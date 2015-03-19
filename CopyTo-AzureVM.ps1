#CopyTo-AzureVM.ps1

Set-StrictMode -Version latest
Set-PSDebug -Strict


################################################################################################

#region support funcs
function New-UserPromptChoice 
{
    PARAM($options)

    $arrOptions = @()
    $i = 1

    foreach ($option in $options)
    {
        $optionDesc = New-Object  -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList ("&$i - $option")
        $arrOptions += $optionDesc
        $i++
    }

    return [System.Management.Automation.Host.ChoiceDescription[]]($arrOptions)
}

function Select-MyAzureSubscription 
{
    [CmdletBinding()]
    Param()

    Write-Host 'Selecting the Azure subscription.'
      
    if (-not (Get-AzureAccount))
    {
        Add-AzureAccount
    }    

    $subscriptions = Get-AzureSubscription | Select-Object -ExpandProperty SubscriptionName 

    if ($subscriptions -is [System.Array])
    {
        $subscriptionPrompt = New-UserPromptChoice -options $subscriptions
        $subscriptionChoice = $Host.UI.PromptForChoice('Subscription Name','Please choose a subscription',$subscriptionPrompt,0)   

        $subscription = $subscriptions[$subscriptionChoice]
        Write-Verbose  -Message "Setting Azure Subscription to: $subscription"
        Select-AzureSubscription -SubscriptionName $subscription
    }
    else
    {
      $subscription = $subscriptions
      Write-Verbose  -Message "Setting Azure Subscription to: $subscription"
      Select-AzureSubscription -SubscriptionName $subscription
    }

}

function Check-AzurePowerShellModule 
{
    Param([Parameter(Mandatory)] $minVer)

    Write-Host  -Object 'Checking if the Azure PowerShell module is installed...'
 
    $minVersion = $minVer -Split "\."
    if ($minVersion.Length -lt 3) {Write-Error 'invalid Minimum Version string.'}

    $minMajor=$minversion[0]
    $minMinor=$minVersion[1]
    $minBuild=$minVersion[2]
    # not testing for build revision, so ignore
    if ($minVersion.Length -gt 3) {
        $minRevision=$minVersion[3]
    }

    if (Get-Module -ListAvailable  -Name 'Azure') 
    {
        Write-Host  -Object 'Loading Azure module...'
        Import-Module  -Name 'Azure' -Force

        $ModVer = (Get-Module -Name "Azure").Version
        
        Write-Verbose  -Message "Version installed: $modver  Minimum required: $minVer"

        $minimumBuild = ($ModVer.Major -gt $minMajor) `
        -OR (($ModVer.Major -eq $minMajor) -AND ($ModVer.Minor -gt $minMinor)) `
        -OR (($ModVer.Major -eq $minMajor) -AND ($ModVer.Minor -eq $minMinor) -AND ($ModVer.Build -ge $minBuild))


        if ($minimumBuild) 
        {
            return $true
        }
        else 
        {
            Write-Host "The Azure PowerShell module is NOT a current build. You will now be directed to the download location. `n" -ForegroundColor Yellow
            Write-Host "You will need to install the update, exit and restart PowerShell"
            Start-Sleep 10
            Start-Process -FilePath 'http://go.microsoft.com/fwlink/p/?linkid=320376&clcid=0x409'
            return $false
        }
    }
    else 
    {
        Write-Host  -Object "The Azure PowerShell module is NOT installed you will now be directed to the download location. `n" -ForegroundColor Yellow
        Write-Host "You will need to install the update, exit and restart PowerShell"
        Start-Sleep 10
        Start-Process -FilePath 'http://go.microsoft.com/fwlink/p/?linkid=320376&clcid=0x409'
        return $false
    }
}


function Test-IsAdmin () {
  $me=[Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
  if (-not ( $me.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) ) ) { return $false }
  else { return $true  }
}


function Install-WinRMCert($ServiceName, $Name)
{
  $winRMCertThumbprint = (Get-AzureVM -ServiceName $serviceName -Name $name | select -ExpandProperty vm).DefaultWinRMCertificateThumbprint
  $AzureX509cert = Get-AzureCertificate -ServiceName $serviceName -Thumbprint $winRMCertThumbprint -ThumbprintAlgorithm sha1
 
  $TempFile = [System.IO.Path]::GetTempFileName()
  $AzureX509cert.Data | Out-File $TempFile
 
  $CertToImport = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $TempFile
 
  try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store "Root", "LocalMachine"
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    $FindByThumbprint = [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint
    $exists = $store.Certificates.Find($FindByThumbprint,$winRMCertThumbprint,$true)

    if (-NOT $exists) {
     if (Test-IsAdmin) {
       Write-Verbose 'Adding certificate.'
       $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
       $store.Add($CertToImport)
     }
     else{
       if ($store) { $store.Close() }
       Write-Error 'You do not have the required Administrator rights to add a Certificate. Run PowerShell as Administrator.'
     }
    }

    $store.Close()
  }
  catch {
    if ($store) { $store.Close() }
    Write-Error "Error when querying for certificate in Certificate store 'root', 'LocalMachine'."
  }
  
  if (Test-Path -Path $TempFile) {  Remove-Item -Path $TempFile }
}

#endregion 

function Send-FileToAzure
{
    param (
        [Parameter(Mandatory = $true)] [string] $Source,
        [Parameter(Mandatory = $true)] [string] $Destination,
        [Parameter(Mandatory = $true)] [System.Management.Automation.Runspaces.PSSession] $Session,
        [bool] $onlyCopyNew = $false
    )


#region remotescript

    # this will be Invoked repeatedly to pass across chunks of the file
    $remoteScript =
    {
        param ($destination, $bytes)

        # Convert the destination path to a full filesystem path (to supportrelative paths)
        $Destination = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Destination)

        # append the content to the file
        $file = [System.IO.File]::Open($Destination, "OpenOrCreate")
        $null = $file.Seek(0, [System.IO.SeekOrigin]::End)
        $null = $file.Write($bytes, 0, $bytes.Length)
        $file.Close()
    }
#endregion


    # Get the source file, and then start reading its content
    $sourceFile = Get-Item -Path $Source

#region remove existing file
    # Delete the previously-existing file if it exists
    $abort = Invoke-Command -Session $Session {
        param ([String] $dest, [bool]$onlyCopyNew)

        if (Test-Path $dest) { 
          if ($onlyCopyNew -eq $true) { return $true }
          Remove-Item $dest
        }

        $destinationDirectory = Split-Path -Path $dest -Parent
        if (-NOT (Test-Path $destinationDirectory))  { New-Item -ItemType Directory -Force -Path $destinationDirectory }

        return $false
    } -ArgumentList $Destination, $onlyCopyNew


    if ($abort -eq $true)
    {
        Write-Output 'Ignored file transfer - already exists'
        return
    }
#endregion

#region stream the data

    Write-Progress -Activity "Sending $Source" -Status "Preparing file"
    New-Variable -Name MAXBUFSIZE -Value 1Mb -Option Constant
    $position = 0
    $rawBytes = New-Object byte[] $MAXBUFSIZE

    $file = [System.IO.File]::OpenRead($sourceFile.FullName)

    while (($bytesRead = $file.Read($rawBytes, 0, $MAXBUFSIZE)) -gt 0) {
        Write-Progress -Activity "Writing $Destination" -Status "Sending file" -PercentComplete ($position / $sourceFile.Length * 100)

        # Ensure that our array is shrunk to the same size as what we actually read from disk
        if ($bytesRead -ne $rawBytes.Length) { [Array]::Resize( [ref] $rawBytes, $bytesRead)   }

        # And send that array to the remote system
        Invoke-Command -Session $session -Scriptblock $remoteScript -ArgumentList $destination, $rawBytes

        # Ensure that our array is reset to the maximum size we may read from disk
        if ($rawBytes.Length -ne $MAXBUFSIZE) { [Array]::Resize( [ref] $rawBytes, $MAXBUFSIZE)  }

        [GC]::Collect()
        $position += $bytesRead
    }

    $file.Close()
#endregion

    # Show the copied file now on the target
    Invoke-Command -Session $session -Scriptblock { Get-Item $args[0] } -ArgumentList $Destination
}


#region Windows Forms

Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms 
$PSicon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command -Name powershell).Path)

function Invoke-FileBrowser
 {
  param([string]$Title, [string]$Directory, [string]$Filter='All Files (*.*)|*.*')

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.InitialDirectory = $Directory
    $FileBrowser.Filter = $Filter
    $FileBrowser.Title = $Title
    $Show = $FileBrowser.ShowDialog()
    if ($Show -eq 'OK') { return $FileBrowser.FileName }
}

function Select-ArrayItem  
{  
    param ([Parameter(Mandatory=$true)] $options,  
           [Parameter(Mandatory=$true)] $displayProperty  
    )
       
    function processOK  
    {  
        if ($ListItems.SelectedIndex -lt 0)  
        {  
            $global:selectedItem = $null  
        }  
        else  
        {  
            $global:selectedItem = $options[$ListItems.SelectedIndex]  
        }  
        $form.Close()  
    }  
    $global:selectedItem = $null  
      
    $form = new-object System.Windows.Forms.Form  
    $form.Icon = $PSicon
    $form.Size = new-object System.Drawing.Size @(300,250)     
    $form.text = "Target VM"    

    $ListItems = New-Object System.Windows.Forms.ListBox  
    $ListItems.Name = "ListItems"  
    $ListItems.Width = 200  
    $ListItems.Height = 175  
    $ListItems.Location = New-Object System.Drawing.Size(5,5)  
    $ListItems.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right  

    $ButtonOK = New-Object System.Windows.Forms.Button   
    $ButtonOK.Width=100  
    $ButtonOK.Location = New-Object System.Drawing.Size(50, 180)  
    $ButtonOK.Text = "OK"  
    $ButtonOK.add_click({processOK})  
    $ButtonOK.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right

    $form.Controls.Add($ListItems)  
    $form.Controls.Add($ButtonOK)  
      
    foreach ($option in $options)  
    {  
        [void]$ListItems.Items.Add($option.$displayProperty)
    }  
    $ListItems.SetSelected(0,$true)

    [void] $form.BringToFront()
    [void] $form.ShowDialog()

    return $global:selectedItem  
}  
#endregion

###################################################################################################

if (Check-AzurePowerShellModule "0.8.15") {

  Select-MyAzureSubscription 

  Write-Output 'Retrieving the list of running VMs...'
  $VMlist = Get-AzureVM | where {$_.Status -eq "ReadyRole"}

  if ($null -ne $VMlist) {
   $TARGET = (Select-ArrayItem $VMlist Name)
   if ($null -ne $TARGET) {
    Write-Output "Retrieving the WinRMuri of '$($TARGET.Name)'"
    $uri = Get-AzureWinRMUri -ServiceName $TARGET.ServiceName -Name $TARGET.Name
    if (-NOT ($null -eq $uri)) {

      # Needs the cert installed locally to allow a connect PSsession cleanly. 
      Write-Output "Checking if a Certificate from the server is available"
      Install-WinRMcert -ServiceName $TARGET.ServiceName -Name $TARGET.Name

#region OpenSession
      Write-Output "Opening PSsession to '$($TARGET.Name)'..."
      $ValidSession = $false
      do {
        $cred = Get-Credential -Message "Enter the domain\user and their password, to allow the session to login to the target VM '$($TARGET.Name)'."

        $AzureSession=New-PSsession -ConnectionUri $uri.AbsoluteUri -Credential $cred
        if ($null -ne $AzureSession) {
           $validSession = $true
           Write-Output "Connected!"
        }
      }
      until ($ValidSession)
#endregion

#region doThecopy
      $source = Invoke-FileBrowser -Title 'Local file to copy'
      Write-Output "From Local : $source"
      $destination = Read-Host 'Enter destination filepath'
      if (-NOT (Test-Path -IsValid $destination)) {write-error 'Invalid filepath syntax.'}
      Write-Output "Copying to '$($TARGET.Name)':   $destination"
      Send-FileToAzure -Source $source -Destination $destination -Session $AzureSession
#endregion

      Remove-PSSession -Session $AzureSession

    }
    else {
      write-error 'No uri was returned. The server may be in a shutdown state.'
    }
   }
   else {
    Write-Error 'The Azure hosted target machine must be selected.'
   }
  }
 else {
   write-output "No VMs are in a 'ReadyRole' state."
 }
}