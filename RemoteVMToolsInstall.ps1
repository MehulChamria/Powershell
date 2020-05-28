<#
  .SYNOPSIS
  This script remotely installs/upgrades VMware tools on Windows guest VM.

  .DESCRIPTION
  The script imports a list of all servers from a text file and performs the following operation:
  1. Connects to the vCenter
  2. Takes snapshot of the guest VM
  3. Copies the VM Tools installer files to the guest VM
  4. Executes the installer using Invoke-Command and specifies switches to install silently
  5. Deletes the VM Tools installer files from the guest VM
  6. Displays the version of VM Tools installed. This will help you verify the VM Tools version after installation.
#>

Function Connect-VMWare{
    param(
    [Parameter(Mandatory=$True)]
    [String]$vCenter,
	
	[ValidateNotNull()]
	[System.Management.Automation.PSCredential]
	[System.Management.Automation.Credential()]
	$Credential = [System.Management.Automation.PSCredential]::Empty
	)

    If (!(Get-Module -Name "VMware.VimAutomation.Core")){
        Try {
            Write-Host -ForegroundColor DarkGray "Loading VMware PowerCLI Module..."
            Import-Module -Name VMware.VimAutomation.Core -ea 0| Out-Null
            Write-Host "`nModule Loaded" -ForegroundColor Green
		}
        Catch {
            Write-Host "`nCould not load PowerCLI module" -ForegroundColor Red
		}
	}
    Else { 
        Write-Host "`nPowerCLI Module Loaded" -ForegroundColor Green
	}
    Write-Host "`nConnecting to $($vCenter) vCenter server..." -ForegroundColor DarkGray
    $Connection_Info = Connect-VIServer $vCenter -Credential $Credentials

    If ($Connection_Info.IsConnected -eq "True") {
        Write-Host "Connected to $($vCenter) vCenter server" -ForegroundColor Green
	}
    Else {
        Write-Host "Failed to connect to VI server" -ForegroundColor Red
	}
}

$vCenter = "vCenter.globomantics.com"
$ESX_Cluster = "GLOBO-PROD-01"

$Servers = Get-Content C:\Temp\Servers.txt

if (!($vCenterCredentials)) {
    $vCenterCredentials = Get-Credential -Message "Enter your credentials for connecting to vcenter."
}

Connect-VMWare -vCenter $vCenter -Credential $vCenterCredentials

if (!($ServerCredentials)) {
    $ServerCredentials = Get-Credential -Message "Enter your credentials for connecting to servers to upgrade VM Tools."
}

$VMToolsLocation = "\\Globomantics.com\VMware\vmware tools\VMware-Tools-core-10.3.2-9925305\vmtools\windows\"
$TempLocation = "\\$($Server)\C$\Temp\Windows"

foreach($Server in $Servers) {
    if(Get-VM -Name $Server) {
        New-Snapshot -VM $Server -Description "Snapshot of the VM before upgrading VMTools." `
                     -Name "Pre VM Tools Upgrade"
        Copy-Item $VMToolsLocation $TempLocation -recurse
        Invoke-Command -ComputerName $Server -Credential $ServerCredentials -ScriptBlock {
        cmd /c c:\temp\windows\setup64.exe /S /v /qn
        }
        Remove-Item $TempLocation -Recurse
        Write-Host "`nThe version of tools installed on $Server is: $((get-vm $Server|Get-VMGuest).Toolsversion)"
	}
    Else {
        Write-Host "`nVM $Server not found!`n" -ForegroundColor Red
	}
}
