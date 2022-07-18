<#
  .SYNOPSIS
  This script disables IE Enhanced Mode on Windows Server.
  .DESCRIPTION
  This script disables IE Enhanced Mode. Open the powershell session in Admin mode as this script makes changes to the registry to disable IE Enhanced Mode. The script will fail if the powershell session is not run as an admin.
  .INPUTS
  None.
  .EXAMPLE
  PS> .\DisableIEEnhancedMode.ps1
  .OUTPUTS
  Success or Fail messages on the console and a log file in Logs folder created in script directory. The log file captures all the output from the console.
  .NOTES
  Version:        1.0
  Author:         Mehul Chamria
  Creation Date:  18/07/2022
  Purpose/Change: Initial script development
#>

$null = Start-Transcript -Path $PSScriptRoot\Logs\$($MyInvocation.MyCommand.Name).log -Append

$AdminKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
$UserKey = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
$flag = 0

Write-Host " - Disable IE Enhanced Security Configuration (ESC): " -NoNewLine

if ((Get-ItemProperty -Path $AdminKey -Name "IsInstalled").IsInstalled -eq 1) {
    try {
        $null = Set-ItemProperty -Path $AdminKey -Name "IsInstalled" -Value 0 -ErrorAction Stop
        $flag = 1
    }
    catch {
        Write-Host "Failed - Unable to set Admin key" -ForegroundColor Red
        $null = Read-Host "Press any key to exit"
        $null = Stop-Transcript
        return
    }
}

if ((Get-ItemProperty -Path $UserKey -Name "IsInstalled").IsInstalled -eq 1) {
    try {
        $null = Set-ItemProperty -Path $UserKey -Name "IsInstalled" -Value 0 -ErrorAction Stop
        $flag = 1
    }
    catch {
        Write-Host "Failed - Unable to set User key" -ForegroundColor Red
        $null = Read-Host "Press any key to exit"
        $null = Stop-Transcript
        return
    }
}

if ($flag -eq 1) {
    try {
        $null = Stop-Process -Name Explorer -ErrorAction Stop
        Write-Host "Successful" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed - Unable to stop process Explorer.exe" -ForegroundColor Red
        $null = Read-Host "Press any key to exit"
        $null = Stop-Transcript
        return
    }
}
else {
    Write-Host "Skipped - Already Disabled"
}

$null = Read-Host "Press any key to exit"
$null = Stop-Transcript