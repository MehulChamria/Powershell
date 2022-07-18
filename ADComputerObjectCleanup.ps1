<#
  .SYNOPSIS
  This script helps with AD cleanup for computer object.
  .DESCRIPTION
  The script imports the list of computer objects and performs the following operations:
  1. Sets a description on the computer object.
  2. Disables the computer object.
  3. Moves the computer object to Disabled OU.
  .INPUTS
  None.
  .EXAMPLE
  PS> .\ADComputerObjectCleanup.ps1
  .OUTPUTS
  Status messages on the console and a log file in Logs folder created in script directory. The log file captures all the output from the console.
  .NOTES
  Version:        1.0
  Author:         Mehul Chamria
  Creation Date:  18/07/2022
  Purpose/Change: Initial script development
#>

$null = Start-Transcript -Path $PSScriptRoot\Logs\$($MyInvocation.MyCommand.Name).log -Append

$Description = "Computer Disabled on XX/XX/XXXX XX:XX:XX CHGXXXXXXX"
$DisabledOU = "OU=Disabled Computers,DC=Globomantics,DC=Local"

$ServerListFile = "C:\Temp\ServersList.txt"
$Servers = Get-Content $ServerListFile

foreach($Server in $Servers) {
    Write-Host "`nSetting description on computer object '$Server' to - '$Description': " -NoNewline
    try {
        Set-ADComputer $Server -Description $Description -ErrorAction Stop
        Write-Host "Successful" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "Error: $($Error[0].Exception.Message)"
        continue
    }

    Write-Host "Disabling '$Server': " -NoNewline
    try {
        $null = Disable-ADAccount -Identity (Get-ADComputer $Server) -ErrorAction Stop
        Write-Host "Successful" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "Error: $($Error[0].Exception.Message)"
        continue
    }

    Write-Host "Moving '$Server' to Disabled OU: " -NoNewline
    try {
        $null = Move-ADObject -Identity (Get-ADComputer $Server) -TargetPath $DisabledOU -ErrorAction Stop
        Write-Host "Successful" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "Error: $($Error[0].Exception.Message)"
        continue
    }
}

$null = Read-Host "Press any key to exit"
$null = Stop-Transcript