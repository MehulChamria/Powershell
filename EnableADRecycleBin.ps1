<#
  .SYNOPSIS
  This script enables Active Directory Recycle Bin on the domain controller.
  .DESCRIPTION
  This script when run on the domain controller enables Active Directory Recycle Bin.
  .INPUTS
  None.
  .EXAMPLE
  PS> .\EnableADRecycleBin.ps1
  .OUTPUTS
  Success or Fail messages on the console and a log file in Logs folder created in script directory. The log file captures all the output from the console.
  .NOTES
  Version:        1.0
  Author:         Mehul Chamria
  Creation Date:  18/07/2022
  Purpose/Change: Initial script development
#>

$null = Start-Transcript -Path $PSScriptRoot\Logs\$($MyInvocation.MyCommand.Name).log -Append

Write-Host "Enable AD Recycle Bin: " -NoNewline
if ((Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"').EnabledScopes) {
    Write-Host "Skipped - Already enabled"
}
else {
    try {
        $null = Enable-ADOptionalFeature 'Recycle Bin Feature' -Scope ForestOrConfigurationSet -Target $env:USERDOMAIN -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "Successful" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed" -ForegroundColor Red
        $null = Read-Host "Press any key to exit"
        $null = Stop-Transcript
		break
    }
}
$null = Read-Host "Press any key to exit"
$null = Stop-Transcript