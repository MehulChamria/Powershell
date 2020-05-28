<#
  .SYNOPSIS
  Retrieves the last time windows updates were installed on a remote host.

  .DESCRIPTION
  Retrieves the last time windows updates were installed on a remote host.
#>

if (!$Credentials){
    $Credentials = Get-Credential
}
$HostName = Read-Host "Enter the host name"
Invoke-Command -ComputerName $HostName -Credential $credentials -ScriptBlock {
Write-Host "The windows updates on $env:COMPUTERNAME was last installed on:" ((New-Object -ComObject Microsoft.Update.AutoUpdate).Results).LastInstallationSuccessDate
}
