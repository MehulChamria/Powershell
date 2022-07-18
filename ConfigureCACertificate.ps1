<#
  .SYNOPSIS
  This script creates a root CA certificate.
  .DESCRIPTION
  This script creates a root CA certificate.
  .INPUTS
  None.
  .EXAMPLE
  PS> .\ConfigureCACertificate.ps1
  .OUTPUTS
  Success or Fail messages on the console and a log file in Logs folder created in script directory. The log file captures all the output from the console.
  .NOTES
  Version:        1.0
  Author:         Mehul Chamria
  Creation Date:  18/07/2022
  Purpose/Change: Initial script development
#>

$null = Start-Transcript -Path $PSScriptRoot\Logs\$($MyInvocation.MyCommand.Name).log -Append

$DomainName = $env:USERDOMAIN
$HostName = $env:COMPUTERNAME

$Cert = Get-ChildItem -Path Cert:\LocalMachine\My\ | Where-Object Subject -Like "CN=$DomainName-$HostName-CA*"

Write-Host " - Configure CA Certificate: " -NoNewLine
if ($null -eq $cert) {
	$parameters = @{
		ValidityPeriod = "Years"
		ValidityPeriodUnits = 5
		CAType = "EnterpriseRootCA"
		CryptoProviderName = "RSA#Microsoft Software Key Storage Provider"
		HashAlgorithmName = "SHA256"
		KeyLength = 2048
		Confirm = $false
		ErrorAction = "Stop"
	}
	try {
		$null = Install-AdcsCertificationAuthority @parameters
		Write-Host "Successful" -ForegroundColor Green
	}
	catch {
		Write-Host "Failed" -ForegroundColor Red
		$null = Read-Host "Press any key to exit"
		$null = Stop-Transcript
		break
	}
}
else {
	Write-Host "Skipped - Already exists"
}

$null = Read-Host "Press any key to exit"
$null = Stop-Transcript