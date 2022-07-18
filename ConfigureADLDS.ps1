<#
  .SYNOPSIS
  This script installs Active Directory Lightweight Directory Services (AD LDS) on the domain controller.
  .DESCRIPTION
  This script when run on the domain controller installs Active Directory Lightweight Directory Services (AD LDS). Change the port number configuration as required in the script.
  .INPUTS
  None.
  .EXAMPLE
  PS> .\ConfigureADLDS.ps1
  .OUTPUTS
  Success or Fail messages on the console and a log file in Logs folder created in script directory. The log file captures all the output from the console.
  .NOTES
  Version:        1.0
  Author:         Mehul Chamria
  Creation Date:  18/07/2022
  Purpose/Change: Initial script development
#>

$null = Start-Transcript -Path $PSScriptRoot\Logs\$($MyInvocation.MyCommand.Name).log -Append

$HostName = $env:COMPUTERNAME
$Client = $env:USERDOMAIN
$LocalLDAPPortToListenOn = 10000
$LocalSSLPortToListenOn = 10001
$AnswerFile = "$PSScriptRoot\Files\ADLDSAnswerFile.txt"
$WindowsFeatures = "AD-Certificate", "RSAT-ADCS-Mgmt", "RSAT-ADCS", "ADLDS"
$ADLDSInstanceName = "LDS-$HostName-$Client"
$ADLDSKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ADAM_$ADLDSInstanceName*"

if ($null -ne (Get-ItemProperty $ADLDSKey)) {
    Write-Host "Configure AD Lightweight Directory Services: Skipped - Instance already exists"
    $null = Read-Host "Press any key to exit"
    $null = Stop-Transcript
    return
}

Write-Host "Configure AD Lightweight Directory Services"
Write-Host " - Installing Windows Features"
$WindowsFeatures | ForEach-Object {
    Write-Host "      - $($_): " -NoNewline
    try {
        $null = Install-Windowsfeature -Name $_ -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Host "Successful" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed" -ForegroundColor Red
        $null = Read-Host "Press any key to exit"
        $null = Stop-Transcript
        return
    }
}

Write-Host " - Configuring AD LDS: " -NoNewline
if (($null -eq (Get-NetTCPConnection -LocalPort $LocalLDAPPortToListenOn -ErrorAction SilentlyContinue)) -and
    ($null -eq (Get-NetTCPConnection -LocalPort $LocalSSLPortToListenOn -ErrorAction SilentlyContinue))) {
        Set-Content -Path $AnswerFile -Value "[ADAMInstall]`nInstallType=Unique`nInstanceName=$ADLDSInstanceName"
        Add-Content -Path $AnswerFile -Value "LocalLDAPPortToListenOn=$LocalLDAPPortToListenOn`nLocalSSLPortToListenOn=$LocalSSLPortToListenOn"
        Add-Content -Path $AnswerFile -Value "DataFilesPath='C:\Program Files\Microsoft ADAM\LDS_$HostName_$Client\Data'"
        Add-Content -Path $AnswerFile -Value "LogFilesPath='C:\Program Files\Microsoft ADAM\LDS_$HostName_$Client\Data'"
        $null = Start-Process -FilePath "C:\Windows\ADAM\adaminstall.exe" -ArgumentList "/answer:$AnswerFile" -Wait
        if (Get-ItemProperty $ADLDSKey) {
            Write-Host "Successful" -ForegroundColor Green
        }
        else {
            Write-Host "Failed - Reboot the server and rerun the script" -ForegroundColor Red
            $null = Read-Host "Press any key to exit"
            $null = Stop-Transcript
            break
        }
}
else {
    Write-Host "Failed - The ports $LocalLDAPPortToListenOn and/or $LocalSSLPortToListenOn are currently in use. Please reboot the host and rerun the script." -ForegroundColor Red
    $null = Read-Host "Press any key to exit"
    $null = Stop-Transcript
    break
}

$null = Read-Host "Press any key to exit"
$null = Stop-Transcript