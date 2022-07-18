<#
  .SYNOPSIS
  This script helps pull all applications installed on remote host/s.
  .DESCRIPTION
  The script imports the list of computer objects from a csv file and performs the following operations:
  1. Checks if they are rechable.
  2. Pulls a list of all the applications installed on the system in the following format:
  	a. Display Name
	b. Display Version
	c. Publisher Name
	d. Uninstall String
  3. Exports the information to Excel format, if the module is available or CSV format.
  4. A use case for this script could be to find all applications installed on one or more remote hosts.
#>

function Get-InstalledApplications {
    [CmdletBinding()]
    param (
        [array]$Computers
    )

    $credentials = Get-Credential
    $liveHosts = @()
    $deadHosts = @()

    $Computers | ForEach-Object {
        if (Test-Connection $_ -Count 1 -Quiet) {
            $liveHosts += $_
        }
        else {
            $deadHosts += $_
        }
    }

    $Output = 
    Invoke-Command -ComputerName $liveHosts -Credential $credentials -ScriptBlock {
        $Apps = @()
        $32BitPath = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $64BitPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        $Apps += Get-ItemProperty "HKLM:\$32BitPath"
        $Apps += Get-ItemProperty "HKLM:\$64BitPath"
        $Apps |
        Select-Object DisplayName , DisplayVersion, Publisher, UninstallString
    }

    Write-Warning "The following hosts were not reachable:"
    $deadHosts

    if ((Get-Module -ListAvailable).Name -contains "ImportExcel") {
        $Parameters = @{
            Path = "C:\temp\$(Get-Date -Format yyyMMdd)_ApplicationsList.xlsx"
            WorkSheetname = "AppsList"
            TableName = "Table_AppsList"
            AutoFilter = $true
            AutoSize = $true
        }
        
        $Output |
            Select-Object PSComputerName, DisplayName, DisplayVersion |
            Export-Excel @Parameters
    }
    else {
        $Output |
            Select-Object PSComputerName, DisplayName, DisplayVersion |
            Export-Csv -Path "C:\temp\$(Get-Date -Format yyyMMdd)_ApplicationsList.csv" -NoTypeInformation
    }
}

$Computers = Import-Csv C:\Temp\Hosts.csv
Get-InstalledApplications -Computers $Computers
