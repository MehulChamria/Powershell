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
            BoldTopRow = $true
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
            Export-Csv -Path "C:\temp\$(Get-Date -Format yyyMMdd)_ApplicationsList.xlsx" -NoTypeInformation
    }
}

$Computers = Import-Csv C:\Temp\Hosts.csv
Get-InstalledApplications -Computers $Computers
