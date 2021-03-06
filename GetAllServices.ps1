<#
  .SYNOPSIS
  This script helps pull all services running on remote server.

  .DESCRIPTION
  The script imports the list of computer objects from a text file and performs the following operations:
  1. Initiates a CIM or WMI connection
  2. Pulls a list of all the services running on the system along with the following information for each service:
  	a. Server Name
	b. Service Name
	c. Service Display Name
	d. Account Used
	e. Service State
	f. Service Description
  3. A use case for this script could be to find what account each service is using to run on a machine.
#>

$Service_Info = @()
#Below ServiceFilter variable can be used if you wish to filter services instead of listing all services on a machine
#$ServiceFilter = "(NOT StartName LIKE '%LocalSystem') AND (NOT StartName LIKE '%LocalService') AND (NOT StartName LIKE '%NetworkService') AND (NOT StartName LIKE 'NT AUTHORITY%')"
$Servers = Get-Content C:\Servers.txt

Write-Host $Servers.Count "servers found."

$i = $Servers.Count
Foreach ($Server in $Servers){
    Write-Host "`n$i`. Checking $Server" -ForegroundColor Gray
    Try {
        #Test if the server is running PS lower than v3 because the below CIm Instance command wont work on PS2 and lower
        $data = test-wsman $Server -ErrorAction Stop
        #if it is lower than 3 then using WMI as CIM wont be supported
        if(!($rx.match($data.ProductVersion).value -eq '3.0')){
            $Services = Get-WmiObject -Class Win32_Service `
                                      -ComputerName $Server `
                                      -ErrorAction Stop
		}
        else {
        $Services = `
        Get-CimInstance -ClassName Win32_Service `
                        -ComputerName $Server `
                        -ErrorAction Stop
		}

        If (($Service_Info.'Server Name'|Get-Unique) -notcontains ($Services.SystemName|get-Unique)){
            Foreach ($Service in $Services){
                $Service_Detail = New-Object PSObject
                $Service_Detail | Add-Member -MemberType NoteProperty -Name "Server Name" -Value $Service.SystemName
                $Service_Detail | Add-Member -MemberType NoteProperty -Name "Service Name" -Value $Service.Name
                $Service_Detail | Add-Member -MemberType NoteProperty -Name "Service Display Name" -Value $Service.Caption
                $Service_Detail | Add-Member -MemberType NoteProperty -Name "Account Used" -Value $Service.StartName
                $Service_Detail | Add-Member -MemberType NoteProperty -Name "Service State" -Value $Service.State
                $Service_Detail | Add-Member -MemberType NoteProperty -Name "Description" -Value $Service.Description

                $Service_Info += $Service_Detail
			}
		}
        Else {
            Write-Host "`n$server has been scanned already." -ForegroundColor Yellow
        }
	}
    Catch {
        $Service_Detail = New-Object PSObject
        $Service_Detail | Add-Member -MemberType NoteProperty -Name "Server Name" -Value $server
        $Service_Detail | Add-Member -MemberType NoteProperty -Name "Service Name" -Value "Failed to connect" 
        $Service_Detail | Add-Member -MemberType NoteProperty -Name "Service Display Name" -Value "Failed to connect"
        $Service_Detail | Add-Member -MemberType NoteProperty -Name "Account Used" -Value "Failed to connect"
        $Service_Detail | Add-Member -MemberType NoteProperty -Name "Service State" -Value "Failed to connect"
        $Service_Detail | Add-Member -MemberType NoteProperty -Name "Description" -Value "Failed to connect"

        $Service_Info += $Service_Detail
	}
    $i--
}

$Parameters = @{
	Path = "C:\temp\$(Get-Date -Format yyyMMdd)_ServiceAccounts.xlsx"
	WorkSheetname = "ServiceAccountsInfo"
    	AutoFilter = $true
    	AutoSize = $true
}

$Service_Info | Export-Excel @Parameters
