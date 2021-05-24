<#
  .SYNOPSIS
  This script gets a list of all LUN and CIFS mapped on remote machine/s

  .DESCRIPTION
  The script imports a list of all servers from a text file and retrieves all the information about CIFS and LUN mapped on the server.
  Please change the $Label, $Type and $Storage to match your environment. There are other ways to retrieve disk model information as well but this was a quick one in my case.
#>


function Get-MountPoints {
    [cmdletbinding()]
    param($Computer)

    $Label = @{
        Name = "Name"
        Expression = {
            if ($_.Label -like "A:\*") {
                $_.Label.Trim('A:\')
            }
            else {
                $_.Label
            }
        }
    }
    
    $Type = @{
        Name = "Type"
        Expression = {
            if ($_.Caption -like "B:\*" -or $_.Caption -like "A:\*") {
                "LUN"    
            }
            elseif ($_.Caption -like "L:\*") {
                "CIFS"
            }
        }
    }
    
    $MountPoints = @{
        Name = "MountPoints"
        Expression={$_.Caption}
    }
    
    $Storage = @{
        Name = "Storage"
        Expression = {
            if ($_.Caption -like "B:\*" -or $_.Caption -like "L:\*") {
                "3Par"
            }
            elseif ($_.Caption -like "A:\*") {
                "DataDomain"
            }   
        }
    }
    
    $TotalGB = @{
        Name="Capacity (GB)"
        Expression={"{0:N2}" -f ($_.Capacity/1GB)}
    }
    
    $UsedSpace = @{
        Name="UsedSpace (GB)"
        Expression={"{0:N2}" -f (($_.Capacity/1GB) - ($_.FreeSpace/1GB))}
    }
    
    $FreeGB = @{
        Name="FreeSpace (GB)"
        Expression={"{0:N2}" -f ($_.FreeSpace/1GB)}
    }
    
    $FreePerc = @{
        Name="Free"
        Expression={"{0:P2}" -f (($_.FreeSpace/1GB)/($_.Capacity/1GB))}
    }

    $Volumes = Get-CIMInstance -ComputerName $Computer Win32_Volume | Where-object {$null -eq $_.DriveLetter}
    $Volumes | Select-Object SystemName, $Label, $Type, $MountPoints, $storage, $TotalGB, $UsedSpace, $FreeGB, $FreePerc `
             | Where-Object MountPoints -NotLike "\\?\Volume*"
}

$DateTime = Get-Date -UFormat "%Y%m%d%H%M"
$Servers = Get-Content C:\Temp\Servers.txt
$Path = "C:\Temp\MappedVolume-$DateTime.xlsx"
$Volumes = @()

Foreach($Server in $Servers){
    Write-Host Querying $Server
    $Volumes += Get-MountPoints -Computer $Server
}

$Parameters = @{
    Path        = $Path
    Append      = $true
    AutoSize    = $true
    TableName   = "Table_MappedVolumeList"
    WorkSheetName = "MappedVolumeList"
}
$Volumes | Export-Excel $Parameters
