<#
  .SYNOPSIS
  This script finds the lock status of BitLockered drives on remote machine/s

  .DESCRIPTION
  The script imports a list of all servers from a text file and retrieves lock status of BitLocker drives on remote machine/s.
#>


$Servers = Get-Content C:\temp\NUIXServers.txt
$Report = @()

$Servers | ForEach-Object {
    $Server = $_
    $Drives = (Get-CimInstance win32_logicaldisk -ComputerName $Server).DeviceID
    $Drives | ForEach-Object {
        $Drive = $_
        $BDStatus = Manage-Bde -ComputerName $Server -Status $Drive
        if($BDStatus -like "*Protection Status:    Protection On*"){
            if ($BDStatus -like "*Lock Status:          Unlocked*") {
				$Status = "Unlocked"
            }
            elseif ($BDStatus -like "*Lock Status:          locked*") {
                $Status = "Locked"
            }
            else {
				$Status = "Error"
            }
        $Output = [Ordered]@{
            Server = $Server
            Drive = $Drive
            Status = $Status
        }
        $Result = [PScustomObject]$Output
        $Result
        $Report += $Result
        }
    }
}
