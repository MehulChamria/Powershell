<#
  .SYNOPSIS
  This script helps with AD cleanup for computer object.

  .DESCRIPTION
  The script imports the list of computer objects and performs the following operations:
  1. Sets a description on the computer object.
  2. Disables the computer object.
  3. Moves the computer object to Disabled OU.
#>

$V_preference = $VerbosePreference
$VerbosePreference = "Continue"
$Description = "Computer Disabled on XX/XX/XXXX XX:XX:XX CHGXXXXXXX"
$DisabledOU = "OU=Disabled Computers,DC=Globomantics,DC=Local"

$Servers = Get-Content C:\Temp\ServersList.txt

foreach($Server in $Servers) {
    
    Write-Verbose  "Setting description on computer object '$Server' to - '$Description'."
    try {
        Set-ADComputer $server -Description $Description -ErrorAction Stop
        Write-Host Description set successfully.`n -ForegroundColor Green
    }
    catch {
        Write-Output "Failed to set the description on the computer object $Server.`n"
        Write-Output "Error: $($Error[0].Exception.Message)`n"
    }

    Write-Verbose "Disabling '$Server'"
    try {
        Disable-ADAccount -Identity (Get-ADComputer $Server) -ErrorAction Stop
        Write-Host Computer account $server is now disabled.`n -ForegroundColor Green
    }
    catch {
        Write-Output "Failed to disable computer object $Server."
        Write-Output "Error: $($Error[0].Exception.Message)`n"
    }

    Write-Verbose "Moving '$Server' to Disabled OU."
    try {
        Move-ADObject -Identity (Get-ADComputer $Server) -TargetPath $DisabledOU -ErrorAction Stop
        Write-Host Computer account $server is now moved to disabled OU.`n -ForegroundColor Green
    }
    catch {
        Write-Output "Failed to move computer object $Server to Disabled OU."
        Write-Output "Error: $($Error[0].Exception.Message)`n"
    }

    Read-Host "Press any key to continue"
}

$VerbosePreference = $V_preference
