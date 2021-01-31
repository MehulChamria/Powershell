<#
  .SYNOPSIS
  This script finds the users in ActiveDirectory who do not have the Company or Department field set and exports the results to a database.

  .DESCRIPTION
  1. The script imports a list of all users from ActiveDirectory.
  2. Filters the users who do not have Company or Department field set
  3. Exports the results to a database
#>


$VP = $VerbosePreference
$VerbosePreference = "Continue"

Function Import-SQLServerModule {
    Write-Verbose "Checking for SQLServer module"
    if (Get-Module -ListAvailable -Verbose:$false | Where-Object Name -eq SQLServer){
        Write-Verbose "SQLServer module found on this system."
        Write-Verbose "Attempting to importing SQLServer module."
        try {
            Import-Module SQLServer -Verbose:$false -ErrorAction Stop
        }
        catch {
            Write-Error -Message "Module SQLServer failed to import."
        }
    }
    else {
        Write-Error "SQLServer module was not found on the system."
        exit
    }
}

Function Import-ADModule {
    Write-Verbose "Checking for ActiveDirectory module"
    if (Get-Module -ListAvailable -Verbose:$false | Where-Object Name -eq ActiveDirectory){
        Write-Verbose "ActiveDirectory module found on this system."
        Write-Verbose "Attempting to importing ActiveDirectory module."
        try {
            Import-Module SQLServer -Verbose:$false -ErrorAction Stop
        }
        catch {
            Write-Error -Message "Module ActiveDirectory failed to import."
        }
    }
    else {
        Write-Error "ActiveDirectory module was not found on the system."
        exit
    }
}

Import-SQLServerModule
Import-ADModule

$domain = @{
    Name = "Domain"
    Expression = {
        $_.UserPrincipalName.split("@")[1]
    }
}
$timeStamp =  @{
    Name = "TimeStamp"
    Expression = {Get-Date}
}
$parameters = @{
    Filter = "*"
    Properties = "Department","Company"
}

$usersReport = Get-ADUser @parameters `
                | Where-Object {$null -eq $_.Department -or $null -eq $_.Company} `
                | Select-Object $domain, SamAccountName, $timeStamp

$credentials = Get-Credential -Message "Enter the credentials to connect to the database"

$usersReport | ForEach-Object {
        $insertQuery = " 
            INSERT INTO [dbo].[ServiceTable] 
                ([Domain] 
                ,[LogonID] 
                ,[TimeStamp]) 
            VALUES 
                ('$_.domain' 
                ,'$_.SamAccountName' 
                ,'$_.TimeStamp') 
            GO
    "
    $parameters = @{
        ServerInstance = "ReportingServer01"
        Query = $insertQuery
        Credentials = $credentials
        Database = "ADReports"
    }
    Invoke-SQLcmd @parameters
}

$VerbosePreference = $VP
