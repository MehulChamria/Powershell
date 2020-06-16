<#
  .SYNOPSIS
  This script helps pull deatils of all local certificates on a server.

  .DESCRIPTION
  The script imports the list of computer objects from a specific OU in AD and performs the following operations:
  1. Gets all the certificates in the local store for each machine
  2. Saves the information in an Excel file
  3. A use case for this script could be to find all the machines that have expiring local certificates.
  4. Supports multiple OUs and multiple domains.
  5. Creates a log file.
#>

If (-not $Credentials){
    $Credentials = Get-Credential
}

$domainTable = [Ordered]@{
    GLOBOMANTICS_UAT = @{
        SearchBase = "OU=Servers,OU=Computers,OU=GLOBOMANTICS_UAT,DC=GLOBOMANTICSPROD,DC=com"
        Server = "GLOBOMANTICSPROD.COM"
    }
    GLOBOMANTICS_PROD = @{
        SearchBase = "OU=Servers,OU=Computers,OU=GLOBOMANTICS,DC=GLOBOMANTICSPROD,DC=com"
        Server = "GLOBOMANTICSPROD.COM"
    }
    GLOBOMANTICS_MGMT = @{
        SearchBase = "OU=Servers,OU=Computers,OU=GLOBOMANTICS,DC=GLOBOMANTICSMGMT,DC=com"
        Server = "GLOBOMANTICSMGMT.COM"
    }
    GLOBOMANTICS_TEST = @{
        SearchBase = "OU=Servers,OU=Computers,OU=GLOBOMANTICS,DC=GLOBOMANTICSTEST,DC=com"
        Server = "GLOBOMANTICSTEST.COM"
    }
    GLOBOMANTICS_DEV = @{
        SearchBase = "OU=Servers,OU=Computers,OU=GLOBOMANTICS_DEV,DC=GLOBOMANTICSTEST,DC=com"
        Server = "GLOBOMANTICSTEST.COM"
    }
}

$logFile = "C:\Temp\List_Certificate.log"
"GetCertificatesList.ps1" > $logFile

$domainTable.Keys | ForEach-Object {
    $Certificates = @()
    $domain = $_
    $GLOBOMANTICSLiveServers = @()
    $GLOBOMANTICSDeadServers = @()
    "[INFO] Initializing scan for certificates on servers in $($domain)." >> $logFile

    $Servers = Get-ADComputer -Filter * `
                    -SearchBase $domainTable[$_]['SearchBase'] `
                    -Server $domainTable[$_]['Server']
    
    "[INFO] Found $($Servers.count) $($domain) servers in Active Directory." >> $logFile
    
    $Servers | ForEach-Object {
        If ($_.DNSHostName) {
            $Computer = $_.DNSHostName
        }
        Else {
            $Computer = $_.Name
        }

        if (Test-Connection -ComputerName $Computer -Quiet -Count 1) {
            "[INFO] $($Computer): Connection test successful." >> $logFile
            $GLOBOMANTICSLiveServers += $Computer
        }
        else {
            "[WARNING] $($Computer): Connection test failed." >> $logFile
            $GLOBOMANTICSDeadServers += $Computer
        }
    }
    
    "[INFO] $($GLOBOMANTICSLiveServers.count) $($domain) servers succeeded connection test." >> $logFile
    "[WARNING] $($GLOBOMANTICSDeadServers.count) $($domain) servers failed connection test." >> $logFile
    $Certificates += `
        Try {
            Invoke-Command -ComputerName $GLOBOMANTICSLiveServers -ErrorAction Stop -Credential $credentials -ScriptBlock {
                $HostName = "$($env:computername).$($env:userdnsdomain)"
                $Certificates = Get-ChildItem Cert:\LocalMachine\My | Select-Object Subject,NotBefore,NotAfter,Issuer
                $Outputs = @()
                If ($Certificates) {
                    Foreach ($Certificate in $Certificates) {
                        $Output = ""| Select-Object HostName,Subject,Expires,Starts,Issue
                        $Output.HostName = $HostName
                        $Output.Subject = $Certificate.Subject
                        $Output.Starts  = $Certificate.NotBefore
                        $Output.Expires = $Certificate.NotAfter
                        $Output.Issue   = $Certificate.Issuer
                        $Outputs        += $Output
                    }
                }
                Else {
                    $Output = ""| Select-Object HostName,Subject,Expires,Starts,Issue
                    $Output.HostName = $HostName
                    $Output.Subject = "No Certificates Found"
                    $Output.Starts  = "No Certificates Found"
                    $Output.Expires = "No Certificates Found"
                    $Output.Issue   = "No Certificates Found"
                    $Outputs        += $Output
                }
                $Outputs
            }
        }
        Catch {
            $Outputs = @()
            $Output = ""| Select-Object HostName,Subject,Expires,Starts,Issue
            $Output.HostName = "Cannot connect"
            $Output.Subject = "Cannot connect"
            $Output.Starts = "Cannot connect"
            $Output.Expires = "Cannot connect"
            $Output.Issue = "Cannot connect"
            $Outputs += $Output
            $Outputs
        }
    
    $Certificates += `
        $GLOBOMANTICSDeadServers | ForEach-Object {
            $Output = ""| Select-Object HostName,Subject,Expires,Starts,Issue
            $Output.HostName = $_
            $Output.Subject = "Ping failed"
            $Output.Starts  = "Ping failed"
            $Output.Expires = "Ping failed"
            $Output.Issue   = "Ping failed"
            $Output
        }
    
    "[INFO] Completed scan for certificates on servers in $($domain)." >> $logFile
    "[INFO] Exporting certificates information for $($domain) to Excel." >> $logFile
    $Parameters = @{
        Path = "C:\temp\$(Get-Date -Format yyyMMdd)_CertificatesList.xlsx"
        WorkSheetname = "$($Domain)"
        BoldTopRow = $true
        AutoFilter = $true
        Verbose = $false
    }
    
    $Certificates | Select-Object HostName,Subject,Starts,Expires,Issue |Export-Excel @Parameters
    "[INFO] Successfully exported certificates information for $($domain) to Excel." >> $logFile
}
