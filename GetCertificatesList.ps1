<#
  .SYNOPSIS
  This script helps pull deatils of all local certificates on a server.

  .DESCRIPTION
  The script imports the list of computer objects from a specific OU in AD and performs the following operations:
  1. Gets all the certificates in the local store for each machine
  2. Saves the information in an Excel file
  3. A use case for this script could be to find all the machines that have expiring local certificates.
#>

If (-not $Credentials){
    $Credentials = Get-Credential
}

$Certificates = @()
$Live_Servers = @()
$Dead_Servers = @()

$Servers = Get-ADComputer -Filter * `
			 -SearchBase "OU=Servers,OU=Computers,DC=globomantics,DC=com" `
			 -Server globomantics.com

$i = 1
Foreach ($Server in $Servers) {
    If ($Server.DNSHostName) {
        $Computer = $Server.DNSHostName
	}
    Else {
        $Computer = $Server.Name
	}
Write-Host "`n$i`. Checking $Computer" -ForegroundColor Gray
$Certificates += `
Try {
    Invoke-Command -ComputerName $Computer -ErrorAction Stop -ArgumentList $Computer -Credential $credentials -ScriptBlock {
    $Server = $args[0]
    $Certificates = Get-ChildItem Cert:\LocalMachine\My |Select Subject,NotBefore,NotAfter,Issuer
    $Outputs = @()
    If ($Certificates){
        Foreach ($Certificate in $Certificates){
            $Output = ""| Select Server,Subject,Expires,Starts,Issue
            $Output.Server = $Server
            $Output.Subject = $Certificate.Subject
            $Output.Starts = $Certificate.NotBefore
            $Output.Expires = $Certificate.NotAfter
            $Output.Issue = $Certificate.Issuer
            $Outputs += $Output
		}
	}
    Else {
        $Output = ""| Select Server,Subject,Expires,Starts,Issue
        $Output.Server = $Server
        $Output.Subject = "No Certificates Found"
        $Output.Starts = "No Certificates Found"
        $Output.Expires = "No Certificates Found"
        $Output.Issue = "No Certificates Found"
        $Outputs += $Output
	}
    $Outputs
	}
    $i++
}
Catch{
    $Outputs = @()
    $Output = ""| Select Server,Subject,Expires,Starts,Issue
    $Output.Server = $Computer
    $Output.Subject = "Cannot connect"
    $Output.Starts = "Cannot connect"
    $Output.Expires = "Cannot connect"
    $Output.Issue = "Cannot connect"
    $Outputs += $Output
    $Outputs
    $i++
}
}

$Parameters = @{
	Path = "C:\temp\$(Get-Date -Format yyyMMdd)_CertificatesList.xlsx"
	WorkSheetname = "CertificatesList"
	BoldTopRow = $true
    	AutoFilter = $true
    	AutoSize = $true
}

$Certificates | Export-Excel @Parameters
