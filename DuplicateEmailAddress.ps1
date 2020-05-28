<#
  .SYNOPSIS
  This script queries a list of all the accounts in AD which uses same email address. This will only work for accounts where a user object is not linked with Exchange and has email address field specified manually. Exchange enabled accounts will have a unique email address and this script might not be for that use case.

  .DESCRIPTION
  The script imports the list of all user objects in a specified domain and performs the following operations:
  1. Filters the list and only includes the accounts that are enabled and has an email address field specified
  2. Groups the output from the previous step by Email Address
  3. Filters the output from the previous step and only extracts those groups that have more than 1 account (Meaning an email address is associated to multiple accounts)
  4. It then exports the output from the previous step to an Excel file
  5. USE CASE: There may be a requirement to find if the same email address has been used for more than one account
#>

$Parameters = @{
	Filter = "*"
	Properties = "EmailAddress"
	Server = "globomantics.com"
}
$DuplicateEmailAccounts = Get-ADUser @Parameters `
   |Where-Object {$_.Enabled -eq $true -and $_.EmailAddress -ne $null} `
   |Group-Object EmailAddress `
   |Where-Object {$_.Count -ge 2}

$Parameters = @{
	Path = "C:\temp\$(Get-Date -Format yyyMMdd)_DuplicateEmailAccounts.xlsx"
	WorkSheetname = "DuplicateEmailAccounts"
	BoldTopRow = $true
  AutoFilter = $true
  AutoSize = $true
}
$DuplicateEmailAccounts| select -ExpandProperty Group `
   | Select Name, EmailAddress, SamAccountName, DistinguishedName `
   | Export-Excel @Parameters
