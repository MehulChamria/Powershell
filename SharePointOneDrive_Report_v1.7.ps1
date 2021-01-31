#region Define variables
#SharePoint Admin URL
$sharepoint_admin_url =  "https://indianassociationoldham-admin.sharepoint.com/"

#Author file location
$author_list_file = "C:\Temp\author_list.txt"

#ListName for eg: Documents. To find out list use the following command and use the Title from output as list name: Get-PNPList
$list_name= "Documents"

#Script output location. Do not put \ at the end of the folder location. Example format: C:\Temp\SharePoint
$sharepoint_output_folder = "$env:USERPROFILE\Desktop\SharePointReport"

#CSV output location
$report_file = "$sharepoint_output_folder\$(Get-Date -Format yyyMMdd_hhmm)_SPReport.csv"

#Log file location
$report_log_file = "$sharepoint_output_folder\$(Get-Date -Format yyyMMdd_hhmm)_SPReport_log.txt"
$downloads_log_file = "$sharepoint_output_folder\$(Get-Date -Format yyyMMdd_hhmm)_SPDownloads_log.txt"
$summary_log_file = "$sharepoint_output_folder\$(Get-Date -Format yyyMMdd_hhmm)_SPSummary_log.txt"
$email_log_file = "$sharepoint_output_folder\$(Get-Date -Format yyyMMdd_hhmm)_SPEmail_log.csv"
#Set page size
$page_size = 5000

#Date range in format YYYY/MM/DD HH:MM:SS
[datetime]$start_date = "2021/01/10 02:25:00"
[datetime]$end_date = "2021/01/23 00:25:00"

#Email recipients
$email_recipients = "Mehul.Chamria@kpmg.co.uk", "Dharmesh.Trivedi@kpmg.co.uk"
$email_sender = "no-reply@kpmgmgmt.com"

#Mail server
$smtp_server = "10.174.100.130"

#Email update frequency in minutes
$email_frequency = 30
#endregion

#region Import PS Modules for SharePoint
try {
    Import-Module -Name Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Verbose "Module 'Microsoft.Online.SharePoint.PowerShell' imported successfully." -Verbose
}
catch {
    Write-Warning "Failed to import module 'Microsoft.Online.SharePoint.PowerShell'"
    break
}

try {
    Import-Module -Name SharePointPnPPowerShellOnline -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Verbose "Module 'SharePointPnPPowerShellOnline' imported successfully." -Verbose
}
catch {
    Write-Warning "Failed to import module 'SharePointPnPPowerShellOnline'"
    break
}
#endregion

#region Initialize variables
#Verify if the output folder exists, if not, create it.
if(-not (Test-Path $sharepoint_output_folder)) {
    try{
        New-Item -ItemType Directory -Path $sharepoint_output_folder -ErrorAction Stop | Out-Null
        New-Item -ItemType Directory -Path $sharepoint_output_folder\Downloads -ErrorAction Stop | Out-Null
    }
    catch{
        Write-Warning "Could not create report folder."
        break
    }
}

#Verify if the Author list file exists, if not, exit execution.
if (-not (Test-Path $author_list_file)){
    Write-Warning "Author list file not found."
    break
}
elseif ($null -eq (Get-Content $author_list_file)){
    Write-Warning "Author list file contains no records."
    break
}

#Initialize Output and log files 
Set-Content $report_file -Value $null
Set-Content $report_log_file -Value $null
Set-Content $downloads_log_file -Value $null
Set-Content $summary_log_file -Value $null

$global:file_list = New-Object -TypeName "System.Collections.ArrayList"
$global:file_count = 0
#endregion Initialize variables

#region Connect to SharePoint Admin site
Connect-SPOService -Url $sharepoint_admin_url
$site_list = Get-SPOSite -IncludePersonalSite $true -Limit all |
                Where-Object {($_.URL -Like "*-my.sharepoint.com/personal/*") -or ($_.URL -Like "*/sites/*")}
#endregion

#region Function definition: Get-SharePointReport 
function Get-SharePointReport {
    [cmdletbinding()]
    param()

    #Initialize variables
    $author_list = Get-Content $author_list_file
    $metadata = New-Object -TypeName "System.Collections.ArrayList"
    $global:sharepoint_metadata = New-Object -TypeName "System.Collections.ArrayList"
    $logs = New-Object -TypeName "System.Collections.ArrayList"
    $output = @()

    $site_list | ForEach-Object {
        #Initialize variables
        $site_title = $_.title
        $site_url = $_.url
        $global:counter = 0
        $item_counter = 0

        #Validate OneDrive/SharePoint URL
        if ($_.url -Like "*-my.sharepoint.com/personal/*") {
            $platform = "OneDrive"
        }
        else {
            $platform = "SharePoint"
        }
        Write-Output "`nPlatform: $platform"
        Write-Output "Name: $($_.Title)"
        Write-Output "URL: $($_.url)"

        if ($platform -eq "OneDrive") {
            try {
                Get-SPOUser -Site $_.url -ErrorAction Stop | Out-Null
                Write-Verbose "Access permissions validated." -Verbose
            }
            catch {
                Write-Warning "Global Admin does not have sufficient permission on this OneDrive."
                return
            }
        }

        #Connect to SharePoint
        try {
            Connect-PnPOnline $_.url -SPOManagementShell -ErrorAction Stop
            Write-Verbose "Successfully connected to site: $($_.Title)" -Verbose
            start-sleep -Seconds 2
        }
        catch {
            Write-Warning "Failed to connect to site: $($_.Title)"
            return
        }

        #Get all list from the SharePoint site and validate if "Documents" list exists
        try {
            Write-Verbose "Querying for all list on this SharePoint site" -Verbose
            $list = Get-PnPList -ErrorAction Stop

            Write-Verbose "Checking if the list contains '$list_name' library" -Verbose
            if ($list.Title -contains $list_name) {
                Write-Verbose "Found '$list_name' library on this SharePoint site" -Verbose
                Write-Verbose "Checking total number of items in '$list_name' library" -Verbose
                $list  = Get-PnPList -Identity $list_name -ErrorAction Stop
                if ($list.ItemCount -eq 0){
                    Write-Warning "No files found on this site."
                    return
                }
                else {
                    Write-Verbose "$($list.ItemCount) items found in '$list_name' library" -Verbose
                }
            }
            else {
                Write-Warning "This site does not contain '$list_name' library"
                return
            }
        }
        catch {
            Write-Warning "Error executing Get-PnPList command"
            return
        }

        #Retrieve each item in the Sharepoint site list
        Write-Verbose "Fetching all items from '$list_name' library matching the specified criteria" -Verbose
        $parameters = @{
            PercentComplete = ($global:counter/$list.ItemCount) * 100
            Activity = "Getting file metadata from Library '$($list.Title)'"
            Status = "Successfully processed $item_counter of $($list.ItemCount) files."
        }
        Write-Progress @parameters
        $script_block = {
            param ($items)
            $global:counter += $items.count
            $percentcomplete = ($global:counter/$list.ItemCount)*100
            $parameters = @{
                PercentComplete = $percentcomplete
                Activity = "Getting file metadata from Library '$($list.Title)'"
                Status = "Successfully processed $counter of $($list.ItemCount) files"
                }
            Write-Progress @parameters
        }
        $parameters= @{
            List = $list_name
            PageSize = $page_size
            Fields = "FileLeafRef", "File_x0020_Type", "FileDirRef", "FileRef", "Created", "Author", "Modified", "Editor", "_dlc_DocId", "File_x0020_Size", "ID", "UniqueID"
            ScriptBlock = $script_block
        }
        $list_items = Get-PnPListItem @parameters |
                Where-Object {
                    ($_.FileSystemObjectType -eq "File") -and
                    ($author_list -contains $_.FieldValues.Author.LookupValue) -and
                    ($_.FieldValues.Created -ge $start_date) -and
                    ($_.FieldValues.Created -le $end_date)
                }
        $parameters = @{
        Activity = "Getting file metadata from Library '$($list.Title)'"
        Status = "Successfully processed $counter of $($list.ItemCount) files"
        Completed = $true
        }
        Write-Progress @parameters
        if ($list_items.count -eq 0){
            Write-Warning "No files matching the specified criteria found in '$list_name' library of this site" -Verbose
            return
        }
        else {
            Write-Verbose "$($list_items.count) files matching the specified criteria found in '$list_name' library of this site" -Verbose
        }

        #Export metadata
        Write-Verbose "Exporting file metadata" -Verbose
        $parameters = @{
            PercentComplete = ($item_counter / ($list_items.Count) * 100)
            Activity = "Exporting file metadata"
            Status = "Successfully processed $item_counter of $($list_items.Count) files."
        }
        Write-Progress @parameters
        $list_items | ForEach-Object {
            [long]$file_size = $_.FieldValues.File_x0020_Size
            if ($file_size -ge 1TB) {
                $auto_file_size = "{0:f2}" -f $($file_size/1TB) + " TB"
            }
            elseif ($file_size -ge 1GB) {
                $auto_file_size = "{0:f2}" -f $($file_size/1GB) + " GB"
            }
            elseif ($file_size -ge 1MB) {
                $auto_file_size = "{0:f2}" -f $($file_size/1MB) + " MB"
            }
            elseif ($file_size -ge 1KB -gt 0) {
                $auto_file_size = "{0:f2}" -f $($file_size/1KB) + " KB"
            }
            else {
                $auto_file_size = "$($file_size)" + " B"
            }
            $output = [PSCustomObject]@{
                DocumentID              = $_.FieldValues._dlc_DocId
                FileName                = $_.FieldValues.FileLeafRef
                FileType                = $_.FieldValues.File_x0020_Type
                FileSizeBytes           = $_.FieldValues.File_x0020_Size
                AutoSize                = $auto_file_size
                CreatedOn               = $_.FieldValues.Created
                CreatedByName           = $_.FieldValues.Author.LookupValue
                CreatedByEmail          = $_.FieldValues.Author.Email
                ModifiedOn              = $_.FieldValues.Modified
                ModifiedByName          = $_.FieldValues.Editor.LookupValue
                ModifiedByEmail         = $_.FieldValues.Editor.Email
                FileRelativeURL         = $_.FieldValues.FileRef
                DirectoryRelativeURL    = $_.FieldValues.FileDirRef
                Platform                = $platform
                Site                    = $site_title
                SiteURL                 = $site_url
            }
            $item_counter++
            $global:sharepoint_metadata.Add($output) | Out-Null
            $metadata.Add($output) | Out-Null

            $timestamp = get-date -Format "dd/MM/yyyy HH:mm:ss"
            $logs.Add("$timestamp - $item_counter of $($list_items.Count) - Exporting metadata from document $($output.FileRelativeURL)") | Out-Null
            if (($item_counter -eq $list_items.Count) -or ($item_counter%10000 -eq 0)) {
                $parameters = @{
                    PercentComplete = ($item_counter / ($list_items.Count) * 100)
                    Activity = "Exporting file metadata"
                    Status = "Successfully processed $item_counter of $($list_items.Count) files."
                }
                Write-Progress @parameters
                $metadata | Export-Csv -Path $report_file -NoTypeInformation -Append
                $logs | Add-Content -Path $report_log_file
                $metadata.Clear()
                $logs.Clear()
            }
        }
    }
    $parameters = @{
        Activity = "Successfully exported file metadata"
        Status = "Completed"
        Completed = $true
    }
    Write-Progress @parameters
    Write-Host "`nA total of $('{0:N0}' -f $sharepoint_metadata.Count) files found in SharePoint Online and OneDrive for Business.`n" -ForegroundColor Green
}