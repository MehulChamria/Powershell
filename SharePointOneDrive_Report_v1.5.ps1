#region Define variables
#SharePoint Admin URL
$sharepoint_admin_url =  "https://indianassociationoldham-admin.sharepoint.com/"

#Author file location
$author_list_file = "C:\Temp\author_list.txt"

#ListName for eg: Documents. To find out list use the following command and use the Title from output as list name: Get-PNPList
$list_name= "Documents"

#CSV output location
$report_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMdd_hhmm)_SPReport.csv"

#Log file location
$report_log_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMdd_hhmm)_SPReport_log.txt"
$downloads_log_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMdd_hhmm)_SPDownloads_log.txt"
$summary_log_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMdd_hhmm)_SPSummary_log.txt"
#Set page size
$page_size = 5000

#Date range in format YYYY/MM/DD HH:MM:SS
[datetime]$start_date = "2021/01/10 02:25:00"
[datetime]$end_date = "2021/01/23 00:25:00"
#endregion

#region Initialize variables
#Verify if the SharePointReport folder exists, if not, create it.
if(-not (Test-Path $env:USERPROFILE\Desktop\SharePointReport)) {
    try{
        New-Item -ItemType Directory -Path $env:USERPROFILE\Desktop\SharePointReport -ErrorAction Stop | Out-Null
        New-Item -ItemType Directory -Path $env:USERPROFILE\Desktop\SharePointReport\Downloads -ErrorAction Stop | Out-Null
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

#Initialize Output and log csv file 
Set-Content $report_file -Value $null
Set-Content $report_log_file -Value $null
Set-Content $downloads_log_file -Value $null

$global:report = New-Object -TypeName "System.Collections.ArrayList"
$global:file_count = 0
#endregion Initialize variables

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

#region Connect to SharePoint Admin site
Connect-SPOService -Url $sharepoint_admin_url
$site_list = Get-SPOSite -IncludePersonalSite $true -Limit all |
                Where-Object {($_.URL -Like "*-my.sharepoint.com/personal/*") -or ($_.URL -Like "*/sites/*")}
#endregion

#region Function definition: Get-SharePointReport 
function Get-SharePointReport {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]$site_url,
        [Parameter(Mandatory = $true)]$site_title,
        [Parameter(Mandatory = $true)]$list_name,
        [Parameter(Mandatory = $true)]$output_file,
        [Parameter(Mandatory = $true)]$log_file,
        [Parameter(Mandatory = $true)]$author_list_file,
        [Parameter(Mandatory = $true)]$page_size,
        [Parameter(Mandatory = $true)]$start_date,
        [Parameter(Mandatory = $true)]$end_date
    )

#Initialize variables
$author_list = Get-Content $author_list_file
$global:counter = 0;
$results = New-Object -TypeName "System.Collections.ArrayList"
$logs = New-Object -TypeName "System.Collections.ArrayList"
$item_counter = 0
$csv_counter = 0
$output = @()

#Validate OneDrive/SharePoint URL
if ($site_url -Like "*-my.sharepoint.com/personal/*") {
    $platform = "OneDrive"
    Write-Output "`nPlatform: OneDrive"
    Write-Output "Name: $site_title"
    Write-Output "URL: $site_url"
    try {
        Get-SPOUser -Site $site_url -ErrorAction Stop | Out-Null
        Write-Verbose "Access permissions validated." -Verbose
    }
    catch {
        Write-Warning "Global Admin does not have sufficient permission on this OneDrive."
        return
    }
}
else {
    $platform = "SharePoint"
    Write-Output "`nPlatform: SharePoint"
    Write-Output "Site Name: $site_title"
    Write-Output "URL: $site_url"
}

#Connect to SharePoint
try {
    Connect-PnPOnline $site_url -SPOManagementShell -ErrorAction Stop
    Write-Verbose "Successfully connected to site" -Verbose
    start-sleep -Seconds 2
}
catch {
    Write-Warning "Failed to connect to site"
    return
}

#Get all list from the SharePoint site and validate if "Documents" list exists
try {
    Write-Verbose "Querying for all list on this SharePoint site" -Verbose
    $pnp_list = Get-PnPList -ErrorAction Stop

    Write-Verbose "Checking if the list contains '$list_name' list" -Verbose
    if ($pnp_list.Title -contains $list_name) {
        Write-Verbose "Found '$list_name' list on this SharePoint site" -Verbose
        Write-Verbose "Checking total number of items in '$list_name' list" -Verbose
        $list  = Get-PnPList -Identity $list_name -ErrorAction Stop
        Write-Verbose "$($list.ItemCount) items found in '$list_name' list"  -Verbose
    }
    else {
        Write-Warning "This site does not contain '$list_name' list"
        return
    }
}
catch {
    Write-Warning "Error executing Get-PnPList command"
    return
}

#Check if the list contains any files
if ($list.ItemCount -eq 0){
    Write-Warning "No files found on this site."
    return
}

#Retrieve each item in the Sharepoint site list
Write-Verbose "Retrieving all items from 'Documents' list on this site" -Verbose
$list_items =
    Get-PnPListItem -List $list_name `
        -PageSize $page_size `
        -Fields Title, Author, Editor, Created, File_x0020_Type, _CopySource, Created_x0020_By, FileDirRef, ServerUrl, FileRef, Document_x0020_version, _dlc_DocId `
        -ScriptBlock {
            param ($items)
            $global:counter += $items.count
            $percentcomplete = ($global:counter/$list.ItemCount)*100
            $parameters = @{
                PercentComplete = $percentcomplete
                Activity = "Getting file metadata from Library '$($list.Title)'"
                Status = "Successfully processed $counter of $($list.ItemCount) files"
                }
            Write-Progress @parameters
        } |
        Where-Object {($_.FileSystemObjectType -eq "File") `
                -and ($author_list -contains $_.fieldvalues.Author.LookupValue) `
                -and ($_.FieldValues.Created -ge $start_date) `
                -and ($_.FieldValues.Created -le $end_date)}

Write-Verbose "$($list_items.count) files matching the specified criteria found in '$list_name' list" -Verbose

if ($list_items.count -eq 0){
    return
}

#Export metadata
Write-Verbose "Exporting file metadata" -Verbose
$parameters = @{
    PercentComplete = ($item_counter / ($list_items.Count) * 100)
    Activity = "Initializing export of file metadata"
    Status = "Successfully processed $item_counter of $($list_items.Count) files."
}
Write-Progress @parameters
Foreach ($item in $list_items) {
    $output = [PSCustomObject][Ordered]@{
        Platform          = $platform
        Type              = $Item.FileSystemObjectType
        FileType          = $Item["File_x0020_Type"]
        SiteName          = $site_title
        FileName          = $Item["FileLeafRef"]
        SiteURL           = $site_url
        FileRelativeURL   = $Item["FileRef"]
        Created           = $Item["Created"]
        CreatedByName     = $Item.fieldvalues.Author.LookupValue
        CreatedByEmail    = $Item["Author"].Email
        Modified          = $Item["Modified"]
        ModifiedByName    = $Item.fieldvalues.Editor.LookupValue
        ModifiedByEmail   = $Item["Editor"].Email
        DocumentID        = $Item["_dlc_DocId"]
    }
    
    $item_counter++
    $csv_counter++
    $report.Add($output) | Out-Null

    $results.Add($output) | Out-Null
    $timestamp = get-date -Format "dd/MM/yyyy HH:mm:ss"
    $logs.Add("$timestamp - $item_counter of $($list_items.Count) - Exporting metadata from document $($output.FileRelativeURL)") | Out-Null

    if (($csv_counter -eq 10000) -or ($item_counter -eq $list_items.Count)) {
        $parameters = @{
            PercentComplete = ($item_counter / ($list_items.Count) * 100)
            Activity = "Exporting file metadata."
            Status = "Successfully processed $item_counter of $($list_items.Count) files."
        }
        Write-Progress @parameters

        $results | Export-Csv -Path $output_file -NoTypeInformation -Append
        $results.Clear()
        
        $logs | Add-Content -Path $log_file
        $logs.Clear()
        
        $csv_counter = 0
    }
}

$global:file_count += $list_items.count
Write-Verbose "Successfully exported file metadata" -Verbose
}
#endregion

#region Function definition: Get-SharePointFiles
function Get-SharePointFiles{
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]$report,
        [Parameter(Mandatory = $true)]$log_file
        )
    
    $total_file_count = ($report.group).count
    $failed_file_list = New-Object -TypeName "System.Collections.ArrayList"
    $statistics = @{}
    $downloaded_file_count = 0
    $skipped_file_count = 0
    $processed_file_count = 0
    $failed_file_count = 0
    $download_start_time = get-date
    $download_end_time = get-date

    $report | Group-Object SiteURL | ForEach-Object {
        $file_list = $_.Group
        $sharepoint_site = $_.Name

#Connect to SharePoint
        try {
            Connect-PnPOnline $sharepoint_site -SPOManagementShell -ErrorAction Stop
            Write-Verbose "Successfully connected to site" -Verbose
        }
        catch {
            Write-Warning "Failed to connect to site"
            return
        }

        $file_list | ForEach-Object {
            $relative_path = Split-Path $_.FileRelativeURL.Replace("/","\")
            $folderpath = "$env:USERPROFILE\Desktop\SharePointReport\Downloads$relative_path"
            $filepath = "$env:USERPROFILE\Desktop\SharePointReport\Downloads$relative_path\$($_.FileName)"
            if (-not (Test-Path $folderpath)) {
                try {
                    New-Item -ItemType Directory $folderpath -ErrorAction Stop | Out-Null
                }
                catch {
                    Write-Warning "Failed to create folder path $folderpath"
                    return
                }
            }
            if(-not (Test-Path $filepath)) {
                try {
                    Get-PnPFile -Url $_.FileRelativeURL -Path $folderpath -AsFile -ErrorAction Stop
                    $downloaded_file_count++
                    $timestamp = get-date -Format "dd/MM/yyyy HH:mm:ss"
                    Write-Verbose "$timestamp - $($_.FileName) downloaded successfully" -Verbose
                    Add-Content -Value "$timestamp - $($_.FileRelativeURL) downloaded successfully" -Path $log_file
                }
                catch {
                    $Error[0]
                    Write-Warning "Error downloading file $($_.FileName)"
                    Add-Content -Value "$timestamp - $($_.FileRelativeURL) Failed to download" -Path $log_file
                    $failed_file_list.Add($_)
                    $failed_file_count++
                }
            }
            else {
                $skipped_file_count++
                $timestamp = get-date -Format "dd/MM/yyyy HH:mm:ss"
                Add-Content -Value "$timestamp - $($_.FileRelativeURL) already exists, skipping download." -Path $log_file
                Write-Verbose "$timestamp - $($_.FileName) already exists, skipping download." -Verbose
            }

            $processed_file_count = $failed_file_count + $downloaded_file_count + $skipped_file_count
            $remaining_file_count = $total_file_count - $processed_file_count
            $download_end_time = get-date
            $download_elapsed_time = New-TimeSpan -Start $download_start_time -End $download_end_time
            $parameters = @{
                PercentComplete = (($downloaded_file_count / $total_file_count) * 100)
                Activity = "Total Files: $total_file_count | Downloaded: $downloaded_file_count | Failed: $failed_file_count | Skipped: $skipped_file_count | Remaining: $remaining_file_count | Elapsed Time: $download_elapsed_time"
                Status = "$timestamp - Processing $($_.FileRelativeURL)"
            }
            Write-Progress @parameters
            Add-Content -Value "$timestamp - $processed_file_count of $total_file_count files processed." -Path $log_file
        }
    }

    $statistics = [PSCustomObject]@{
        "Total" = $total_file_count
        "Successful" = $downloaded_file_count
        "Failed" = $failed_file_count
        "Skipped" = $skipped_file_count
        "Remaining" = $remaining_file_count
        "Start time" = $download_start_time
        "End time" = $download_end_time
        "Elapsed time" = $download_elapsed_time
    }
    Write-Output "File downloads statistics"
    Write-Output $statistics
    $statistics | Out-File $log_file -Append
}
#endregion

#region Function call: Get-SharePointReport
Start-Transcript -Path $summary_log_file -Append
Write-Output "`nInitiating scan for SharePoint/OneDrive sites."
$site_list | ForEach-Object {
    $parameters = @{
        site_url = $_.URL
        site_title = $_.Title
        list_name =$list_name
        output_file = $report_file
        log_file = $report_log_file
        author_list_file = $author_list_file
        page_size = $page_size
        start_date = $start_date
        end_date = $end_date
    }
    Get-SharePointReport @parameters
}

Write-Progress -Activity "Successfully exported metadata from all sites" -Status "Completed" -Completed
Write-Host "`nA total of $('{0:N0}' -f $file_count) files found in SharePoint Online and OneDrive for Business." -ForegroundColor Green
Stop-Transcript
#endregion

#region Function call: Get-SharePointFiles
Get-SharePointFiles -report $report -log_file $downloads_log_file
Read-Host "All files have been processed. Press enter to exit"
Write-Progress -Activity "Successfully downloaded files from all sites" -Status "Completed" -Completed
#endregion