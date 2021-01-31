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

#region Validate OneDrive/SharePoint URL

# Check if the URL is SharePoint for OneDrive site
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

#endregion

#region Connect to SharePoint

try {
    Connect-PnPOnline $site_url -SPOManagementShell -ErrorAction Stop
    Write-Verbose "Successfully connected to site" -Verbose
}
catch {
    Write-Warning "Failed to connect to site"
    return
}

#endregion

#region Validate list

#Get all lists from the SharePoint site ad validate if "Documents" list exists
try {
    $pnp_list = Get-PnPList -ErrorAction Stop
    if ($pnp_list.Title -contains $list_name) {
        $list  = Get-PnPList -Identity $list_name -ErrorAction Stop
    }
    else {
        Write-Warning "This site does not contain `"Documents`" list"
        return
    }
}
catch {
    Write-Warning "Error executing Get-PnPList command"
    return
}

#endregion

#region list item count

#Check if the list contains any files
if ($list.ItemCount -eq 0){
    Write-Warning "No files found on this site."
    return
}

#endregion

#Location of text file containing list of all authors
$author_list = Get-Content $author_list_file

#region Initialize variables

$global:counter = 0;
$results = New-Object -TypeName "System.Collections.ArrayList"
$logs = New-Object -TypeName "System.Collections.ArrayList"
$item_counter = 0
$csv_counter = 0
$output = @()

#endregion

#region Get-PnPListItem

#List each item in the Sharepoint site list
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

Write-Verbose "$($list_items.count) files found in Documents library." -Verbose

#endregion

#region Export metadata

Foreach ($item in $list_items) {
    $output = New-Object PSCustomObject -Property ([ordered]@{
        FileName          = $Item["FileLeafRef"]
        Platform          = $platform
        Type              = $Item.FileSystemObjectType
        FileType          = $Item["File_x0020_Type"]
        SiteName          = $site_title
        SiteURL           = $site_url
        FileRelativeURL   = $Item["FileRef"]
        Created           = $Item["Created"]
        CreatedByName     = $Item.fieldvalues.Author.LookupValue
        CreatedByEmail    = $Item["Author"].Email
        Modified          = $Item["Modified"]
        ModifiedByName    = $Item.fieldvalues.Editor.LookupValue
        ModifiedByEmail   = $Item["Editor"].Email
        DocumentID        = $Item["_dlc_DocId"]
    })
    
    $item_counter++
    $csv_counter++
    $report.Add($output) | Out-Null

    $results.Add($output) | Out-Null
    $logs.Add("$item_counter of $($list_items.Count) - Exporting metadata from document $($output.Name)") | Out-Null

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

#endregion

$global:file_count += $list_items.count
Write-Verbose "File metadata exported to CSV Successfully" -Verbose
}

function Get-SharePointFiles{
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]$report,
        [Parameter(Mandatory = $true)]$log_file
        )
    
    $report | Group-Object SiteURL | ForEach-Object {
        $siteURL = $_.Name
        $file_list = $_.Group

#region Connect to SharePoint

        try {
            Connect-PnPOnline $_.Name -SPOManagementShell -ErrorAction Stop
            Write-Verbose "Successfully connected to site" -Verbose
        }
        catch {
            Write-Warning "Failed to connect to site"
            return
        }

#endregion

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
                Get-PnPFile -Url $_.FileRelativeURL -Path $folderpath -AsFile -WarningAction Stop
                Write-Verbose "$($_.FileName) downloaded successfully" -Verbose
                Add-Content -Value "$($_.FileRelativeURL) downloaded successfully" -Path $log_file
            }
            else {
                Write-Verbose "$($_.FileName) already exists, skipping download." -Verbose
            }
        }
    }
}

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
$SharePointAdminURL =  "https://indianassociationoldham-admin.sharepoint.com/"
$Global_Admin_UserID = "mehul@indianassociationoldham.onmicrosoft.com"
Connect-SPOService -Url $SharePointAdminURL

$site_list = Get-SPOSite -IncludePersonalSite $true -Limit all
$onedrive_site_list = $site_list | Where-Object URL -Like "*-my.sharepoint.com/personal/*"
$sharepoint_site_list = $site_list | Where-Object URL -Like "*/sites/*"

$global:file_count = 0

#ListName for eg: Documents. To find out list use the following command and use the Title from output as list name: Get-PNPList
$list_name= "Documents"

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

#CSV output location
$output_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMddhhmm)_SharePointOneDrive_report.csv"

#Log file location
$log_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMddhhmm)_SharepointOneDrive_log.txt"

#Author file location
$author_list_file = "C:\Temp\author_list.txt"
if (-not (Test-Path $author_list_file)){
    Write-Warning "Author list file not found."
    break
}
elseif ($null -eq (Get-Content $author_list_file)){
    Write-Warning "Author list file contains no records."
    break
}

#Set page size
$page_size = 5000

#Date range in format YYYY/MM/DD HH:MM:SS
[datetime]$start_date = "2021/01/10 02:25:00"
[datetime]$end_date = "2021/01/23 00:25:00"

#Initialize Output and log csv file 
Set-Content $output_file -Value $null
Set-Content $log_file -Value $null

$global:report = New-Object -TypeName "System.Collections.ArrayList"

#endregion Initialize variables

Write-Output "`nInitiating scan for SharePoint/OneDrive sites."

$site_list | ForEach-Object {
    $parameters = @{
        site_url = $_.URL
        site_title = $_.Title
        list_name =$list_name
        output_file = $output_file
        log_file = $log_file
        author_list_file = $author_list_file
        page_size = $page_size
        start_date = $start_date
        end_date = $end_date
    }
    Get-SharePointReport @parameters
}

Write-Host "`nA total of $('{0:N0}' -f $file_count) files found in SharePoint Online and OneDrive for Business." -ForegroundColor Green

Get-SharePointFiles -report $report -log_file $log_file