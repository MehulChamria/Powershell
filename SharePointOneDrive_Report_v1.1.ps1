function Get-SharePointReport {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory = $true)]$site_url,
        [Parameter(Mandatory = $true)]$list_name,
        [Parameter(Mandatory = $true)]$output_file,
        [Parameter(Mandatory = $true)]$log_file,
        [Parameter(Mandatory = $true)]$author_list_file,
        [Parameter(Mandatory = $true)]$page_size,
        [Parameter(Mandatory = $true)]$start_date,
        [Parameter(Mandatory = $true)]$end_date
    )

if ($site_url -Like "*-my.sharepoint.com/personal/*") {
    $platform = "OneDrive"
}
elseif ($site_url -Like "*/sites/*") {
    $platform = "SharePoint"
}

#Connect to SharePoint Online site using browser based authentication
try {
    Connect-PnPOnline $site_url -SPOManagementShell -ErrorAction Stop
    Write-Verbose "Successfully connected to site"
}
catch {
    Write-Warning "Failed to connect to site"
    return
}

#Get all Documents from the document library
try {
    $list  = Get-PnPList -Identity $list_name -ErrorAction Stop
}
catch {
    Write-Warning "Error executing Get-PnPList command"
    return
}

if ($list.ItemCount -eq 0){
    Write-Warning "No files found on this site."
    return
}
#Location of text file containing list of all authors
$author_list = Get-Content $author_list_file

#Initialize variables
$global:counter = 0;
$results = New-Object -TypeName "System.Collections.ArrayList"
$logs = New-Object -TypeName "System.Collections.ArrayList"
$item_counter = 0
$csv_counter = 0
$output = @()

#Fetch each item (File) from the list
$list_items =
    Get-PnPListItem -List $list_name `
        -PageSize $page_size `
        -Fields Title, Author, Editor, Created, File_x0020_Type, _CopySource, Created_x0020_By, FileDirRef, ServerUrl, FileRef, Document_x0020_version `
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
                -and ($_.FieldValues.Created -GT $start_date) `
                -and ($_.FieldValues.Created -LT $end_date)}
Write-Verbose "$($list_items.count) files found in Document library."
$global:file_count += $list_items.count

#Iterate through each item and creates required field in to object
Foreach ($item in $list_items)
{
        $output = New-Object PSCustomObject -Property ([ordered]@{
            Name              = $Item["FileLeafRef"]
            Platform          = $platform
            Type              = $Item.FileSystemObjectType
            FileType          = $Item["File_x0020_Type"]
            Site              = $site_url
            RelativeURL       = $Item["FileRef"]
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
    #Get-PnPFile -Url $Results.RelativeURL -Path C:\temp\sp -FileName $Results.Name -AsFile  
}
Write-Verbose "File metadata exported to CSV Successfully"
}

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
    }
    catch{
        Write-Warning "Could not create report folder."
        break
    }
}

#CSV output location
$output_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMddhhmmss)_SharePointOneDrive_report.csv"

#Log file location
$log_file = "$env:USERPROFILE\Desktop\SharePointReport\$(Get-Date -Format yyyMMddhhmmss)_SharepointOneDrive_log.txt"

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
[datetime]$end_date = "2021/01/21 00:25:00"

#Initialize Output and log csv file 
Set-Content $output_file -Value $null
Set-Content $log_file -Value $null

Write-Information "`nInitiating inventory for SharePoint/OneDrive sites." -InformationAction Continue

$sharepoint_site_list.URL | ForEach-Object {
    Write-Output "`nPlatform: SharePoint"
    Write-Output "Site name: $_"
    $parameters = @{
        site_url = $_
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

$onedrive_site_list | ForEach-Object {
    $parameters = @{
    site_url = $_.URL
    list_name =$list_name
    output_file = $output_file
    log_file = $log_file
    author_list_file = $author_list_file
    page_size = $page_size
    start_date = $start_date
    end_date = $end_date
    }

    $User = $_.Title
    $OneDrive_URL = $_.URL
    Write-Output "`nPlatform: OneDrive"
    Write-Output "User: $User"
    Write-Output "Site: $OneDrive_URL"
    
    try {
        Get-SPOUser -Site $OneDrive_URL | Out-Null
        Write-Verbose "Access permissions validated." -Verbose
        Get-SharePointReport @Parameters -Verbose
    }
    catch {
        Write-Warning "Global Admin does not have sufficient permission." -WarningAction Continue
    }
}

Write-Host "`nA total of $('{0:N0}' -f $file_count) files found in SharePoint Online." -ForegroundColor Green