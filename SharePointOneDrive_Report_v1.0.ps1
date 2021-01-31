Function Set-OneDrivePermissions {
    param(
        [Parameter(Mandatory = $true)]$Global_Admin_UserID,
        [Parameter(Mandatory = $true)]$User,
        [Parameter(Mandatory = $true)]$OneDrive_URL
    )
    Write-Host "Adding $Global_Admin_UserID as site collection admin for user's OneDrive."
        try{
            Set-SPOUser -Site $OneDrive_URL -LoginName $Global_Admin_UserID -IsSiteCollectionAdmin $true | Out-Null
            Write-Verbose "Permissions applied successfully." -Verbose
        }
        catch{
            Write-Warning "Failed to apply permissions." -WarningAction Continue
        }
}

Function Get-OneDrivePermissions {
    param(
        [Parameter(Mandatory = $true)]$Global_Admin_UserID,
        [Parameter(Mandatory = $true)]$User,
        [Parameter(Mandatory = $true)]$OneDrive_URL,
        [Parameter(Mandatory = $true)]$Parameters,
        [Parameter(Mandatory = $false)][Switch]$Appy_Permissions
    )
    Write-Host "`nChecking if $Global_Admin_UserID has sufficient rights" -InformationAction Continue
    try {
        Get-SPOUser -Site $OneDrive_URL | Out-Null
        Write-Verbose "Access permissions validated." -Verbose
        Get-SharePointReport @Parameters
    }
    catch {
        Write-Warning "Global Admin does not have sufficient permission." -WarningAction Continue
        if ($Appy_Permissions) {
            Set-OneDrivePermissions -Global_Admin_UserID $Global_Admin_UserID -User $User -OneDrive_URL $OneDrive_URL
        }
        $User
    }
}

function Get-SharePointReport {
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
#Connect to SharePoint Online site using browser based authentication
Connect-PnPOnline $site_url -SPOManagementShell

#Get all Documents from the document library
$list  = Get-PnPList -Identity $list_name
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
            $percentcomplete = ($items.count/$global:counter)*100
            $parameters = @{
                PercentComplete = $percentcomplete
                Activity = "Getting documents metadata from Library '$($list.Title)'"
                Status = "Successfully processed $counter of $($list.ItemCount) documents"
            }
            Write-Progress @parameters
        } |
        Where-Object {($_.FileSystemObjectType -eq "File") `
                -and ($author_list -contains $_.fieldvalues.Author.LookupValue) `
                -and ($_.FieldValues.Created -GT $start_date) `
                -and ($_.FieldValues.Created -LT $end_date)}

#Iterate through each item and creates required field in to object
Foreach ($item in $list_items)
{
        $output = New-Object PSCustomObject -Property ([ordered]@{
            Name              = $Item["FileLeafRef"]
            Type              = $Item.FileSystemObjectType
            FileType          = $Item["File_x0020_Type"]
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
            "PercentComplete" = ($item_counter / ($list_items.Count) * 100)
            "Activity" = "Successfully processed $item_counter of $($list_items.Count) documents"
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

Write-host "Document Library Inventory Exported to CSV Successfully!"
}

$SharePointAdminURL =  "https://indianassociationoldham-admin.sharepoint.com/"
$Global_Admin_UserID = "mehul@indianassociationoldham.onmicrosoft.com"
Connect-SPOService -Url $SharePointAdminURL

$site_list = Get-SPOSite -IncludePersonalSite $true -Limit all
$onedrive_site_list = $site_list | Where-Object URL -Like "*-my.sharepoint.com/personal/*"
$sharepoint_site_list = $site_list | Where-Object URL -Like "*/sites/*"

#ListName for eg: Documents. To find out list use the following command and use the Title from output as list name: Get-PNPList
$list_name= "Documents"

#CSV output location
$output_file = "C:\temp\$(Get-Date -Format yyyMMdd)_SharePointOneDrive_report.csv"

#Log file location
$log_file = "C:\temp\$(Get-Date -Format yyyMMdd)_log.txt"

#Author file location
$author_list_file = "C:\Temp\author_list.txt"

#Set page size
$page_size = 5000

#Date range in format YYYY/MM/DD HH:MM:SS
[datetime]$start_date = "2021/01/10 02:25:00"
[datetime]$end_date = "2021/01/21 00:25:00"

#Initialize Output and log csv file 
Set-Content $output_file -Value $null
Set-Content $log_file -Value $null

$sharepoint_site_list.URL | ForEach-Object {
    Write-Information "Scanning $_" -InformationAction Continue
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

$NoPermissions_Profiles_list = $null
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
    Write-Host "`nScanning following user's OneDrive: $User"
    $NoPermissions_Profiles_list += Get-OneDrivePermissions -Global_Admin_UserID $Global_Admin_UserID -User $User -OneDrive_URL $OneDrive_URL -Parameters $parameters
}