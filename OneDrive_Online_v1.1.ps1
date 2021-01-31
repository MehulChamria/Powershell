#Config Variables
$SharePointAdminURL =  "https://indianassociationoldham-admin.sharepoint.com/"
#$credentials = get-credential -Message "Login using a global admin account"
Connect-SPOService -Url $SharePointAdminURL

$onedrive_sites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"
#$onedriveurl = "https://indianassociationoldham-my.sharepoint.com/personal/iao_indianassociationoldham_onmicrosoft_com"

Write-Host "`nAdding globaladmin as site collection admin on OneDrive site collection" -ForegroundColor Blue
# Set current admin as a Site Collection Admin on both OneDrive Site Collections
Set-SPOUser -Site $onedriveurl -LoginName "mehul@indianassociationoldham.onmicrosoft.com" -IsSiteCollectionAdmin $true

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
$list  = Get-PnPList -Identity $ListName
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

$site_url = $onedriveurl

#ListName for eg: Documents. To find out list use the following command and use the Title from output as list name: Get-PNPList
$list_name= "Documents"

#CSV output location
$output_file = "C:\temp\$(Get-Date -Format yyyMMdd)_onedrive_report.csv"

#Log file location
$log_file = "C:\temp\$(Get-Date -Format yyyMMdd)_log.txt"

#Author file location
$author_list_file = "C:\Temp\author_list.txt"

#Set page size
$page_size = 5000

#Date range in format YYYY/MM/DD HH:MM:SS
[datetime]$start_date = "2021/01/10 02:25:00"
[datetime]$end_date = "2021/01/15 00:25:00"

#Initialize Output and log csv file 
Set-Content $output_file -Value $null
Set-Content $log_file -Value $null

$parameters = @{
    site_url = $site_url
    list_name =$list_name
    output_file = $output_file
    log_file = $log_file
    author_list_file = $author_list_file
    page_size = $page_size
    start_date = $start_date
    end_date = $end_date
}

Get-SharePointReport @parameters