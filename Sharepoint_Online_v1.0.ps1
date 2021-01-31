Function Write-Log{
    param(
        [Parameter(Mandatory = $true)][String]$msg,
        [Parameter(Mandatory = $true)][String]$logfile
    )
    Add-Content $logfile $msg
}

#Enter the name of sharepoint site https://<SHAREPOINTSITENAME.com>/sites/<SITENAME>
$site_url = "https://indianassociationoldham.sharepoint.com/sites/public"

#ListName for eg: Documents. To find out list use the following command and use the Title from output as list name: Get-PNPList
$list_name= "Documents"

#CSV output location
$output_file = "C:\temp\$(Get-Date -Format yyyMMdd)_sharepoint_report.csv"

#Log file location
$log_file = "C:\temp\$(Get-Date -Format yyyMMdd)_log.txt"

#Author file location
$author_list_file = "C:\Temp\author_list.txt"

#Set page size
$page_size = 5000

#Date range in format YYYY/MM/DD HH:MM:SS
[datetime]$start_date = "2021/01/10 02:25:00"
[datetime]$end_date = "2021/01/15 00:25:00"

#Connect to SharePoint Online site using browser based authentication
Connect-PnPOnline $site_url -UseWebLogin

#Initialize Output csv file 
Set-Content $output_file -Value $null
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
            $parameters = @{
                PercentComplete = ($counter / ($list.ItemCount) * 100)
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
            CreatedOn         = $Item["Created"]
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

    #<#
    if (($csv_counter -eq 10000) -or ($item_counter -eq $list_items.Count)) {
        $parameters = @{
            "PercentComplete" = ($item_counter / ($list_items.Count) * 100)
            "Activity" = "Successfully processed $item_counter of $($list_items.Count) documents"
        }
        Write-Progress @parameters

        $results | Export-Csv -Path $output_file -NoTypeInformation -Append
        $results.Clear()
        
        Write-Log -msg $logs -logfile $log_file
        $logs.Clear()
        
        $csv_counter = 0
    }#>
    #Get-PnPFile -Url $Results.RelativeURL -Path C:\temp\sp -FileName $Results.Name -AsFile  
}

Write-host "Document Library Inventory Exported to CSV Successfully!"