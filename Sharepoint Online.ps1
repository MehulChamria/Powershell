Function Write-Log{
    param(
        [Parameter(Mandatory = $true)][String]$msg,
        [Parameter(Mandatory = $true)][String]$logfile
    )
    Add-Content $logfile $msg
}

#Enter the name of sharepoint site https://<SHAREPOINTSITENAME.com>/sites/<SITENAME>
$SiteURL = "https://indianassociationoldham.sharepoint.com/sites/public"

#ListName for eg: Documents. To find out list use the following command and use the Title from output as list name: Get-PNPList
$ListName= "Documents"

#CSV output location
$ReportOutput = "C:\temp\filelist.csv"
Set-Content $ReportOutput -Value $null

#Log file location
$logfile = "C:\temp\log.txt"

#Set page size
$Pagesize = 5000

#Array to store results
$Results = @()

#Connect to SharePoint Online site using browser based authentication
Connect-PnPOnline $SiteURL -UseWebLogin

#Get all Documents from the document library
$List  = Get-PnPList -Identity $ListName
$counter = 0;

#Location of text file containing list of all authors
$createdBy = Get-Content C:\temp\name.txt

#Date range in format YYYY/MM/DD HH:MM:SS
[datetime]$startDate = "2021/01/10 02:25:00"
[datetime]$EndDate = "2021/01/15 00:25:00"

#Get the each item (File) from the list
$ListItems =
    Get-PnPListItem -List $ListName `
        -PageSize $Pagesize `
        -Fields Title,Author, Editor, Created, File_x0020_Type,_CopySource,Created_x0020_By,FileDirRef,ServerUrl,FileRef,Document_x0020_version `
        -ScriptBlock {
            param ($items)
            $global:counter += $items.Count
            $parameters = @{
                PercentComplete = ($counter / ($List.ItemCount) * 100)
                Activity = "Getting Documents from Library '$($List.Title)'"
                Status = "Successfully processed $counter of $($List.ItemCount) documents"
            }
            Write-Progress @parameters
        } |
        Where-Object {($_.FileSystemObjectType -eq "File") `
                -and ($createdBy -contains $_.fieldvalues.Author.LookupValue) `
                -and ($_.FieldValues.Created -GT $startDate) `
                -and ($_.FieldValues.Created -LT $EndDate)}

$Output = New-Object -TypeName "System.Collections.ArrayList"
$Logs = New-Object -TypeName "System.Collections.ArrayList"
$ItemCounter = 0
$csvcounter = 0

#Iterate through each item and creates required field in to object
Foreach ($Item in $ListItems)
{
        $Results = New-Object PSCustomObject -Property ([ordered]@{
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
    $Output.Add($Results)
    $Logs.Add("$ItemCounter of $($ListItems.Count) - Exporting metadata from document $($Results.Name)")
    $ItemCounter++
    $csvcounter++

    #<#
    if (($csvcounter -eq 10000) -or ($ItemCounter -eq $ListItems.Count)) {
        $parameters = @{
            "PercentComplete" = ($ItemCounter / ($ListItems.Count) * 100)
            "Activity" = "Successfully processed $ItemCounter of $($ListItems.Count) documents"
        }
        Write-Progress @parameters

        $Output | Export-Csv -Path $ReportOutput -NoTypeInformation -Append
        $Output.Clear()
        
        Write-Log -msg $Logs -logfile $logfile
        $Logs.Clear()
        
        $csvcounter = 0
    }#>
    #Get-PnPFile -Url $Results.RelativeURL -Path C:\temp\sp -FileName $Results.Name -AsFile  
}

Write-host "Document Library Inventory Exported to CSV Successfully!"