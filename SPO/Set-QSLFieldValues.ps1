[CmdletBinding()]
param(
    
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=1, HelpMessage="URL of the site where the QSL list resides")]
    [string] $siteURL="https://murphyoil.sharepoint.com/GlobalProcurement/purchasing",

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Display name of the QSL list")]
    [string] $listName="QSL/Procurement Plan Database",

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Source field to extract data from")]
    [string] $sourceFieldName="CascadingField2",

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Comma separated IDs of Specific list items to update")]
    [int[]] $listItemIDs=@(1243),

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Run for all items in the list?")]
    [switch] $processAll = $true
)

$ErrorActionPreference = "Continue"

function Write-Message {
    param (
        [string] $msg, 
        [switch] $NoNewLine = $false, 
        [System.ConsoleColor] $BackgroundColor = "DarkBlue", 
        [System.ConsoleColor] $ForegroundColor = "White"
    )
    try{
        if(!$NoNewLine) {
            Write-Host $msg -BackgroundColor $backgroundColor -f $foregroundColor
            Write-Output $msg | out-file $global:logFilePath -Append -ErrorAction Stop
        }
        else {
            Write-Host $msg -BackgroundColor $backgroundColor -f $foregroundColor -NoNewline
            Write-Output $msg | out-file $global:logFilePath -NoNewline -Append -ErrorAction Stop
        }
    }
    catch{}
}

#Function to get the ID from Parent Lookup List - Based on Provided value 
function Get-ParentLookupID($parentListName, $parentListLookupField, $parentListLookupValue, $parentListFilterField, $parentListFilterFieldValue) 
{ 
    $parentList = $global:ctx.Web.lists.GetByTitle($parentListName)
     
    #Get the Parent List Item Filtered by given Lookup Value 
    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    if($parentListFilterField -eq $null) { 
        $query.ViewXml="<View Scope='RecursiveAll'><Query><Where><BeginsWith><FieldRef Name='$($parentListLookupField)'/><Value Type='Text'>$($parentListLookupValue)</Value></BeginsWith></Where></Query></View>"
    }
    else {
        $query.ViewXml="<View Scope='RecursiveAll'><Query><Where><And><BeginsWith><FieldRef Name='$($parentListLookupField)'/><Value Type='Text'>$($parentListLookupValue)</Value></BeginsWith><BeginsWith><FieldRef Name='$($parentListFilterField)'/><Value Type='Text'>$($parentListFilterFieldValue)</Value></BeginsWith></And></Where></Query></View>"
    }
    
    $listItems = $parentList.GetItems($query) 
    
    $global:ctx.Load($listItems) 
    $global:ctx.ExecuteQuery() 
    
    #Get the ID of the List Item 
    If($listItems.count -gt 0) { 
        return $listItems[0].ID            #Get the first item - If there are more than One 
    } 
    else { 
        return $null 
    } 
} 

function Set-LookupFieldValue($listItem, $listItemLookupFieldToSet, $parentListName, $parentListLookupField, $parentListLookupValue, $parentListFilterField, $parentListFilterFieldValue){
    $success = $false

    $lookupID = Get-ParentLookupID $parentListName $parentListLookupField $parentListLookupValue $parentListFilterField $parentListFilterFieldValue
    
    if($lookupID -ne $null) { 
        #Update Lookup Field using Parent Lookup Item ID 
        $listItem[$listItemLookupFieldToSet] = $lookupID 
        $listItem.Update()

        $success = $true
    }

    return $success
}

function Process-Item($listName, $listItemID) {
    Write-Message "`n-----------------------------------------------------------------------------------"
    Write-Message "Processing item id '$listItemID' on list '$listName'..."

    $list = $global:ctx.Web.lists.GetByTitle($listName) 
    $listItem = $list.GetItemById($listItemID)

    Write-Message "Retrieving list item..." -NoNewLine

    $global:ctx.Load($listItem) 
    $global:ctx.ExecuteQuery()

    Write-Message "Done" -BackgroundColor Green

    $sourceField = $listItem[$sourceFieldName]

    if($sourceField -ne $null){
        Write-Message "`nSource field '$sourceFieldName' = $($sourceField.ToString())"

        switch($sourceFieldName){
            "CascadingField1"{
                $materialGroupMatches = ($sourceField.ToString() | Select-String -pattern '{"Material Group":"([^"]*)').Matches
                if($materialGroupMatches -ne $null) {
                    $materialGroup = $materialGroupMatches.Groups[1].Value

                    Write-Message "Setting 'QSLMaterialGroup' field for item id '$listItemID' to value '$materialGroup'..." -NoNewLine

                    $materialGroupSetResult = Set-LookupFieldValue $listItem "QSLMaterialGroup" "QSLMaterialGroup" "Title" $materialGroup $null $null

                    if($materialGroupSetResult) {
                        Write-Message "Done" -BackgroundColor Green
                    }
                    else {
                        Write-Message "Failed" -BackgroundColor Red
                    }
                }
                
                $categoryMatches = ($sourceField.ToString() | Select-String -pattern '{"Category":"([^"]*)').Matches
                if($categoryMatches -ne $null) {
                    $category = $categoryMatches.Groups[1].Value

                    Write-Message "Setting 'QSLCategory' field for item id '$listItemID' to value '$category'..." -NoNewLine

                    $materialCategorySetResult = Set-LookupFieldValue $listItem "QSLCategory" "QSLMaterialCategory" "Title" $category "MaterialGroup" $materialGroup
                    
                    if($materialCategorySetResult) {
                        Write-Message "Done" -BackgroundColor Green
                    }
                    else {
                        Write-Message "Failed" -BackgroundColor Red
                    }
                }

                $subCategoryMatches = ($sourceField.ToString() | Select-String -pattern '{"Sub-Category":"([^"]*)').Matches
                if($subCategoryMatches -ne $null) {
                    $subCategory = $subCategoryMatches.Groups[1].Value

                    Write-Message "Setting 'QSLSubCategory' field for item id '$listItemID' to value '$subCategory'..." -NoNewLine

                    $materialSubCategorySetResult = Set-LookupFieldValue $listItem "QSLSubCategory" "QSLMaterialSubCategory" "Title" $subCategory "MaterialCategory" $category
                    
                    if($materialSubCategorySetResult) {
                        Write-Message "Done" -BackgroundColor Green
                    }
                    else {
                        Write-Message "Failed" -BackgroundColor Red
                    }
                }

                if($materialGroupSetResult -and $materialCategorySetResult -and $materialSubCategorySetResult) {
                    try {
                        Write-Message "`nCommitting updates..." -NoNewLine

                        $global:ctx.ExecuteQuery()

                        Write-Message "Done" -BackgroundColor Green
                    }
                    catch {
                        Write-Message $_ -ForegroundColor Red
                    }
                } 
                else { 
                    Write-Message "Failed. Either lookup values were not found in source list or some other error occurred." -BackgroundColor Red 
                }
            }
            "CascadingField2"{
                $supplierMatches = ($sourceField.ToString() | Select-String -pattern '{"SAP Supplier Name":"([^"\.]*)').Matches
                if($supplierMatches -ne $null) {
                    $supplier = $supplierMatches.Groups[1].Value

                    Write-Message "Setting 'QSLSupplier' field for item id '$listItemID' to value '$supplier'..." -NoNewLine

                    $supplierSetResult = Set-LookupFieldValue $listItem "QSLSupplier" "QSLSuppliers" "Title" $supplier $null $null
                    
                    if($supplierSetResult) {
                        Write-Message "Done" -BackgroundColor Green
                    }
                    else {
                        Write-Message "Failed" -BackgroundColor Red
                    }
                }

                $contactMatches = ($sourceField.ToString() | Select-String -pattern '{"Supplier Contact Name":"([^"]*)').Matches
                if($contactMatches -ne $null){
                    $contact = $contactMatches.Groups[1].Value

                    Write-Message "Setting 'QSLSupplierContact' field for item id '$listItemID' to value '$contact'..." -NoNewLine

                    $contactSetResult = Set-LookupFieldValue $listItem "QSLSupplierContact" "QSLSupplierContacts" "Title" $contact "Supplier" $supplier
                    
                    if($contactSetResult) {
                        Write-Message "Done" -BackgroundColor Green
                    }
                    else {
                        Write-Message "Failed" -BackgroundColor Red
                    }
                }

                if($supplierSetResult -and $contactSetResult ) {
                    try {
                        Write-Message "`nCommitting updates..." -NoNewLine

                        $global:ctx.ExecuteQuery()

                        Write-Message "Done" -BackgroundColor Green
                    }
                    catch {
                        Write-Message $_ -ForegroundColor Red
                    }
                } 
                else { 
                    
                    Write-Message "Failed. Either lookup values were not found in source list or some other error occurred." -BackgroundColor Red 
                }
            }
            default{}
        }
    }
    else {
        Write-Message "`nSource field '$sourceFieldName' = ''"
    }
}

function Process-AllItems($listName) {
    Write-Message "Retrieving all list items..." -NoNewLine

    $list = $global:ctx.Web.lists.GetByTitle($listName) 

    $query = New-Object Microsoft.SharePoint.Client.CamlQuery

    $query.ViewXml="<View Scope='RecursiveAll'><Query><Where><IsNotNull><FieldRef Name='ID' /></IsNotNull></Where></Query></View>"

    $listItems = $list.GetItems($query) 
    
    $global:ctx.Load($listItems) 
    $global:ctx.ExecuteQuery() 
    
    Write-Message "Done. Total Item Count: $($listItems.count)" -BackgroundColor Green

    if($listItems.count -gt 0) { 
        for($i = 0; $i -lt ($listItems.Count); $i++) {
            [int] $listItemID = $listItems[$i].ID

            Process-Item $listName $listItemID
        }          
    } 
}

#------------------ main script --------------------------------------------------

Write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

# Global Variables
$sCSOMPath = “C:\Users\prabhvx\OneDrive - Murphy Oil\SPLibs\” # Path to CSOM DLLs

#Load SharePoint CSOM Assemblies 
$sCSOMRuntimePath=$sCSOMPath + “Microsoft.SharePoint.Client.Runtime.dll”
$sCSOMPath=$sCSOMPath + “Microsoft.SharePoint.Client.dll”

Add-Type -Path $sCSOMPath
Add-Type -Path $sCSOMRuntimePath

#Get Credentials to connect 
$global:cred= Get-Credential -Message "Please enter your organizational credential for Office 365"
$global:credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:cred.Username, $global:cred.Password) 

$timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
{
    $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $logFolderPath = $currentDir + "\QSLMigrationLog\" + $timestamp + "\"

    Write-Host "You did not specify a path for the activity log. The log files will be available at '$logFolderPath'" -ForegroundColor Cyan

    if(-not (Test-Path $logFolderPath -PathType Container)) {
        Write-Host "`nCreating log folder '$logFolderPath'..." -NoNewline
        md -Path $logFolderPath | out-null
        Write-Host "Done" -ForegroundColor White -BackgroundColor Green
    }

    $global:logFolderPath = $logFolderPath
}
else {
    $global:logFolderPath = $outputLogFolderPath
}

$global:logFilePath = $global:logFolderPath + "DetailedActionLog.log"

Write-Host "The log file will be created at '$global:logFilePath'" -ForegroundColor Cyan

Write-Message "Script started ---> $(Get-Date)`n"

Write-Message "Connecting to site $siteURL..." -NoNewLine

#Setup the context 
$global:ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL) 
$global:ctx.Credentials = $global:credentials 

Write-Message "Done" -BackgroundColor Green

if($processAll) {
    Write-Message "`nProcessing ALL items in the list.."

    Process-AllItems $listName
}
else {
    for($i = 0; $i -lt ($listItemIDs.Count); $i++) {
        [int] $listItemID = $listItemIDs[$i]

        Process-Item $listName $listItemID
    }
}

Write-Message "Script ended ---> $(Get-Date)`n"

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow