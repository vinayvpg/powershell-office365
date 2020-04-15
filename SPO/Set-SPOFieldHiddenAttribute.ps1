[CmdletBinding()]
param(
    
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=1, HelpMessage="URL of the site where the QSL list resides")]
    [string] $siteURL="https://murphyoil.sharepoint.com/GlobalProcurement/purchasing",

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Display name of the QSL list")]
    [string] $listName="International QSL/Procurement Plan Database",

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Comma separated field internal names whose hidden attribute is to be toggled")]
    [string[]] $fieldNames=@("CascadingField1","SubCategory","Material","Proposed","Copyoffield1","GL","Copyoffield2","ISNCompanyID","SupplierContactEmail0","SupplierContactName0","SupplierContactPhone0","SAPSupplierContactEmail","SAPSupplierNumber","SAPSupplierName1")
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

function Set-FieldLinkHiddenAttribute($ctype, $fieldLink, $fieldName){
    $success = $false
    
    if($fieldLink -ne $null) { 
        Write-Message "`nSource field link attributes '$fieldName'"

        Write-Message $fieldLink

        Write-Message "`nSetting Hidden=TRUE..." -NoNewLine
        
        $fieldLink.Hidden = $true
        $fieldLink.ShowInDisplayForm = $false

        $ctype.Update($false)

        $success = $true

        Write-Message "Done" -BackgroundColor Green
    }

    return $success
}

function Set-FieldHiddenAttribute($fieldName, $field){
    $success = $false
    
    if($field -ne $null) { 
        Write-Message "`nSource field schema '$fieldName' = $($field.SchemaXml)"

        Write-Message "`nSetting Hidden='TRUE'..." -NoNewLine
        
        $newSchema = $field.SchemaXml.Replace('Hidden="FALSE"','Hidden="TRUE"')

        $field.SchemaXml = $newSchema

        $field.Update()

        $success = $true

        Write-Message "Done" -BackgroundColor Green
    }

    return $success
}

function Process-ListField($list, $fieldName) {
    $field = $list.Fields.GetByInternalNameOrTitle($fieldName)

    $global:ctx.Load($field) 
    $global:ctx.ExecuteQuery()

    if($field -ne $null) {
        Write-Message "Done" -BackgroundColor Green

        $res = Set-FieldHiddenAttribute $fieldName $field
        
        if($res) {
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
                    
            Write-Message "Failed to update field schema" -BackgroundColor Red 
        }
    }
    else {
        Write-Message "Field not found on list" -BackgroundColor Red
    }
}


function Process-Field($listName, $fieldName) {
    Write-Message "`n-----------------------------------------------------------------------------------"
    Write-Message "Processing field '$fieldName' on list '$listName'..."

    $list = $global:ctx.Web.lists.GetByTitle($listName)
    $listCTypes = $list.ContentTypes

    Write-Message "Retrieving all content types for list '$listName'..." -NoNewLine

    $global:ctx.Load($listCTypes) 
    $global:ctx.ExecuteQuery()

    Write-Message "Done" -BackgroundColor Green

    $itemCType = $listCTypes | ? {$_.Name -eq "Item"}

    Write-Message "Retrieving 'Item' content type field link named '$fieldName'..." -NoNewLine

    $ctypeFieldLinks = $itemCType.FieldLinks

    $global:ctx.Load($ctypeFieldLinks) 
    $global:ctx.ExecuteQuery()

    $fieldLink = $ctypeFieldLinks | ? {$_.Name -eq $fieldName}

    if($fieldLink -ne $null) {
        Write-Message "Done" -BackgroundColor Green

        $res = Set-FieldLinkHiddenAttribute $itemCType $fieldLink $fieldName
        
        if($res) {
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
                    
            Write-Message "Failed to update field link" -BackgroundColor Red 
        }

        Write-Message "FieldLink hidden, now retrieving field created in the list..." -NoNewLine

        Process-ListField $list $fieldName
    }
    else {

        Write-Message "Not found" -BackgroundColor Red

        Write-Message "Field not part of content type. Trying to retrieve list field..." -NoNewLine

        Process-ListField $list $fieldName

        <#
        $field = $list.Fields.GetByInternalNameOrTitle($fieldName)

        $global:ctx.Load($field) 
        $global:ctx.ExecuteQuery()

        if($field -ne $null) {
            Write-Message "Done" -BackgroundColor Green

            $res = Set-FieldHiddenAttribute $fieldName $field
        
            if($res) {
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
                    
                Write-Message "Failed to update field schema" -BackgroundColor Red 
            }
        }
        else {
            Write-Message "Field not found on list" -BackgroundColor Red
        }
        #>
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

for($i = 0; $i -lt ($fieldNames.Count); $i++) {
    $fieldName = $fieldNames[$i]

    Process-Field $listName $fieldName
}

Write-Message "Script ended ---> $(Get-Date)`n"

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow