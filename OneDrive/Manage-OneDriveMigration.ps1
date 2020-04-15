<#
 .NOTES
 ===========================================================================
 Created On:   9/15/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Manage-OneDriveMigration.ps1
 Version:      1.0.10
 Copyright:    Vinay Prabhugaonkar (Sparkhound Inc.)

 MIT License

 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the ""Software""), to deal
 in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

 THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 ===========================================================================
 
.SYNOPSIS
    Bulk or individual import/export from file share to to One Drive for Business in Office 365.

.DESCRIPTION  
    The script relies on SharePointPnPPowerShellOnline as well as Sharegate modules which must be installed on the machine where the script is run. 
    Script prompts for user credential on every run. It should be run by a user that has full control rights on the file share as well as SharePoint Administrator role for the Office 365 tenant.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing batch of one drives to import to. The csv schema contains following columns
    OneDriveUrl,SourceFolderPath,DestinationFolderName,FromDateString,ToDateString
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - destinationOneDriveUrl (string)
    Full url of user one drive
.PARAMETER - sourceFolderPath (string - comma separated)
    Full path to the file system folder being imported. Can be drive or mapped location or UNC path. Can be folder and/or file. Separate multiple by commas
.PARAMETER - destinationFolderName (string)
    Name of folder in which to import in destination one drive
.PARAMETER - createDestinationFolder (switch)
    Create destination folder if one does not exist? - default True - only relevant if destinationFolderName in provided.
.PARAMETER - onlyImportSourceChildren (switch)
    Should source folder be imported directly or only its children? - default True - only children of specified source folder
    are imported. False to import folder(s)/file(s) specified as comma separated list in 'sourceFolder' attribute
.PARAMETER - from (string)
    Date string in yyyy-MM-dd format. Specifies the modified date from filter for source item
.PARAMETER - to (string)
    Date string in yyyy-MM-dd format. Specifies the modified date to filter for source item. Leave empty for current date.
.PARAMETER - excludeExtensions (string array)
    File extensions to exclude from import. Default value pst. Leave empty if specifying limitToExtensions parameter.
.PARAMETER - limitToExtensions (string array)
    File extensions to include in import. All other extensions will be excluded. Default empty. Leave empty if specifying excludeExtensions parameter.
.PARAMETER - downloadFolderPath (string)
    Path where one drive files should be exported. Also the location where ShareGate migration reports will be exported.
.PARAMETER - migratePermissions (switch)
    Should permissions be migrated from source to destination? - default False
.PARAMETER - incrementalCopy (switch)
    Is this incremental migration to one drive? - default False
.PARAMETER - waitToExportReport (switch)
    Wait until O365 import job finishes before exporting migration job report? - default True
.PARAMETER - action (string)
    Import, Export. Default is empty string which means an import pre-check is performed to verify potential issues.

.USAGE 
    Bulk migrate to one drives specified in csv in a single batch operation
     
    PS >  Manage-OneDriveMigration.ps1 -inputCSVPath "c:/temp/onedrives.csv" -action "Import"
.USAGE 
    Bulk incremental copy to one drives specified in csv in a single batch operation
     
    PS >  Manage-OneDriveMigration.ps1 -inputCSVPath "c:/temp/onedrives.csv" -incrementalCopy -action "Import"
.USAGE 
    Perform a pre-check on one drives specified in a csv
     
    PS >  Manage-OneDriveMigration.ps1 -inputCSVPath "c:/temp/onedrives.csv"
.USAGE 
    Import to individual user one drive folder named 'UDrive' from a single source folder path 
     
    PS >  Manage-OneDriveMigration.ps1 -destinationOneDriveUrl "https://tenant-my.sharepoint.com/personal/xxx_domain_com/" -sourceFolderPath "\\share\myfolder or mapped drive:\myfolder" -destinationFolderName "UDrive" -action "Import"
.USAGE 
    Import to individual user one drive folder named 'UDrive' from a multiple source folders or files
     
    PS >  Manage-OneDriveMigration.ps1 -destinationOneDriveUrl "https://tenant-my.sharepoint.com/personal/xxx_domain_com/" -sourceFolderPath "\\share1\myfolder1,\\share2\myfolder2,\\share3\myfile.ext" -destinationFolderName "UDrive" -onlyImportSourceChildren $false -action "Import"
.USAGE 
    Import to individual user one drive root all items from a single source folder that were modified after Jan 01,2019
     
    PS >  Manage-OneDriveMigration.ps1 -destinationOneDriveUrl "https://tenant-my.sharepoint.com/personal/xxx_domain_com/" -sourceFolderPath "\\share1\myfolder1" -from "2019-01-01" -action "Import"
.USAGE 
    Import to individual user one drive root all items from a single source folder that were modified between Jan 01,2019 and Jan 31, 2019
     
    PS >  Manage-OneDriveMigration.ps1 -destinationOneDriveUrl "https://tenant-my.sharepoint.com/personal/xxx_domain_com/" -sourceFolderPath "\\share1\myfolder1" -from "2019-01-01" -to "2019-01-31" -action "Import"
.USAGE 
    Import to individual user one drive root all items from a single source folder that are of file type docx, txt or pdf
     
    PS >  Manage-OneDriveMigration.ps1 -destinationOneDriveUrl "https://tenant-my.sharepoint.com/personal/xxx_domain_com/" -sourceFolderPath "\\share1\myfolder1" -limitToExtensions @("txt","docx","pdf") -action "Import"
.USAGE 
    Export all the contents of a specific one drive to a specified location
     
    PS >  Manage-OneDriveMigration.ps1 -destinationOneDriveUrl "https://tenant-my.sharepoint.com/personal/xxx_domain_com/" -downloadFolderPath "drive:\\export-to-folder-name" -action "Export"
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on one drive urls and corresponding import paths, date filters and destination folder")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Full url to user one drive")]
    [string] $destinationOneDriveUrl,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Full path to the folder(s)/file(s) being imported. Separate multiple with comma.")]
    [string] $sourceFolderPath,

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Name of folder in which to import")]
    [string] $destinationFolderName,

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Create destination folder prior to import?")]
    [switch] $createDestinationFolder = $true,

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Only import source folder children? Should be FALSE if multiple source folder paths specified.")]
    [switch] $onlyImportSourceChildren = $true,

    [Parameter(Mandatory=$false, Position=7, HelpMessage="Migrate content with create/modify date FROM this date (format: yyyy-MM-dd)")]
    [string] $from,

    [Parameter(Mandatory=$false, Position=8, HelpMessage="Migrate content with create/modify date TO this date (format: yyyy-MM-dd)")]
    [string] $to,

    [Parameter(Mandatory=$false, Position=9, HelpMessage="File extension(s) to exclude. Leave empty if specifying limitToExtensions")]
    [string[]] $excludeExtensions=@("pst"),

    [Parameter(Mandatory=$false, Position=10, HelpMessage="File extension(s) to include. Leave empty if specifying excludeExtensions")]
    [string[]] $limitToExtensions,

    [Parameter(Mandatory=$false, Position=11, HelpMessage="Path to folder where files will be downloaded if action=export and/or export migration reports")]
    [string] $downloadFolderPath,

    [Parameter(Mandatory=$false, Position=12, HelpMessage="Migrate permissions during import?")]
    [switch] $migratePermissions=$false,

    [Parameter(Mandatory=$false, Position=13, HelpMessage="Incremental migration?")]
    [switch] $incrementalCopy=$false,
    
    [Parameter(Mandatory=$false, Position=14, HelpMessage="Wait until O365 import completes before exporting report?")]
    [switch] $waitToExportReport=$true,

    [Parameter(Mandatory=$false, Position=15, HelpMessage="Action to take")]
    [ValidateSet('','Import','Export')] [string] $action = ""
)

$ErrorActionPreference = "Continue"

function Create-DestinationFolder() {
    param(
        [Parameter(Mandatory=$true)][string]$siteUrl,
        [Parameter(Mandatory=$true)][string]$siteRelativeFolderPath
    )

    [boolean] $success = $false

    try {
        Write-Host "Connecting over PnP..." -NoNewline
        Write-Output "Connecting over PnP..." | out-file $global:logFilePath -NoNewline -Append

        Connect-PnPOnline -Url $siteUrl -Credentials $global:cred

        Write-Host "Done" -BackgroundColor Green
        Write-Output "Done" | out-file $global:logFilePath -Append

        Write-Host "Creating folder at '$siteRelativeFolderPath' if one does not exist..." -NoNewline
        Write-Output "Creating folder at '$siteRelativeFolderPath' if one does not exist..." | out-file $global:logFilePath -NoNewline -Append

        Resolve-PnPFolder -SiteRelativePath $siteRelativeFolderPath -ErrorAction Stop

        Write-host "Done" -BackgroundColor Green
        Write-Output "Done" | out-file $global:logFilePath -Append

        $success = $true

        Disconnect-PnPOnline
    } 
    catch { 
        Write-host "...Failed. Error: $($_.Exception.Message)" -BackgroundColor Red 
        Write-Output "...Failed. Error: $($_.Exception.Message)" | out-file $global:logFilePath -Append
    }

    return $success
}

function Create-PropertyTemplate() {
    param(
        [Parameter(Mandatory=$true)][PSCustomObject]$row
    )

    $optionalParams = @{}

    $from = $($row.FromDateString.Trim())
    if(![string]::IsNullOrEmpty($from)){
        $optionalParams.Add("From", $from)
    }

    $to = $($row.ToDateString.Trim())
    if(![string]::IsNullOrEmpty($to)){
        $optionalParams.Add("To", $to)
    }

    if($excludeExtensions -ne $null){
        $optionalParams.Add("ExcludeFileExtension", $excludeExtensions)
    }

    if($limitToExtensions -ne $null){
        $optionalParams.Add("LimitToFileExtension", $limitToExtensions)
    }

    if($migratePermissions) {
        $optionalParams.Add("Permissions", $true)
    }

    [Sharegate.Automation.Entities.PropertyTemplate] $sgPropertyTemplate = New-PropertyTemplate -VersionLimit 5 -NoLinkCorrection -WebParts -AuthorsAndTimestamps -VersionHistory @optionalParams
    
    return $sgPropertyTemplate
}

function Import-ToOneDrive() {
    param(
        [Parameter(Mandatory=$true)][PSObject] $oneDriveSite,
        [Parameter(Mandatory=$true)][PSCustomObject] $row,
        [Parameter(Mandatory=$false)][switch] $preCheck
    )

    Set-Variable destinationODList, msg, copyResult
    Clear-Variable destinationODList, msg, copyResult

    $sourceFolderPath = $($row.SourceFolderPath.Trim())
    $destinationFolderName = $($row.DestinationFolderName.Trim())

    if([string]::IsNullOrWhiteSpace($sourceFolderPath) -or -not (Test-Path $sourceFolderPath -PathType Container)) {
        Write-Host "Source folder path '$sourceFolderPath' does not exist...Skipping" -BackgroundColor Red
        Write-Output "Source folder path '$sourceFolderPath' does not exist...Skipping" | out-file $global:logFilePath -Append
        
        #log csv
        "$($row.OneDriveUrl) `t Yes `t $($row.SourceFolderPath) `t $($row.FromDateString) `t $($row.ToDateString) `t N/A `t N/A `t N/A `t N/A `t N/A `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append
        
        return 
    }

    # verify report download folder. Set to log folder if it doesn't exist
    if(-not (Test-Path $downloadFolderPath -PathType Container)) {
        $downloadFolderPath = $global:logFolderPath
    }

    $otherParams = @{}

    if($onlyImportSourceChildren){
        # default scenario - contents of specified folder are brought in. Helps reduce url length.
        $otherParams.Add("SourceFolder", $sourceFolderPath)
    }
    else{
        # the folder(s)/file(s) specified in parameter are imported
        $otherParams.Add("SourceFilePath", $sourceFolderPath)
    }

    if($incrementalCopy) {
        $copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
        $otherParams.Add("CopySettings", $copysettings)
        # insane mode incremental copy replaces all existing versions due to limitations of Azure Import API so switch to normal mode
        $otherParams.Add("NormalMode", $true)
    }
    else {
        # regular copy and replace will use insane mode
        $otherParams.Add("InsaneMode", $true)
    }

    if($waitToExportReport) {
        $otherParams.Add("WaitForImportCompletion", $true)
    }

    $destinationODList = Get-List -Name Documents -Site $oneDriveSite

    [Sharegate.Automation.Entities.PropertyTemplate] $propTemplate = Create-PropertyTemplate -row $row

    $importInODRoot = $true

    $msg = "`n$(Get-Date) - Importing '$sourceFolderPath' content, children only - '$onlyImportSourceChildren', from date - '$($row.FromDateString)', to date - '$($row.ToDateString)', excluding extensions - '$excludeExtensions', or limiting to extensions - '$limitToExtensions', into '$oneDriveUrl/Documents/$destinationFolderName'. Incremental copy? - '$incrementalCopy'. Insane mode copy? - '$(-not $incrementalCopy)'. This is a potentially long running operation. You can follow progress within ShareGate desktop 'Tasks' menu..."

    if(![string]::IsNullOrWhiteSpace($destinationFolderName)) {
        if($createDestinationFolder) {
            $importInODRoot = $false
            
            $folderUrl = "Documents/" + $destinationFolderName
            
            if(Create-DestinationFolder -siteUrl $oneDriveUrl -siteRelativeFolderPath $folderUrl) {
                Write-Host $msg -ForegroundColor Cyan -NoNewline
                Write-Output $msg | out-file $global:logFilePath -NoNewline -Append
        
                if(!$preCheck) {
                    $copyResult = Import-Document -Template $propTemplate -DestinationList $destinationODList -DestinationFolder $destinationFolderName @otherParams
                }
                else {
                    $copyResult = Import-Document -Template $propTemplate -DestinationList $destinationODList -DestinationFolder $destinationFolderName -WhatIf @otherParams
                }
            }
        }
    }

    if($importInODRoot) {
        Write-Host $msg -NoNewline -f Cyan
        Write-Output $msg | out-file $global:logFilePath -NoNewline -Append
                        
        if(!$preCheck) {
            $copyResult = Import-Document -Template $propTemplate -DestinationList $destinationODList @otherParams
        }
        else {
            $copyResult = Import-Document -Template $propTemplate -DestinationList $destinationODList -WhatIf @otherParams
        }
    }
    
    Write-Host "Done. $(Get-Date)" -BackgroundColor Green
    Write-Output "Done. $(Get-Date)" | out-file $global:logFilePath -Append

    Write-Host "`nSessionId: $($copyResult.SessionId) Result: $($copyResult.Result) ItemsCopied: $($copyResult.ItemsCopied)" -f Green
    Write-Host "Errors: $($copyResult.Errors) Warnings: $($copyResult.Warnings)" -f Red

    $copyResult | out-file $global:logFilePath -Append

    #log csv
    "$($row.OneDriveUrl) `t Yes `t $($row.SourceFolderPath) `t $($row.FromDateString) `t $($row.ToDateString) `t $($copyResult.SessionId) `t $($copyResult.Result) `t $($copyResult.ItemsCopied) `t $($copyResult.Errors) `t $($copyResult.Warnings) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append

    # export report
    Write-Host "`nExporting ShareGate migration report to '$downloadFolderPath$($copyResult.SessionId).xlsx'..." -NoNewline
    Write-Output "`nExporting ShareGate migration report to '$downloadFolderPath$($copyResult.SessionId).xlsx'.." | out-file $global:logFilePath -NoNewline -Append

    Export-Report -CopyResult $copyResult -Path $downloadFolderPath

    Write-Host "Done" -BackgroundColor Green
    Write-Output "Done" | out-file $global:logFilePath -Append
}

function Export-FromOneDrive() {
    param(
        [Parameter(Mandatory=$true)][PSObject] $oneDriveSite,
        [Parameter(Mandatory=$true)][string] $downloadFolderPath
    )

    if(-not (Test-Path $downloadFolderPath -PathType Container)) {
        Write-Host "Download folder path $downloadFolderPath does not exist...Skipping" -BackgroundColor Red
        Write-Output "Download folder path $downloadFolderPath does not exist...Skipping" | out-file $global:logFilePath -Append
                
        return 
    }

    Write-Host "Exporting from '$oneDriveUrl' into '$downloadFolderPath'. This is a potentially long running operation. You can follow progress within ShareGate desktop 'Tasks' menu..." -NoNewline
    Write-Output "Exporting from '$oneDriveUrl' into '$downloadFolderPath'. This is a potentially long running operation. You can follow progress within ShareGate desktop 'Tasks' menu..." | out-file $global:logFilePath -NoNewline -Append
                        
    Export-List -SourceSite $oneDriveSite -Name Documents -DestinationFolder $downloadFolderPath
    
    Write-Host "Done" -BackgroundColor Green
    Write-Output "Done" | out-file $global:logFilePath -Append
}

function Action-AOneDrive([PSCustomObject] $row) {
    $oneDriveUrl = $($row.OneDriveUrl.Trim())  
    #$sourceFolderPath = $($row.SourceFolderPath.Trim())
    #$destinationFolderName = $($row.DestinationFolderName.Trim())

    if([string]::IsNullOrWhiteSpace($oneDriveUrl)) {
        Write-Host "Destination OneDrive Url must be specified..." -NoNewline
        Write-Host "Skipping" -BackgroundColor Red

        return
    }

    Write-Host "`n--------------------------------------------------------------------------------------"
    Write-Host "Checking if one drive '$oneDriveUrl' exists..." -NoNewline
    Write-Output "Checking if one drive '$oneDriveUrl' exists..." | out-file $global:logFilePath -NoNewline -Append 

    Set-Variable destinationODSite

    Clear-Variable destinationODSite

    # check if one drive exists
    try {

        $destinationODSite = Connect-Site -Url $oneDriveUrl -Credential $global:cred -ErrorAction Stop

        Write-Host "...Yes" -BackgroundColor Green
        Write-Output "...Yes" | out-file $global:logFilePath -Append 
    }
    catch {

        Write-Host "...No. Skipping. Exception: $($Error[0].Exception.Message)" -BackgroundColor Red
        Write-Output "...No. Skipping. Exception: $($Error[0].Exception.Message)" | out-file $global:logFilePath -Append 

        #log csv
        "$($row.OneDriveUrl) `t No `t $($row.SourceFolderPath) `t $($row.FromDateString) `t $($row.ToDateString) `t N/A `t N/A `t N/A `t N/A `t N/A `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append

        return
    }

    #Add-SiteCollectionAdministrator -Site $destinationODSite

    switch($action)
    {
        'Import'{
            Import-ToOneDrive -oneDriveSite $destinationODSite -row $row
        }
        'Export'{
            Export-FromOneDrive -oneDriveSite $destinationODSite -downloadFolderPath $downloadFolderPath
        }
        default{
            Write-Host "`nYou are NOT taking any migration action on one drive '$oneDriveUrl'"
            Write-Host "THIS WILL RUN A PRE-CHECK AND IDENTIFY POTENTIAL ISSUES WITH IMPORT..." -BackgroundColor Magenta
            Write-Output "`nYou are NOT taking any migration action on one drive '$oneDriveUrl'. THIS WILL RUN A PRE-CHECK AND IDENTIFY POTENTIAL ISSUES WITH IMPORT..." | out-file $global:logFilePath -Append
            
            Import-ToOneDrive -oneDriveSite $destinationODSite -row $row -preCheck
        }
    }
                            
    #Remove-SiteCollectionAdministrator -Site $destinationODSite
}

function ProcessCSV([string] $csvPath)
{
    if(![string]::IsNullOrEmpty($csvPath))
    {
        Write-host "`nProcessing csv file $csvPath..." -ForegroundColor Green
        Write-Output "`nProcessing csv file $csvPath..." | out-file $global:logFilePath -Append

        $global:csv = Import-Csv -Path $csvPath
    }

    if($global:csv -ne $null)
    {
        $global:csv | % {
            Action-AOneDrive $_ | out-null
        }
    }
}

function CheckAndLoadRequiredModules() {
    $res = [PSCustomObject]@{SPOPnPModuleSuccess = $false;ShareGateModuleSuccess = $false}

    if(!(Get-Module -ListAvailable | ? {$_.Name -like "SharePointPnPPowerShellOnline"})) {
        Write-Host "Installing SharePointPnPPowerShellOnline module from https://www.powershellgallery.com/packages/SharePointPnPPowerShellOnline/3.13.1909.0..." -NoNewline

        Install-Module SharePointPnPPowerShellOnline -AllowClobber -Force

        Write-Host "Done" -BackgroundColor Green
    }

    try {
        Write-Host "Loading SharePointPnPPowerShellOnline module..." -ForegroundColor Cyan -NoNewline
        
        Import-Module SharePointPnPPowerShellOnline -DisableNameChecking -Force

        Write-Host "Done" -BackgroundColor Green

        $res.SPOPnPModuleSuccess = $true
    }
    catch {
        Write-Host "Failed" -BackgroundColor Red

        $res.SPOPnPModuleSuccess = $false
    }

    if(!(Get-Module -ListAvailable | ? {$_.Name -like "ShareGate"})) {
        Write-Host "ShareGate must be installed on this machine to avail of the ShareGate PowerShell module..."
    }
    else {
        try {
            Write-Host "Loading ShareGate module..." -ForegroundColor Cyan -NoNewline
        
            Import-Module ShareGate -Force

            Write-Host "Done" -BackgroundColor Green

            $res.ShareGateModuleSuccess = $true
        }
        catch {
            Write-Host "Failed" -BackgroundColor Red

            $res.ShareGateModuleSuccess = $false
        }
    }
    
    return $res
}

function CheckExecutionPolicy() {

    $res = [PSCustomObject]@{Success = $false}

    $currentUserPolicy = Get-ExecutionPolicy -Scope CurrentUser

    if($currentUserPolicy -eq "Unrestricted" -or $currentUserPolicy -eq "RemoteSigned") {
        Write-Host "Execution policy for current user is already set to $currentUserPolicy"
        
        $res.Success = $true
    }
    else {
        Write-Host "Setting CurrentUser execution policy to RemoteSigned. Check https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-6 for details about execution policy settings" -ForegroundColor Cyan -NoNewline
        
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

        Write-Host "...Done" -ForegroundColor White -BackgroundColor Green

        $res.Success = $true
    }

    return $res
}

#------------------ main script --------------------------------------------------

Write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

$psExecutionPolicyLevelCheck = CheckExecutionPolicy 

$modulesCheck = CheckAndLoadRequiredModules

if($psExecutionPolicyLevelCheck.Success -and $modulesCheck.SPOPnPModuleSuccess -and $modulesCheck.ShareGateModuleSuccess) {
    $global:cred = Get-Credential -Message "Please enter the organizational credential for Office 365 for an account that is SharePoint Administrator or Global Administrator"

    $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

    if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
    {
        $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $logFolderPath = $currentDir + "\ODBImportLog\" + $timestamp + "\"

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

    $global:logCSVPath = $global:logFolderPath + "QuickLog.csv"
    $global:logFilePath = $global:logFolderPath + "DetailedActionLog.log"

    if([string]::IsNullOrWhiteSpace($downloadFolderPath)) {
        Write-Host "You did not specify a path for downloading files or migration reports. The report files will be available at '$global:logFolderPath'" -ForegroundColor Cyan

        $downloadFolderPath = $global:logFolderPath
    }

    #log csv
    "OneDriveUrl `t OneDriveExists? `t SourcePath `t FromDate `t ToDate `t SGSessionId `t SGResult `t SGItemsCopied `t SGErrors `t SGWarnings `t RunDate `t Action" | out-file $global:logCSVPath

    Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

    if(![string]::IsNullOrWhiteSpace($inputCSVPath))
    {
        ProcessCSV $inputCSVPath
    }
    else 
    {
        Write-Host "`nYou did not specify a csv file containing input data...." -ForegroundColor Cyan

        $csvPathEntryReponse = Read-Host "Would you like to enter the full path of the csv file? [y|n]"
        if($csvPathEntryReponse -eq 'y') {
            do {
                $path = Read-Host "Enter full path to the csv file containing input data."
            }
            until (![string]::IsNullOrWhiteSpace($path))

            ProcessCSV $path
        }
        else {
            Write-Host "`nActioning single OneDrive..." -BackgroundColor White -ForegroundColor Black
            
            if([string]::IsNullOrWhiteSpace($destinationOneDriveUrl))
            {
                do {
                    $destinationOneDriveUrl = Read-Host "Specify the one drive destination Url"
                }
                until (![string]::IsNullOrWhiteSpace($destinationOneDriveUrl))
            }

            if([string]::IsNullOrWhiteSpace($sourceFolderPath)) {
                do {
                    $sourceFolderPath = Read-Host "Specify full path to the source folder"
                }
                until (![string]::IsNullOrWhiteSpace($sourceFolderPath))
            }

            $row = @{
                    OneDriveUrl=$destinationOneDriveUrl;
                    SourceFolderPath=$sourceFolderPath;
                    DestinationFolderName=$destinationFolderName;
                    FromDateString=$from;
                    ToDateString=$to;
                }
    
            Action-AOneDrive $row | Out-Null
        }
    }

    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow