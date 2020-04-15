<#
 .NOTES
 ===========================================================================
 Created On:   9/1/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Manage-OneDriveDeletion.ps1
 Version:      1.0.0
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
    Bulk or individual restore or permanent removal of deleted OneDrive site(s) in Office 365.

.DESCRIPTION  
    The script relies on Microsoft SharePoint Online management shell which must be installed on the machine where the script is run. 
    Script prompts for user credential on every run. It should be run by a user who has Office 365 SharePoint Admin or Global Admin role for the tenant.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing urls of deleted one drives to restore/remove. The csv schema contains following columns
    SiteCollectionUrls,SiteCollectionAdmins
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - tenantAdminUrl (string)
    Url of tenant administration site https://tenantname-admin.sharepoint.com
.PARAMETER - siteCollUrls (string semicolon separated)
    Full url of one drives to restore. Separate multiple urls with semicolon
.PARAMETER - siteCollAdminUsers (string semicolon separated)
    Email address of user that should be made administrator on restored one drive. Separate multiple with semicolon.
.PARAMETER - action (string)
    Restore or Remove. Default is empty string which means no action.

.USAGE 
    Bulk restore deleted one drives specified in csv in a single batch operation
     
    PS >  Manage-OneDriveDeletion.ps1 -inputCSVPath "c:/temp/onedrivestorestore.csv" -action "Restore"
.USAGE 
    Restore individual deleted one drive
     
    PS >  Manage-OneDriveDeletion.ps1 -siteCollUrls 'https://tenantname-my.sharepoint.com/personal/userupnsuffix_userupndomain' -siteCollAdminUsers 'admin1@domain.com;admin2@domain.com' -action "Restore"
.USAGE 
    Restore multiple deleted one drives without batch operation
     
    PS >  Manage-OneDriveDeletion.ps1 -siteCollUrls 'https://tenantname-my.sharepoint.com/personal/usr1_domain_com;https://tenantname-my.sharepoint.com/personal/usr2_domain_com' -siteCollAdminUsers 'admin1@domain.com;admin2@domain.com' -action "Restore"
.USAGE 
    Permanently remove individual one drive
     
    PS >  Manage-OneDriveDeletion.ps1 -siteCollUrls 'https://tenantname-my.sharepoint.com/personal/userupnsuffix_userupndomain' -action "Remove"
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing one drive urls to restore")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Tenant admin url")]
    [string] $tenantAdminUrl,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Onedrive site URLs, semicolon separated")]
    [string] $siteCollUrls,

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Set admin users, semicolon separated")]
    [string] $siteCollAdminUsers,

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Action to take")]
    [ValidateSet('','Restore','Remove')] [string] $action = ""
)

$ErrorActionPreference = "Continue"

function Add-AdminUsers($siteColl, $adminUsers)
{
    $adminUsersColl = $adminUsers -split ";" 

    if($adminUsersColl)
    {
        for($i =0; $i -le ($adminUsersColl.count - 1) ; $i++)
        {
            try {
                Write-Host "---> Setting '$($adminUsersColl[$i])' as one drive administrator on $(Get-Date)..." -NoNewline
                Write-Output "---> Setting '$($adminUsersColl[$i])' as one drive administrator on $(Get-Date)..." | out-file $global:logFilePath -NoNewline -Append

                Set-SPOUser -Site $($siteColl.Url) -LoginName $($adminUsersColl[$i]) -IsSiteCollectionAdmin $true 

                "$($siteColl.Title) `t $($siteColl.Url) `t Yes `t $($siteColl.Owner) `t $($adminUsersColl[$i]) `t $($siteColl.Status) `t $($siteColl.StorageUsageCurrent) `t $($siteColl.LastContentModifiedDate) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append

                Write-Host "Done. $(Get-Date)" -BackgroundColor Green
                Write-Output "Done. $(Get-Date)" | out-file $global:logFilePath -Append
            }
            catch {
                "$($siteColl.Title) `t $($siteColl.Url) `t Yes `t $($siteColl.Owner) `t N/A `t $($siteColl.Status) `t $($siteColl.StorageUsageCurrent) `t $($siteColl.LastContentModifiedDate) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append

                Write-Host "Failed. $(Get-Date)" -BackgroundColor Red
                Write-Output "Failed. $(Get-Date)" | out-file $global:logFilePath -Append

                $_ | out-file $global:logFilePath -Append
            }
        }
    }
}


function Action-OneDrive([PSCustomObject] $row)
{       
    $urls = $($row.SiteCollectionUrls.Trim())  
    $admins = $($row.SiteCollectionAdmins.Trim())
    
    if([string]::IsNullOrWhiteSpace($urls)) {
        Write-Host "Either one drive url(s) are missing..." -NoNewline
        Write-Host "Quitting" -BackgroundColor Red

        return
    }

    $siteColls = $urls -split ";" 
    
    if($siteColls)
    {
        for($i =0; $i -le ($siteColls.count - 1) ; $i++)
        {
            Write-Host "`n--------------------------------------------------------------------------------------"
            Write-Output "`n--------------------------------------------------------------------------------------" | out-file $global:logFilePath -Append 

            Write-Host "Checking if deleted one drive with url '$($siteColls[$i])' is still available. Deleted one drive is only available for restore upto 93 days after initial deletion..." -NoNewline
            Write-Output "Checking if deleted one drive with url '$($siteColls[$i])' is still available. Deleted one drive is only available for restore upto 93 days after initial deletion..." | out-file $global:logFilePath -NoNewline -Append 

            $siteColl = $null
            $exists = $false

            try {
                # check if deleted one drive exists
                $siteColl = $global:allDeletedOneDrives | ? { $_.Url -like $($siteColls[$i]) }
            }
            catch{
                $_ | out-file $global:logFilePath -NoNewline -Append 
            }

            if($siteColl -eq $null) {
                Write-Host "No" -BackgroundColor Red
                Write-Output "No" | out-file $global:logFilePath -Append
            }
            else {
                $exists = $true

                Write-Host "Yes" -BackgroundColor Green
                Write-Output "Yes" | out-file $global:logFilePath -Append 
            }

            switch($action)
            {
                'Restore'{ 
                    if(![string]::IsNullOrWhiteSpace($admins)) {
                        if($exists) {                        
                            try {
                                Write-Host "Restoring deleted one drive with url '$($siteColls[$i])'..." -NoNewline
                                Write-Output "Restoring deleted one drive with url '$($siteColls[$i])'..." | out-file $global:logFilePath -NoNewline -Append 

                                Restore-SPODeletedSite -Identity $($siteColl.Url) -ErrorAction Stop

                                Write-Host "Done" -BackgroundColor Green
                                Write-Output "Done" | out-file $global:logFilePath -Append
                            
                                # populate log
                                "$($siteColl.SiteId) `t $($siteColl.Url) `t $($siteColl.Status) `t Restored `t $($siteColl.DeletionTime) `t $($siteColl.DaysRemaining) `t $($siteColl.StorageQuota) `t $($siteColl.ResourceQuota) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 
 
                                Add-AdminUsers $siteColl $admins
                            }
                            catch {
                        
                                Write-Host "Failed. $(Get-Date)" -BackgroundColor Red
                                Write-Output "Failed. $(Get-Date)" | out-file $global:logFilePath -Append

                                # populate log
                                "$($siteColl.SiteId) `t $($siteColl.Url) `t $($siteColl.Status) `t NotRestored `t $($siteColl.DeletionTime) `t $($siteColl.DaysRemaining) `t $($siteColl.StorageQuota) `t $($siteColl.ResourceQuota) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 

                                $_ | out-file $global:logFilePath -Append
                            }
                        }
                        else {
                            # populate log
                            "N/A `t $($siteColls[$i]) `t NotAvailable `t NotRestored `t N/A `t N/A `t N/A `t N/A `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 

                            Write-Host "...CANNOT be restored" -BackgroundColor Red
                            Write-Output "...CANNOT be restored" | out-file $global:logFilePath -Append 
                        }
                    }
                    else {
                        # populate log
                        "N/A `t $($siteColls[$i]) `t NotAvailable `t NotRestored `t N/A `t N/A `t N/A `t N/A `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 

                        Write-Host "...No admin user specified. CANNOT be restored" -BackgroundColor Red
                        Write-Output "...No admin user specified. CANNOT be restored" | out-file $global:logFilePath -Append 
                    }          
                }
                'Remove'{            
                    if($exists) {                        
                        try {
                            Write-Host "Permanently deleting one drive with url '$($siteColls[$i])'. All content will be permanently removed..." -NoNewline
                            Write-Output "Permanently deleting one drive with url '$($siteColls[$i])'. All content will be permanently removed.." | out-file $global:logFilePath -NoNewline -Append 

                            Remove-SPODeletedSite -Identity $($siteColl.Url) -ErrorAction Stop -Confirm

                            Write-Host "Done" -BackgroundColor Green
                            Write-Output "Done" | out-file $global:logFilePath -Append
                            
                            # populate log
                            "$($siteColl.SiteId) `t $($siteColl.Url) `t $($siteColl.Status) `t Removed `t $($siteColl.DeletionTime) `t $($siteColl.DaysRemaining) `t $($siteColl.StorageQuota) `t $($siteColl.ResourceQuota) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 
 
                            Add-AdminUsers $siteColl $admins
                        }
                        catch {
                        
                            Write-Host "Failed. $(Get-Date)" -BackgroundColor Red
                            Write-Output "Failed. $(Get-Date)" | out-file $global:logFilePath -Append

                            # populate log
                            "$($siteColl.SiteId) `t $($siteColl.Url) `t $($siteColl.Status) `t NotRestored `t $($siteColl.DeletionTime) `t $($siteColl.DaysRemaining) `t $($siteColl.StorageQuota) `t $($siteColl.ResourceQuota) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 

                            $_ | out-file $global:logFilePath -Append
                        }
                    }
                    else {
                        # populate log
                        "N/A `t $($siteColls[$i]) `t NotAvailable `t NotRemoved `t N/A `t N/A `t N/A `t N/A `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 

                        Write-Host "...CANNOT be removed" -BackgroundColor Red
                        Write-Output "...CANNOT be removed" | out-file $global:logFilePath -Append 
                    }
                }
                default{
                    Write-Host "`nYou are NOT taking any action on '$($siteColls[$i])'..."
                    Write-Output "`nYou are NOT taking any action on '$($siteColls[$i])'..." | out-file $global:logFilePath -Append
                }
            }
        }
    }
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
            Action-OneDrive $_ | out-null
        }
    }
}

function CheckAndLoadRequiredModules() {
    $res = [PSCustomObject]@{SPOModuleSuccess = $false}

    if(!(Get-Module -ListAvailable | ? {$_.Name -like "Microsoft.Online.SharePoint.PowerShell"})) {
        Write-Host "Installing Microsoft.Online.SharePoint.PowerShell module from https://www.powershellgallery.com/packages/Microsoft.Online.SharePoint.PowerShell/16.0.19223.12000..." -NoNewline

        Install-Module Microsoft.Online.SharePoint.PowerShell -AllowClobber -Force

        Write-Host "Done" -BackgroundColor Green
    }

    try {
        Write-Host "Loading Microsoft.Online.SharePoint.PowerShell module..." -ForegroundColor Cyan -NoNewline
        
        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -Force

        Write-Host "Done" -BackgroundColor Green

        $res.SPOModuleSuccess = $true
    }
    catch {
        Write-Host "Failed" -BackgroundColor Red

        $res.SPOModuleSuccess = $false
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

if($psExecutionPolicyLevelCheck.Success -and $modulesCheck.SPOModuleSuccess) {
    if([string]::IsNullOrWhiteSpace($tenantAdminUrl)) {
        do {
            $tenantAdminUrl = Read-Host "Specify the tenant admin url (https://tenantname-admin.sharepoint.com)"
        }
        until (![string]::IsNullOrWhiteSpace($tenantAdminUrl))
    }

    Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

    Connect-SPOService -Url $tenantAdminUrl -Credential (Get-Credential)

    Write-Host "Connected to $tenantAdminUrl..." -ForegroundColor Cyan

    $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

    if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
    {
        $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $logFolderPath = $currentDir + "\OneDriveRestoreLog\" + $timestamp + "\"

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

    #log csv
    "SiteId `t Url `t RecycleStatus `t RestoreStatus `t DeletionTime `t DaysRemaining `t StorageQuota `t ResourceQuota `t ActionDate `t Action" | out-file $global:logCSVPath

    $global:allDeletedOneDrives = Get-SPODeletedSite -Limit All -IncludeOnlyPersonalSite

    if($global:allDeletedOneDrives -ne $null) {
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
                $row = @{}
            
                Write-Host "`nActioning specific one drive(s)..." -BackgroundColor White -ForegroundColor Black
            
                if([string]::IsNullOrWhiteSpace($siteCollUrls)) {
                    do {
                        $siteCollUrls = Read-Host "Specify the url of the one drive site(s) (if multiple, separate with semicolon)"
                    }
                    until (![string]::IsNullOrWhiteSpace($siteCollUrls))
                }

                if([string]::IsNullOrWhiteSpace($siteCollAdminUsers)) {
                    do {
                        $siteCollAdminUsers = Read-Host "Specify the email address(es) of the site collection admin user(s) (if multiple, separate with semicolon)"
                    }
                    until (![string]::IsNullOrWhiteSpace($siteCollAdminUsers))
                }
            
                $row = @{
                        SiteCollectionUrls=$siteCollUrls;
                        SiteCollectionAdmins=$siteCollAdminUsers
                    }
            
                Action-OneDrive $row | Out-Null
            }
        }
    }
    else {
        Write-Host "No deleted one drives found..." -NoNewline
        Write-Host "Quitting" -BackgroundColor Red
    }

    Disconnect-SPOService

    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow