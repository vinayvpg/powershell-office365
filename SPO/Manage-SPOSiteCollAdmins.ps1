[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing site collection urls")]
    [string] $inputCSVPath="C:\Users\prabhvx\OneDrive - Murphy Oil\Desktop\SP Management Scripts\SPO\malo365users.csv",
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Tenant admin url")]
    [string] $tenantAdminUrl = "https://murphyoil-admin.sharepoint.com",

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Site collection URLs, semicolon separated")]
    [string] $siteCollUrls="https://murphyoil-my.sharepoint.com/personal/Erick_Anthony_contractor_murphyoilcorp_com",

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Secondary admin users to set, semicolon separated")]
    [string] $siteCollAdminUsers="dxc_support@murphyoilcorp.com",

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Do it for all site collections?")]
    [switch] $all = $false,

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Action to take")]
    [ValidateSet('','Add','Remove')] [string] $action = "Add"
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
                Write-Host "---> Setting '$($adminUsersColl[$i])' as site collection administrator on $(Get-Date)..." -NoNewline
                Write-Output "---> Setting '$($adminUsersColl[$i])' as site collection administrator on $(Get-Date)..." | out-file $global:logFilePath -NoNewline -Append

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

function Remove-AdminUsers($siteColl, $adminUsers)
{
    $adminUsersColl = $adminUsers -split ";" 

    if($adminUsersColl)
    {
        for($i =0; $i -le ($adminUsersColl.count - 1) ; $i++)
        {
            try {
                Write-Host "---> Removing '$($adminUsersColl[$i])' as site collection administrator on $(Get-Date)..." -NoNewline
                Write-Output "---> Removing '$($adminUsersColl[$i])' as site collection administrator on $(Get-Date)..." | out-file $global:logFilePath -NoNewline -Append

                Remove-SPOUser -Site $($siteColl.Url) -LoginName $($adminUsersColl[$i])

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

function Action-SiteCollection([PSCustomObject] $row)
{   
    $urls = [string]::Empty
    $admins = [string]::Empty

    if($all) {
        $admins = $($row.SiteCollectionAdmins.Trim())
    
        if([string]::IsNullOrWhiteSpace($admins)) {
            Write-Host "Admin users are missing..." -NoNewline
            Write-Host "Quitting" -BackgroundColor Red

            return
        }

        $siteColls = Get-SPOSite -Limit All
    }
    else {
    
        $urls = $($row.SiteCollectionUrls.Trim())  
        $admins = $($row.SiteCollectionAdmins.Trim())
    
        if([string]::IsNullOrWhiteSpace($urls) -or [string]::IsNullOrWhiteSpace($admins)) {
            Write-Host "Either site collection url or admin user are missing..." -NoNewline
            Write-Host "Quitting" -BackgroundColor Red

            return
        }

        $siteColls = $urls -split ";" 
    }


    if($siteColls)
    {
        for($i =0; $i -le ($siteColls.count - 1) ; $i++)
        {
            Write-Host "`n--------------------------------------------------------------------------------------"
            Write-Output "`n--------------------------------------------------------------------------------------" | out-file $global:logFilePath -Append 

            Write-Host "Checking if a site collection with url '$($siteColls[$i])' exists..." -NoNewline
            Write-Output "Checking if a site collection with url '$($siteColls[$i])' exists..." | out-file $global:logFilePath -NoNewline -Append 

            $siteColl = $null
            $exists = $false

            try {
                # check if site exists
                $siteColl = Get-SPOSite $($siteColls[$i])
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
                'Add'{            
                    if($exists) {
                        Add-AdminUsers $siteColl $admins
                    }
                    else {
                        # populate log
                        "N/A `t $($siteColls[$i]) `t No `t N/A `t N/A `t N/A `t N/A `t N/A `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 

                        Write-Host "...Will NOT be added" -BackgroundColor Red
                        Write-Output "...Will NOT be added" | out-file $global:logFilePath -Append 
                    }
                }
                'Remove'{
                    if(!$exists) {
                        "N/A `t $($siteColls[$i]) `t No `t N/A `t N/A `t N/A `t N/A `t N/A `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 

                        Write-Host "No action taken" -BackgroundColor Green
                        Write-Output "No action taken" | out-file $global:logFilePath -Append 
                    }
                    else {
                        Remove-AdminUsers $siteColl $admins
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
            Action-SiteCollection $_ | out-null
        }
    }
}

#------------------ main script --------------------------------------------------

Write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow


    if([string]::IsNullOrWhiteSpace($tenantAdminUrl)) {
        do {
            $tenantAdminUrl = Read-Host "Specify the tenant admin url (https://tenantname-admin.sharepoint.com)"
        }
        until (![string]::IsNullOrWhiteSpace($tenantAdminUrl))
    }

    Connect-SPOService -Url $tenantAdminUrl -Credential (Get-Credential)

    Write-Host "Connected to $tenantAdminUrl..."

    $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

    if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
    {
        $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $logFolderPath = $currentDir + "\SiteCollAdminLog\" + $timestamp + "\"

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
    "SiteTitle `t Url `t Exists `t PrimaryOwner `t SecondaryAdmin `t Status `t StorageUsageCurrent `t LastContentModifiedDate `t ActionDate `t Action" | out-file $global:logCSVPath

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
            $row = @{}

            if($all) {
                if([string]::IsNullOrWhiteSpace($siteCollAdminUsers)) {
                    do {
                        $siteCollAdminUsers = Read-Host "Specify the email address(es) of the site collection admin user(s) (if multiple, separate with semicolon)"
                    }
                    until (![string]::IsNullOrWhiteSpace($siteCollAdminUsers))
                }
            
                $row = @{
                        SiteCollectionUrls=[string]::Empty;
                        SiteCollectionAdmins=$siteCollAdminUsers
                    }
            }
            else {
            
                Write-Host "`nActioning specific site collection(s)..." -BackgroundColor White -ForegroundColor Black
            
                if([string]::IsNullOrWhiteSpace($siteCollUrls)) {
                    do {
                        $siteCollUrls = Read-Host "Specify the url of the site collection(s) (if multiple, separate with semicolon)"
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
            }
    
            Action-SiteCollection $row | Out-Null
        }
    }

    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow