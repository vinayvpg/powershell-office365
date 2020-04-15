<#
 .NOTES
 ===========================================================================
 Created On:   9/15/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Purge-SPOSiteCollRecycleBin.ps1
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
 ============================================================================
 
.SYNOPSIS
    Purge first/second stage recycle bin of a site collection.

.DESCRIPTION  
    The script relies on SharePointOnlinePnPPowerShell which must be installed on the machine where the script is run. 
    Script prompts for user credential on every run. It should be run by a user who has Office 365 SharePoint Admin or Global Admin role for the tenant.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing site collection urls whose recycle bin needs to be purged. The csv schema contains following columns - SiteCollectionUrl
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - siteCollectionUrl (string)
    Url of site collection
.PARAMETER - stage (string)
    Empty of second. Default is empty string which means purge both first and second stage recycle bins.

.USAGE 
    Bulk purge all recycle bins of multiple site collections in a single batch operation
     
    PS >  Purge-SPOSiteCollRecycleBin.ps1 -inputCSVPath "c:/temp/sitecollurls.csv"
.USAGE 
    Purge only the second stage recycle bin for a specific site collection
     
    PS >  Purge-SPOSiteCollRecycleBin.ps1 -siteCollectionUrl "https://xxx.sharepoint.com/sites/site" -stage "Second"
.USAGE 
    Purge both recycle bins for a specific site collection
     
    PS >  Purge-SPOSiteCollRecycleBin.ps1 -siteCollectionUrl "https://xxx.sharepoint.com/sites/site"
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing one drive urls to restore")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Url of site collection whose recycle bin needs to be purged")]
    [string] $siteCollectionUrl="https://murphyoil-my.sharepoint.com/personal/stan_stanbrook_murphyoilcorp_com",

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Which recycle bin to purge")]
    [ValidateSet('','Second')] [string] $stage = ""
)

$ErrorActionPreference = "Continue"

function Purge-SiteCollection([PSCustomObject] $row)
{       
    $url = $($row.SiteCollectionUrl.Trim())
    
    if([string]::IsNullOrWhiteSpace($url)) {
        Write-Host "Site collection url must be specified..." -NoNewline
        Write-Host "Skipping" -BackgroundColor Red

        return
    }
    
    Write-Host "`n--------------------------------------------------------------------------------------"
    Write-Output "`n--------------------------------------------------------------------------------------" | out-file $global:logFilePath -Append 

    Write-Host "Connecting to '$url'..." -NoNewline

    Connect-PnPOnline -Url $url -Credentials (Get-Credential)

    Write-Host "Done" -f Green

    try {
        switch($stage)
        {
            'Second'{    
                Write-Host "Purging '$stage' stage recycle bin for '$url'...." -NoNewline
                Write-Output "Purging '$stage' recycle bin for '$url'...." | out-file $global:logFilePath -NoNewline -Append 

                Clear-PnPRecycleBinItem -SecondStageOnly -Force -ErrorAction Stop
                
                Write-Host "Done" -BackgroundColor Green
                Write-Output "Done" | out-file $global:logFilePath -Append     
            }
            default{
                Write-Host "Purging ALL recycle bins for '$url'...." -NoNewline
                Write-Output "Purging ALL recycle bins for '$url'...." | out-file $global:logFilePath -NoNewline -Append 

                Clear-PnPRecycleBinItem -All -Force -ErrorAction Stop

                Write-Host "Done" -BackgroundColor Green
                Write-Output "Done" | out-file $global:logFilePath -Append
            }
        }
    }
    catch {
        Write-Host "Failed" -BackgroundColor Red
        Write-Output "Failed" | out-file $global:logFilePath -Append 

        $_ | out-file $global:logFilePath -Append
    }
    
    # populate log
    "$url `t $(Get-Date) `t $stage" | out-file $global:logCSVPath -Append        
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
            Purge-SiteCollection $_ | out-null
        }
    }
}

function CheckAndLoadRequiredModules() {
    $res = [PSCustomObject]@{SPOPnPModuleSuccess = $false}

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

if($psExecutionPolicyLevelCheck.Success -and $modulesCheck.SPOPnPModuleSuccess) {

    $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

    if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
    {
        $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $logFolderPath = $currentDir + "\RecycleBinPurgeLog\" + $timestamp + "\"

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

    Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

    #log csv
    "Url `t ActionDate `t Stage" | out-file $global:logCSVPath

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
            
            Write-Host "`nActioning single site collection..." -BackgroundColor White -ForegroundColor Black
            
            if([string]::IsNullOrWhiteSpace($siteCollectionUrl)) {
                do {
                    $siteCollectionUrl = Read-Host "Specify the url of the site collection that needs to be purged"
                }
                until (![string]::IsNullOrWhiteSpace($siteCollectionUrl))
            }
            
            $row = @{
                    SiteCollectionUrl=$siteCollectionUrl
                }
            
            Purge-SiteCollection $row | Out-Null
        }
    }
 
    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow