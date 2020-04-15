<#
 .NOTES
 ===========================================================================
 Created On:   9/15/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Provision-OneDriveWithSG.ps1
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
    Bulk or individual provisioning of one drive for licensed users using Sharegate PowerShell.

.DESCRIPTION  
    The script relies on ShareGate powershell module which must be installed on the machine where the script is run. 
    Script prompts for user credential on every run. It should be run by a user who has Office 365 SharePoint Admin or Global Admin role for the tenant.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing UPN of users whose one drive is to be provisioned. The csv schema contains following columns - UPN
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - tenantAdminUrl (string)
    Url of tenant administration site https://tenantname-admin.sharepoint.com
.PARAMETER - UPN (string)
    UPN of the user whoes one drive is to be provisioned
.PARAMETER - action (string)
    Provision. Default is empty string which means no action.

.USAGE 
    Bulk provision one drives in a single batch operation
     
    PS >  Provision-OneDriveWithSG.ps1 -inputCSVPath "c:/temp/onedrivestoprovision.csv" -action "Provision"
.USAGE 
    Provision a single user's one drive
     
    PS >  Provision-OneDriveWithSG.ps1 -UPN 'useremail@domain.com' -action "Provision"
.USAGE 
    Check if one drive exists for all users in a batch
     
    PS >  Provision-OneDriveWithSG.ps1 -inputCSVPath "c:/temp/onedrivestoprovision.csv"
.USAGE 
    Check if a single user's one drive exists
     
    PS >  Provision-OneDriveWithSG.ps1 -UPN 'useremail@domain.com'
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing one drive urls to restore")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Tenant admin url")]
    [string] $tenantAdminUrl,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="User UPN")]
    [string] $UPN,

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Action to take")]
    [ValidateSet('','Provision')] [string] $action = ""
)

$ErrorActionPreference = "Continue"

function Action-OneDrive([PSCustomObject] $row)
{       
    $userUPN = $($row.UPN.Trim())
    
    if([string]::IsNullOrWhiteSpace($userUPN)) {
        Write-Host "UPN of the user must be specified..." -NoNewline
        Write-Host "Quitting" -BackgroundColor Red

        return
    }
    
    Write-Host "`n--------------------------------------------------------------------------------------"
    Write-Output "`n--------------------------------------------------------------------------------------" | out-file $global:logFilePath -Append 

    $od = $null

    try {
        switch($action)
        {
            'Provision'{    
                Write-Host "Provisioning one drive for '$userUPN' if one doesn't exist. This is a potentially long running operation and can take upto 24 hrs. The command does not wait for provisioning to finish...." -NoNewline
                Write-Output "Provisioning one drive for '$userUPN' if one doesn't exist. This is a potentially long running operation and can take upto 24 hrs. The command does not wait for provisioning to finish...." | out-file $global:logFilePath -NoNewline -Append 

                $od = Get-OneDriveUrl -Tenant $global:tenant -Email $userUPN -ProvisionIfRequired -DoNotWaitForProvisioning -ErrorAction Stop

                Write-Host "Done" -BackgroundColor Green
                Write-Output "Done" | out-file $global:logFilePath -Append     
            }
            default{
                Write-Host "Checking if one drive has been provisioned for '$userUPN'..." -NoNewline
                Write-Output "Checking if one drive has been provisioned for '$userUPN'..." | out-file $global:logFilePath -NoNewline -Append 

                $od = Get-OneDriveUrl -Tenant $global:tenant -Email $userUPN -ErrorAction Stop

                Write-Host "Yes" -BackgroundColor Green
                Write-Output "Yes" | out-file $global:logFilePath -Append
            }
        }
    }
    catch {
        Write-Host "Failed" -BackgroundColor Red
        Write-Output "Failed" | out-file $global:logFilePath -Append 

        $_ | out-file $global:logFilePath -Append
    }
    
    # populate log
    "$userUPN `t $od `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append        
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
    $res = [PSCustomObject]@{ShareGateModuleSuccess = $false}

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

if($psExecutionPolicyLevelCheck.Success -and $modulesCheck.ShareGateModuleSuccess) {

    $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

    if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
    {
        $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $logFolderPath = $currentDir + "\ODProvisionLog\" + $timestamp + "\"

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
    
    if([string]::IsNullOrWhiteSpace($tenantAdminUrl)) {
        do {
            $tenantAdminUrl = Read-Host "Specify the tenant admin url (https://tenantname-admin.sharepoint.com)"
        }
        until (![string]::IsNullOrWhiteSpace($tenantAdminUrl))
    }

    $global:tenant = Connect-Site -Url $tenantAdminUrl -Browser

    Write-Host "Connected to $tenantAdminUrl with ShareGate PowerShell..." -ForegroundColor Cyan

    #log csv
    "UPN `t OneDriveUrl `t ActionDate `t Action" | out-file $global:logCSVPath

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
            
            Write-Host "`nActioning single user..." -BackgroundColor White -ForegroundColor Black
            
            if([string]::IsNullOrWhiteSpace($UPN)) {
                do {
                    $UPN = Read-Host "Specify the UPN of the user whose one drive needs to be provisioned"
                }
                until (![string]::IsNullOrWhiteSpace($UPN))
            }
            
            $row = @{
                    UPN=$UPN
                }
            
            Action-OneDrive $row | Out-Null
        }
    }
 
    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow