<#  
 .NOTES
 ===========================================================================
 Created On:   12/1/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Update-UnifiedGroup.ps1
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
    Bulk or individual update of a Office 365 Groups (Unified Group).

.Description  
    Bulk or individual update of a Office 365 Groups (Unified Group). Script prompts for user credential on every run.
    It should be run by a user who has Office 365 group creation rights in Azure AD.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing batch of Groups to update. The csv schema contains following columns - GroupId,GroupDisplayName
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - groupDisplayName (string < 256 characters required)
    Display name of the Group.
.PARAMETER - groupId (string - guid format)
    Id of the group. Either id or display name must be specified
.PARAMETER - hiddenFromAddressListsEnabled (boolean)
    
.PARAMETER - hiddenFromExchangeClientsEnabled (boolean)
    
.PARAMETER - action (string)
    Update. Default is empty string which means no action.

.USAGE 
    Bulk update Groups specified in csv in a single batch operation
     
    PS >  Update-UnifiedGroup.ps1 -inputCSVPath "c:/temp/O365groups.csv" -action "Update"
.USAGE 
    Update individual Group with specific parameters
     
    PS >  Update-UnifiedGroup.ps1 -groupDisplayName "Company function Group" -hiddenFromExchangeClientsEnabled $false -action "Update"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on Office 365 Groups to be update")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Group display name")]
    [string] $groupDisplayName,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="O365 Group Id")]
    [string] $groupId,

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Group setting")]
    [boolean] $hiddenFromAddressListsEnabled,

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Group setting")]
    [boolean] $hiddenFromExchangeClientsEnabled,

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Group setting")]
    [boolean] $unifiedGroupWelcomeMessageEnabled,

    [Parameter(Mandatory=$false, Position=7, HelpMessage="Action to take")]
    [ValidateSet('','Update')] [string] $action = ""
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

function Update-AUnifiedGroup([PSCustomObject] $row)
{
    $groupDisplayName = $($row.GroupDisplayName.Trim())
    $groupId = $($row.GroupId.Trim())

    $optionalParams = @{}

    if($row.HiddenFromExchangeClientsEnabled -ne $null) {
        $optionalParams.Add("HiddenFromExchangeClientsEnabled", [System.Convert]::ToBoolean($row.HiddenFromExchangeClientsEnabled))
    }

    if($row.HiddenFromAddressListsEnabled -ne $null) {
        $optionalParams.Add("HiddenFromAddressListsEnabled", [System.Convert]::ToBoolean($row.HiddenFromAddressListsEnabled))
    }

    if($row.UnifiedGroupWelcomeMessageEnabled -ne $null) {
        $optionalParams.Add("UnifiedGroupWelcomeMessageEnabled", [System.Convert]::ToBoolean($row.UnifiedGroupWelcomeMessageEnabled))
    }

    if(![string]::IsNullOrWhiteSpace($groupId)) {
        # update based on group Id

        try {

            Write-Message "You are updating a group id '$groupId' or named '$groupDisplayName'..." -NoNewline

            Set-UnifiedGroup -Identity $groupId @optionalParams -ErrorAction Stop

            Write-Message "Done" -BackgroundColor Green
            
            Get-UnifiedGroup -Identity $groupId -IncludeAllProperties | fl

            # populate log
            #"$teamDisplayName `t $visibility `t $teamGroupId `t $owners `t $members `t $channels `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 
        }
        catch {
            Write-Message "Failed" -BackgroundColor Red
        }
    }
}

function Action-AUnifiedGroup([PSCustomObject] $row)
{     
    $groupId = $($row.GroupId.Trim()) 
    $groupDisplayName = $($row.GroupDisplayName.Trim())

    if([string]::IsNullOrWhiteSpace($groupId) -and [string]::IsNullOrWhiteSpace($groupDisplayName)) {
        Write-Message "Either group Id or group display name must be specified..." -NoNewline
        Write-Message "Quitting" -BackgroundColor Red

        return
    }

    Write-Host "`n--------------------------------------------------------------------------------------"
    Write-Message "Checking if a group with name '$groupDisplayName' or id '$groupId' exists..." -NoNewline

    $group = $null
    $exists = $false

    # check if Group exists
    if(![string]::IsNullOrWhiteSpace($groupId)) {
        $group = Get-UnifiedGroup -Identity $groupId
    }
    else {
        $group = Get-UnifiedGroup -Identity $groupDisplayName
    }

    if($group -eq $null) {
        Write-Message "...No..." -BackgroundColor Red
    }
    else {
        $exists = $true

        Write-Message "...Yes. GroupId: $($group.ExternalDirectoryObjectId)..." -BackgroundColor Green

        if([string]::IsNullOrEmpty($groupId)) {
            # team discovered from display name, update group id to pass on
            Write-Message "Updating GroupId of '$groupDisplayName'..." -NoNewline

            $row.GroupId = $($group.ExternalDirectoryObjectId)

            Write-Message "Done" -BackgroundColor Green
        }
    }

    switch($action)
    {
        'Update'{            
            if($exists) {
                Update-AUnifiedGroup $row
            }
            else {
                Write-Message "Will NOT be updated" -BackgroundColor Red
            }
        }
        default{
            Write-Message "`nYou are NOT taking any action on a group named '$groupDisplayName'..."
        }
    }
}

function ProcessCSV([string] $csvPath)
{
    if(![string]::IsNullOrEmpty($csvPath))
    {
        Write-Message "`nProcessing csv file $csvPath..." -ForegroundColor Green

        $global:csv = Import-Csv -Path $csvPath
    }

    if($global:csv -ne $null)
    {
        $global:csv | % {
            Action-AUnifiedGroup $_ | out-null
        }
    }
}

function CheckAndLoadRequiredModules($cred) {
    $res = [PSCustomObject]@{Success = $false}

    try {
        Write-Host "Loading Exchange Online Powershell module..." -ForegroundColor Cyan -NoNewline
        
        # Connect to Exchange Online
        $global:exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri “https://outlook.office365.com/powershell-liveid/” -Credential $cred -Authentication “Basic” –AllowRedirection

        Import-PSSession $global:exchangeSession -AllowClobber

        Write-Host "Done" -BackgroundColor Green

        $res.Success = $true
    }
    catch{
        Write-Host "Failed" -BackgroundColor Red
    }
    
    return $res
}

function CheckPowerShellVersion()
{
    $res = [PSCustomObject]@{ Success = $false }

    $h = Get-Host

    if($h.Version.ToString() -like "5.1*") {

         Write-Host "PowerShell version check succeeded" -ForegroundColor Cyan

         $res.Success = $true
    }
    else {
        Write-Host "PowerShell version must be 5.1 or higher. Please download Windows Management Framework which contains PowerShell 5.1 from https://www.microsoft.com/en-us/download/details.aspx?id=54616" -ForegroundColor Cyan

        $res.Success = $false
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

$psVersionCheck = CheckPowerShellVersion

$psExecutionPolicyLevelCheck = CheckExecutionPolicy

if($psVersionCheck.Success -and $psExecutionPolicyLevelCheck.Success) {
    $global:cred = Get-Credential
    $global:exchangeSession = $null

    $modulesCheck = CheckAndLoadRequiredModules $global:cred

    if($modulesCheck.Success) {
        $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

        if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
        {
            $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
            $logFolderPath = $currentDir + "\O365GroupsLog\" + $timestamp + "\"

            Write-Host "You did not specify a path for the activity log. The log files will be available at '$logFolderPath'" -ForegroundColor Cyan

            if(-not (Test-Path $logFolderPath -PathType Container)) {
                Write-Host "Creating log folder '$logFolderPath'..." -NoNewline
                md -Path $logFolderPath | out-null
                Write-Host "Done" -ForegroundColor White -BackgroundColor Green
            }

            $global:logFolderPath = $logFolderPath
        }
        else {

            $global:logFolderPath = $outputLogFolderPath
        }

        #$global:logCSVPath = $global:logFolderPath + "UpdateLog.csv"
        $global:logFilePath = $global:logFolderPath + "ActionLog.log"

        Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

        #log csv
        #"GroupName `t Type `t GroupId `t Owners `t Members `t Alias `t Date `t Action" | out-file $global:logCSVPath

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
                Write-Host "`nActioning single group..." -BackgroundColor White -ForegroundColor Black

                if([string]::IsNullOrWhiteSpace($groupId)) {
                    if([string]::IsNullOrWhiteSpace($groupDisplayName))
                    {
                        do {
                            $groupDisplayName = Read-Host "Specify the display name for the group"
                        }
                        until (![string]::IsNullOrWhiteSpace($groupDisplayName))
                    }
                }

                $row = @{
                        GroupId=$groupid;
                        GroupDisplayName=$groupDisplayName;
                        HiddenFromAddressListsEnabled=$hiddenFromAddressListsEnabled;
                        HiddenFromExchangeClientsEnabled=$hiddenFromExchangeClientsEnabled;
                        UnifiedGroupWelcomeMessageEnabled=$unifiedGroupWelcomeMessageEnabled
                    }
    
                Action-AUnifiedGroup $row | Out-Null
            }
        }
    }

    if($global:exchangeSession -ne $null){
        Remove-PSSession -Session $global:exchangeSession
    }
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow