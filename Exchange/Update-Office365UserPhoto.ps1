<#
.SYNOPSIS
    Bulk or individual cloning of a Microsoft Team in Office 365 using Microsoft Graph API.

.NOTES
 ===========================================================================
 Created On:   11/18/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Update-Office365UserPhoto.ps1
 Version:      1.0.1
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
 
.DESCRIPTION  
    The script relies on Azure AD App that exposes Graph API permissions. Group.ReadWrite.All, User.Read.All, Files.ReadWrite.All and Notes.ReadWrite.All 
    delegated permissions to Graph API must be enabled on this app. It should be run by a user who has Office 365 group/team creation rights for the tenant in Azure AD.         

.PARAMETER - inputCSVPath (string)
    Path to csv file containing batch of Teams to provision. The csv schema contains following columns
    TeamDisplayName,TeamDescription,TeamPrivacy,TemplateTeamName,OwnerUPNs,MemberUPNs
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - teamDisplayName (string < 256 characters)
    Display name of the Team. Required.
.PARAMETER - teamDescription (string)
    Description setting for the Team. Default is empty.
.PARAMETER - templateTeamName
    Display name of the Template Team to clone. Required.
.PARAMETER - tenantRootUrl
    Tenant root url https://tenantname.sharepoint.com. Required.
.PARAMETER - tenantId
    Tenant Id from Azure Enterprise App Registration. Required.
.PARAMETER - clientId
    Client Id from Azure Enterprise App Registration. Required.
.PARAMETER - clientSecret
    Client secret from Azure Enterprise App Registration. Required.
.PARAMETER - redirectUri
    Redirect URI from Azure Enterprise App Registration. Required.
.PARAMETER - teamPrivacy (string)
    Visibility setting for the Team. Default is private.
.PARAMETER - owners (string)
    Semicolon separated list of UPNs.
.PARAMETER - members (string)
    Semicolon separated list of UPNs.
.PARAMETER - removeWikiTab (switch)
    If any wiki tabs are cloned and not needed then remove them
.PARAMETER - removeCreatorFromTeam (switch)
    Script runner becomes automatic owner of team. Remove with this switch.
.PARAMETER - action (string)
    Clone. Default is empty string which means no action.

.USAGE 
    Bulk provision Teams specified in csv in a single batch operation
     
    PS >  Clone-TeamWithGraphAPI.ps1 -inputCSVPath "c:/temp/teams.csv" -action "Clone"
.USAGE 
    Clone individual Team with specific parameters. Keep script runner as owner (default)
     
    PS >  Clone-TeamWithGraphAPI.ps1 -teamDisplayName "My team" -templateTeamName "My Template Team Name" -teamPrivacy "Private" -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -action "Clone"
.USAGE 
    Clone individual Team with specific parameters. Remove script runner from Team ownership
     
    PS >  Clone-TeamWithGraphAPI.ps1 -teamDisplayName "My team" -templateTeamName "My Template Team Name" -teamPrivacy "Private" -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -action "Clone" -removeCreatorFromTeam
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on users whose photo to update")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Office 365 admin credential - Exchange or Global Admin role")]
    [string] $adminUserUPN = "vinay@office365ninja.onmicrosoft.com",

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Office 365 admin credential - Exchange or Global Admin role")]
    [string] $adminUserPassword = "Montecarlo1#",

    [Parameter(Mandatory=$false, Position=4, HelpMessage="UPN of user whose photo to update")]
    [string] $userUPN = "vinay@office365ninja.onmicrosoft.com",

    [Parameter(Mandatory=$false, Position=4, HelpMessage="UPN of user whose photo to update")]
    [string] $employeeId = "12345",

    [Parameter(Mandatory=$false, Position=5, HelpMessage="UNC or mapped drive path to the photos directory")]
    [string] $photosDirectoryPath = "c:\photos",

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Do it for all users?")]
    [switch] $allUsers = $true,

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
            Action-AUser $_ | out-null
        }
    }
}

function Action-AUser([PSCustomObject] $row)
{
    $userUPN = $($row.UPN.Trim())
    $employeeId = $($row.EmployeeId.Trim())

    Write-Message "`n-----------------------------------------------------------------------------------------"
    Write-Message "Processing user '$userUPN' with employeeId '$employeeId'..."

    if([string]::IsNullOrWhiteSpace($userUPN) -or [string]::IsNullOrWhiteSpace($employeeId)) {
        Write-Message "User UPN and EmployeeId must be specified..." -NoNewLine 
        Write-Message "Quitting" -BackgroundColor Red

        return
    }

    $photoFilePath = $photosDirectoryPath + "\" + $employeeId + ".jpg"

    if(-not (Test-Path $photoFilePath)) {
        Write-Message "File '$photoFilePath' does not exist..." -NoNewLine
        Write-Message "Quitting" -BackgroundColor Red

        return
    }

    try {
        Write-Message "Retrieving file '$photoFilePath' for user '$userUPN'..." -NoNewLine

        $pictureData = Get-Item $photoFilePath -ErrorAction Stop

        Write-Message "Done" -BackgroundColor Green
    }
    catch{
        Write-Message "Failed to retrieve file. Exception: $($_.Exception.Message)"
    }

    try {
        Write-Message "Setting profile photo for user '$userUPN'..." -NoNewLine

        Set-UserPhoto -Identity $userUPN -PictureData ([System.IO.File]::ReadAllBytes($pictureData)) -Confirm:$false -ErrorAction Stop

        Write-Message "Done" -BackgroundColor Green
    }
    catch{
        Write-Message "Failed to set profile photo. Exception: $($_.Exception.Message)"
    }
}

#------------------ main script --------------------------------------------------

Write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

$timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
{
    $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $logFolderPath = $currentDir + "\PhotoLog\" + $timestamp + "\"

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
$global:logFilePath = $global:logFolderPath + "ActionLog.log"

if([string]::IsNullOrWhiteSpace($photosDirectoryPath))
{
    do {
        $photosDirectoryPath = Read-Host "Specify the path to the directory (UNC or mapped drive) that holds user photos"
    }
    until (![string]::IsNullOrWhiteSpace($photosDirectoryPath))
}

if(-not (Test-Path $photosDirectoryPath)) {
    Write-Message "`nPath '$photosDirectoryPath' is invalid..." -NoNewLine
    Write-Message "Quitting" -BackgroundColor Red

    return
}

if([string]::IsNullOrWhiteSpace($adminUserUPN))
{
    do {
        $adminUserUPN = Read-Host "Specify the UPN of the admin user (should have User Management or Exchange admin or Global Admin role)"
    }
    until (![string]::IsNullOrWhiteSpace($adminUserUPN))
}

if([string]::IsNullOrWhiteSpace($adminUserPassword))
{
    do {
        $adminUserPassword = Read-Host "Specify the password of the admin user (should have User Management or Exchange admin or Global Admin role)"
    }
    until (![string]::IsNullOrWhiteSpace($adminUserPassword))
}

$pwd = $adminUserPassword | ConvertTo-SecureString -AsPlainText -Force

#$credential = Get-Credential -Message "Please enter your organizational tenant admin credential for Office 365"

$credential = New-Object System.Management.Automation.PsCredential($adminUserUPN,$pwd)

Write-Message "Connecting to Exchange Online..." -ForegroundColor Cyan -NoNewLine

# Connect to Exchange Online
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri “https://outlook.office365.com/powershell-liveid/” -Credential $credential -Authentication “Basic” –AllowRedirection
Import-PSSession $exchangeSession -AllowClobber

Write-Message "Done" -BackgroundColor Green

#log csv
"UserName `t UPN `t EmployeeNumber `t PhotoFilePath `t Date `t Action" | out-file $global:logCSVPath

Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

if(![string]::IsNullOrWhiteSpace($inputCSVPath))
{
    ProcessCSV $inputCSVPath
}
else 
{
    $users = New-Object System.Collections.ArrayList

    if($allUsers) {
        Write-Message "`nActioning all users..." -BackgroundColor White -ForegroundColor Black

        $obj = Get-Mailbox -ResultSize unlimited #-Filter { (RecipientType -eq "UserMailbox") -and (IsValid -eq $true) -and (IsResource -eq $false) -and (IsLinked -eq $false) -and (IsShared -eq $false) -and (AccountDisabled -eq $false)}

        if($obj.GetType().BaseType.Name -eq "Array") {
            $users.AddRange($obj)
        }

        Write-Message “`nFound $($users.Count) users to process” -ForegroundColor Cyan

        for($i = 0; $i -le ($users.Count - 1); $i++) {            
            $row = @{
                UPN=$($users[$i].UserPrincipalName);
                EmployeeId=$($users[$i].CustomAttribute1)
            }

            Action-AUser $row | Out-Null
        }
    }
    else {
        Write-Message "`nActioning single user..." -BackgroundColor White -ForegroundColor Black
            
        if([string]::IsNullOrWhiteSpace($userUPN))
        {
            do {
                $userUPN = Read-Host "Specify the UPN of the user whose photo is to be updated"
            }
            until (![string]::IsNullOrWhiteSpace($userUPN))
        }

        if([string]::IsNullOrWhiteSpace($employeeId))
        {
            do {
                $employeeId = Read-Host "Specify the EmployeeId of the user whose photo is to be updated"
            }
            until (![string]::IsNullOrWhiteSpace($employeeId))
        }

        $obj = Get-Mailbox -ResultSize unlimited | ? { $_.UserPrincipalName -eq $userUPN }

        $row = @{
                UPN=$userUPN;
                EmployeeId=$employeeId
            }
    
        Action-AUser $row | Out-Null
    }
}

Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append

Write-Message "Disconnect Exchange Online session..." -ForegroundColor Cyan -NoNewLine

# Disconnect exchange online session
Remove-PSSession $exchangeSession

Write-Message "Done" -BackgroundColor Green

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow