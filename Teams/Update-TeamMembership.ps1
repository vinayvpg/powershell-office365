<#
 .NOTES
 ===========================================================================
 Created On:   8/1/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Update-TeamMembership.ps1
 Version:      1.0.4
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
    Bulk or individual membership update of a Microsoft Team in Office 365.

.DESCRIPTION  
    The script relies on MicrosoftTeams as well as SkypeOnline Connector modules which must be installed on the machine where the script is run. 
    Script prompts for user credential on every run. It should be run by a user who has Office 365 group/team creation rights for the tenant in Azure AD.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing batch of Teams to update. The csv schema contains following columns
    TeamGroupId,TeamDisplayName,OwnerUPNs,MemberUPNs
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - groupId (string - guid format)
    If Team is being provisioned from existing O365 group, then the Id of the group
.PARAMETER - teamDisplayName (string < 256 characters)
    Display name of the Team. Either display name or groupId are required
.PARAMETER - allTeams (switch)
    Perform the action (add/remove) on ALL teams in the tenant. The input csv and individual team id/display name parameters will be ignored
.PARAMETER - owners (string)
    Semicolon separated list of UPNs.
.PARAMETER - members (string)
    Semicolon separated list of UPNs.
.PARAMETER - action (string)
    Add, Remove. Default is empty string which means no action.

.USAGE 
    Bulk update Teams membership specified in csv in a single batch operation
     
    PS >  Update-TeamMembership.ps1 -inputCSVPath "c:/temp/teams.csv" -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -action "Add"
.USAGE 
    Update individual Team from existing O365 group
     
    PS >  Update-TeamMembership.ps1 -groupId 'xxxxxxx-xxxx-xxxx-xxxxxxxx' -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -action "Remove"
.USAGE 
    Update individual Team with specific parameters
     
    PS >  Update-TeamMembership.ps1 -teamDisplayName 'My team' -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -action "Add"
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on Teams to be updated")]
    [string] $inputCSVPath="C:\Users\prabhvx\OneDrive - Murphy Oil\Desktop\SP Management Scripts\Book1.csv",
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="O365 Group Id. If this is given then team display name parameter given here will be ignored and taken from the group")]
    [string] $groupId,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Team display name")]
    [string] $teamDisplayName="Delete - Test Team 2",

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Perform action on all Teams?")]
    [switch] $allTeams = $false,

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Owners (semicolon separated list of UPNs)")]
    [string] $owners="vinay_prabhugaonkar@murphyoilcorp.com",

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Members (semicolon separated list of UPNs)")]
    [string] $members,

    [Parameter(Mandatory=$false, Position=7, HelpMessage="Action to take")]
    [ValidateSet('','Add','Remove')] [string] $action = "Add"
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

function Add-Users
{   
    param(   
             $users,$groupId,$role
          )   
    Process
    {
        $teamusers = $users -split ";" 
        if($teamusers)
        {
            for($j =0; $j -le ($teamusers.count - 1) ; $j++)
            {
                $user = $teamusers[$j].Trim()

                Write-Message "---> Adding '$user' to $groupId as $role ..." -NoNewline

                try {
                    
                    Add-TeamUser -GroupId $groupId -User $user -Role $role -Verbose

                    Write-Message "Done" -BackgroundColor Green
                }
                catch {
                    Write-Message "Failed. Check that user UPN is valid and the user is not already present in this role." -BackgroundColor Red
                }
            }
        }
    }
}

function Remove-Users
{   
    param(   
             $users,$groupId,$role
          )   
    Process
    {
        $teamusers = $users -split ";" 
        
        if($teamusers)
        {
            for($j =0; $j -le ($teamusers.count - 1) ; $j++)
            {
                $user = $teamusers[$j].Trim()

                Write-Message "---> Removing '$user' from '$groupId' as '$role' ..." -NoNewline

                try {
                    if($role  -eq "Owner") {
                        # first demote to member
                        MicrosoftTeams\Remove-TeamUser -GroupId $groupId -User $user -Role $role -Verbose
                    }

                    MicrosoftTeams\Remove-TeamUser -GroupId $groupId -User $user

                    Write-Message "Done" -BackgroundColor Green
                }
                catch {
                    Write-Message "Failed. Check that the user UPN is valid and is already present in this role." -BackgroundColor Red
                }
            }
        }
    }
}

function Action-AllTeams(){
    $teams = Get-Team

    Write-Message "Found $($teams.count) Teams in the tenant..." -ForegroundColor Magenta

    for($j = 0; $j -le ($teams.count - 1); $j++){
        $team = $teams[$j]
        $row = @{
                TeamGroupId=$($team.GroupId);
                TeamDisplayName=$($team.DisplayName);
                OwnerUPNs=$owners;
                MemberUPNs=$members
            }
    
        Action-ATeam $row | Out-Null
    }
}

function Action-ATeam([PSCustomObject] $row)
{   
    $teamGroupId = $($row.TeamGroupId.Trim())  
    $teamDisplayName = $($row.TeamDisplayName.Trim())

    $owners = $($row.OwnerUPNs.Trim())
    $members = $($row.MemberUPNs.Trim())

    if([string]::IsNullOrWhiteSpace($teamGroupId) -and [string]::IsNullOrWhiteSpace($teamDisplayName)) {
        Write-Message "Either group Id or team display name must be specified..." -NoNewline
        Write-Message "Quitting" -BackgroundColor Red

        return
    }

    Write-Host "`n--------------------------------------------------------------------------------------"
    Write-Message "Checking if a team with name '$teamDisplayName' or id '$teamGroupId' exists..." -NoNewline

    $team = $null
    $exists = $false

    # check if Team with the same name/id already exists
    if(![string]::IsNullOrWhiteSpace($teamGroupId)) {
        $team = Get-Team -GroupId $teamGroupId
    }
    else {
        $team = Get-Team -DisplayName $teamDisplayName
    }

    if($team -eq $null) {
        Write-Message "No" -BackgroundColor Red
    }
    else {
        $exists = $true

        Write-Message "...Yes. GroupId: $($team.GroupId)..." -BackgroundColor Green

        if([string]::IsNullOrEmpty($teamGroupId)) {
            # team discovered from display name, update group id to pass on
            Write-Message "Updating TeamGroupId of team '$teamDisplayName'..." -NoNewline

            $row.TeamGroupId = $($team.GroupId)

            $teamGroupId = $($team.GroupId)

            Write-Message "Done" -BackgroundColor Green
        }
    }

    switch($action)
    {
        'Add'{            
            if($exists) {
                if(![string]::IsNullOrEmpty($owners)){
                    Add-Users -users $owners -groupId $teamGroupId -role Owner
                }

                if(![string]::IsNullOrEmpty($members)){
                    Add-Users -users $members -groupId $teamGroupId -role Member 
                }
            }
            else {
                Write-Message "Will NOT be updated" -BackgroundColor Red
            }
        }
        'Remove'{            
            if($exists) {
                if(![string]::IsNullOrEmpty($owners)){
                    Remove-Users -users $owners -groupId $teamGroupId -role Owner
                }

                if(![string]::IsNullOrEmpty($members)){
                    Remove-Users -users $members -groupId $teamGroupId -role Member 
                }
            }
            else {
                Write-Message "Will NOT be updated" -BackgroundColor Red
            }
        }
        default{
            Write-Message "`nYou are NOT taking any action on a team named '$teamDisplayName'..."
        }
    }

    "$teamDisplayName `t $teamGroupId `t $owners `t $members `t $(Get-Date) `t $action" | out-file $global:logCSVPath
}

function ProcessCSV([string] $csvPath)
{
    if(![string]::IsNullOrEmpty($csvPath))
    {
        Write-message "`nProcessing csv file $csvPath..." -ForegroundColor Green

        $global:csv = Import-Csv -Path $csvPath
    }

    if($global:csv -ne $null)
    {
        $global:csv | % {
            Action-ATeam $_ | out-null
        }
    }
}

function CheckAndLoadRequiredModules() {
    $res = [PSCustomObject]@{TeamsModuleSuccess = $false;SkypeModuleSuccess = $false}

    if(!(Get-InstalledModule -Name MicrosoftTeams)) {
        Write-Host "Installing MicrosoftTeams module from https://www.powershellgallery.com/packages/MicrosoftTeams/1.0.1..." -NoNewline

        Install-Module MicrosoftTeams -Force

        Write-Host "Done" -BackgroundColor Green
    }

    try {
        Write-Host "Loading MicrosoftTeams module..." -ForegroundColor Cyan -NoNewline
        
        Import-Module MicrosoftTeams -Force

        Write-Host "Done" -BackgroundColor Green

        $res.TeamsModuleSuccess = $true
    }
    catch{
        Write-Host "Failed" -BackgroundColor Red

        $res.TeamsModuleSuccess = $false
    }

    try {
        Write-Host "Loading Skype for Business Online Powershell module..." -ForegroundColor Cyan -NoNewline

        Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1"

        Write-Host "Done" -BackgroundColor Green

        $res.SkypeModuleSuccess = $true
    }
    catch {

        Write-Host "Skype for Business Online Powershell module is not installed. This module is required for managing meeting policies and configurations for Teams. Install from https://www.microsoft.com/en-us/download/details.aspx?id=39366" -BackgroundColor Red

        $res.SkypeModuleSuccess = $false
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

$modulesCheck = CheckAndLoadRequiredModules

if($psVersionCheck.Success -and $psExecutionPolicyLevelCheck.Success -and $modulesCheck.TeamsModuleSuccess -and $modulesCheck.SkypeModuleSuccess) {
    $global:cred = Get-Credential -Message "Please enter your organizational credential for Office 365"
    $global:currentUser = $global:cred.UserName

    # Connect Microsoft Team
    Connect-MicrosoftTeams -Credential $global:cred

    $timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

    if([string]::IsNullOrWhiteSpace($outputLogFolderPath))
    {
        $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        $logFolderPath = $currentDir + "\TeamsLog\" + $timestamp + "\"

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

    $global:logCSVPath = $global:logFolderPath + "ProvisioningLog.csv"
    $global:logFilePath = $global:logFolderPath + "ActionLog.log"

    #log csv
    "TeamName `t GroupId `t Owners `t Members `t Date `t Action" | out-file $global:logCSVPath

    Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

    if($allTeams) {
        Write-Host "`nActioning ALL Teams in the tenant..." -BackgroundColor White -ForegroundColor Black

        if([string]::IsNullOrWhiteSpace($owners) -and [string]::IsNullOrWhiteSpace($members)) {
            Write-Message "Either owners or members to be added/removed must be specified..." -NoNewline
            Write-Message "Quitting" -BackgroundColor Red

            return
        }

        Action-AllTeams
    }
    else {
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
                Write-Host "`nActioning single team..." -BackgroundColor White -ForegroundColor Black
            
                if([string]::IsNullOrWhiteSpace($groupId)) {
                    if([string]::IsNullOrWhiteSpace($teamDisplayName))
                    {
                        do {
                            $teamDisplayName = Read-Host "Specify the display name for the Team"
                        }
                        until (![string]::IsNullOrWhiteSpace($teamDisplayName))
                    }
                }

                $row = @{
                        TeamGroupId=$groupId;
                        TeamDisplayName=$teamDisplayName;
                        OwnerUPNs=$owners;
                        MemberUPNs=$members
                    }
    
                Action-ATeam $row | Out-Null
            }
        }
    }

    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append

    Disconnect-MicrosoftTeams
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow