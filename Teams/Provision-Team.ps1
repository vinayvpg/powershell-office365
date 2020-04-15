<#
 .NOTES
 ===========================================================================
 Created On:   8/1/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Provision-Team.ps1
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
    Bulk or individual provisioning/deprovisioning of a Microsoft Team in Office 365.

.DESCRIPTION  
    The script relies on MicrosoftTeams as well as SkypeOnline Connector modules which must be installed on the machine where the script is run. 
    Script prompts for user credential on every run. It should be run by a user who has Office 365 group/team creation rights for the tenant in Azure AD.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing batch of Teams to provision. The csv schema contains following columns
    TeamGroupId,TeamDisplayName,TeamDescription,MailNickName,TeamPrivacy,Channels,OwnerUPNs,MemberUPNs,AllowGiphy,GiphyContentRating,AllowStickersAndMemes,AllowCustomMemes,AllowGuestCreateUpdateChannels,AllowGuestDeleteChannels,AllowAddRemoveApps,AllowCreateUpdateRemoveTabs,AllowCreateUpdateRemoveConnectors,AllowUserEditMessages,AllowUserDeleteMessages,AllowOwnerDeleteMessages,AllowTeamMentions,AllowChannelMentions,ShowInTeamsSearchAndSuggestions
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - groupId (string - guid format)
    If Team is being provisioned from existing O365 group, then the Id of the group
.PARAMETER - teamDisplayName (string < 256 characters)
    Display name of the Team. Either display name or groupId are required
.PARAMETER - teamPrivacy (string)
    Visibility setting for the Team. Default is private.
.PARAMETER - owners (string)
    Semicolon separated list of UPNs.
.PARAMETER - members (string)
    Semicolon separated list of UPNs.
.PARAMETER - channels (string)
    Semicolon separated list of channels to provision.
.PARAMETER - action (string)
    Provision, deprovision. Default is empty string which means no action.

.USAGE 
    Bulk provision Teams specified in csv in a single batch operation
     
    PS >  Provision-Team.ps1 -inputCSVPath "c:/temp/teams.csv" -action "Provision"
.USAGE 
    Provision individual Team from existing O365 group
     
    PS >  Provision-Team.ps1 -groupId 'xxxxxxx-xxxx-xxxx-xxxxxxxx' -action "Provision"
.USAGE 
    Provision individual Team with specific parameters
     
    PS >  Provision-Team.ps1 -teamDisplayName 'My team' -teamPrivacy "Public" -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -channels "channel 1;channel 2" -action "Provision"
.USAGE 
    Deprovision individual Team using group Id
     
    PS >  Provision-Team.ps1 -groupId 'xxxxxxx-xxxx-xxxx-xxxxxxxx' -action "Deprovision"
.USAGE 
    Deprovision individual Team using team display name
     
    PS >  Provision-Team.ps1 -teamDisplayName 'My team' -action "Deprovision"
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on Teams to be provisioned or deprovisioned")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="O365 Group Id. If this is given then team display name/description/visibility parameters given here will be ignored and taken from the group")]
    [string] $groupId,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Team display name")]
    [string] $teamDisplayName,

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Team privacy setting")]
    [ValidateSet('Private','Public','HiddenMembership')][string] $teamPrivacy="Private",

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Owners (semicolon separated list of UPNs)")]
    [string] $owners,

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Members (semicolon separated list of UPNs)")]
    [string] $members,

    [Parameter(Mandatory=$false, Position=7, HelpMessage="Channels to provision (semicolon separated list)")]
    [string] $channels,

    [Parameter(Mandatory=$false, Position=8, HelpMessage="Action to take")]
    [ValidateSet('','Provision','Deprovision')] [string] $action = "",

    [Parameter(Mandatory=$false, Position=9, HelpMessage="Prevent script runner from being owner of Team?")]
    [switch] $removeCreatorFromTeam = $true
)

$ErrorActionPreference = "Continue"

function Add-Channels
{   
   param (   
             $channels,$groupId
         )   
    Process
    {

        $teamchannels = $channels -split ";" 
        if($teamchannels)
        {
            for($i =0; $i -le ($teamchannels.count - 1) ; $i++)
            {
                Write-Host "---> Adding channel '$($teamchannels[$i])' to $groupId..." -NoNewline
                Write-Output "---> Adding channel '$($teamchannels[$i])' to $groupId..." | out-file $global:logFilePath -NoNewline -Append

                try {
                    New-TeamChannel -GroupId $groupId -DisplayName $($teamchannels[$i])

                    Write-Host "Done" -BackgroundColor Green
                    Write-Output "Done" | out-file $global:logFilePath -Append
                }
                Catch{
                    Write-Host "Failed. A channel with this name may already exist." -BackgroundColor Red
                    Write-Output "Failed. A channel with this name may already exist." | out-file $global:logFilePath -Append
                }
            }
        }
    }
}

function Add-Users
{   
    param(   
             $users,$groupId,$currentUsername,$role
          )   
    Process
    {
        
        $teamusers = $users -split ";" 
        if($teamusers)
        {
            for($j =0; $j -le ($teamusers.count - 1) ; $j++)
            {
                if($teamusers[$j] -ne $currentUsername)
                {
                    Write-Host "---> Adding '$($teamusers[$j])' to $groupId as $role ..." -NoNewline
                    Write-Output "---> Adding '$($teamusers[$j])' to $groupId as $role ..." | out-file $global:logFilePath -NoNewline -Append

                    try {
                        Add-TeamUser -GroupId $GroupId -User $($teamusers[$j]) -Role $Role

                        Write-Host "Done" -BackgroundColor Green
                        Write-Output "Done" | out-file $global:logFilePath -Append
                    }
                    catch {
                        Write-Host "Failed. Check that user UPN is valid and the user is not already present in this role." -BackgroundColor Red
                        Write-Output "Failed. Check that user UPN is valid and the user is not already present in this role." | out-file $global:logFilePath -Append
                    }
                }
            }
        }
    }
}

function Provision-ATeam([PSCustomObject] $row)
{
    $teamGroupId = $($row.TeamGroupId.Trim())
    $teamDisplayName = $($row.TeamDisplayName.Trim())
    $visibility = $($row.TeamPrivacy.Trim())
    
    if([string]::IsNullOrWhiteSpace($visibility)) {
        $visibility = "Private"
    }

    $channels = $($row.Channels.Trim())
    $owners = $($row.OwnerUPNs.Trim())
    $members = $($row.MemberUPNs.Trim())

    $optionalParams = @{}

    $teamDescription = $($row.TeamDescription.Trim())
    if(![string]::IsNullOrEmpty($teamDescription)){
        $optionalParams.Add("Description", $teamDescription)
    }

    $mailNickName = $($row.MailNickName.Trim())
    if(![string]::IsNullOrEmpty($mailNickName)){
        $optionalParams.Add("MailNickName", $mailNickName)
    }

    if($row.AllowGiphy -ne $null) {
        $optionalParams.Add("AllowGiphy", [System.Convert]::ToBoolean($row.AllowGiphy))
    }

    $giphyContentRating = $($row.GiphyContentRating.Trim())
    if(![string]::IsNullOrEmpty($giphyContentRating)){
        $optionalParams.Add("GiphyContentRating", $giphyContentRating)
    }

    if($row.AllowStickersAndMemes -ne $null) {
        $optionalParams.Add("AllowStickersAndMemes",[System.Convert]::ToBoolean($row.AllowStickersAndMemes))
    }

    if($row.AllowCustomMemes -ne $null) {
        $optionalParams.Add("AllowCustomMemes",[System.Convert]::ToBoolean($row.AllowCustomMemes))
    }

    if($row.AllowGuestCreateUpdateChannels -ne $null) {
        $optionalParams.Add("AllowGuestCreateUpdateChannels",[System.Convert]::ToBoolean($row.AllowGuestCreateUpdateChannels))
    }

    if($row.AllowGuestDeleteChannels -ne $null) {
        $optionalParams.Add("AllowGuestDeleteChannels",[System.Convert]::ToBoolean($row.AllowGuestDeleteChannels))
    }

    if($row.AllowCreateUpdateChannels -ne $null) {
        $optionalParams.Add("AllowCreateUpdateChannels",[System.Convert]::ToBoolean($row.AllowCreateUpdateChannels))
    }

    if($row.AllowDeleteChannels -ne $null) {
        $optionalParams.Add("AllowDeleteChannels",[System.Convert]::ToBoolean($row.AllowDeleteChannels))
    }

    if($row.AllowAddRemoveApps -ne $null) {
        $optionalParams.Add("AllowAddRemoveApps",[System.Convert]::ToBoolean($row.AllowAddRemoveApps))
    }
        
    if($row.AllowCreateUpdateRemoveTabs -ne $null) {
        $optionalParams.Add("AllowCreateUpdateRemoveTabs",[System.Convert]::ToBoolean($row.AllowCreateUpdateRemoveTabs))
    }

    if($row.AllowCreateUpdateRemoveConnectors -ne $null) {
        $optionalParams.Add("AllowCreateUpdateRemoveConnectors",[System.Convert]::ToBoolean($row.AllowCreateUpdateRemoveConnectors))
    }

    if($row.AllowUserEditMessages -ne $null) {
        $optionalParams.Add("AllowUserEditMessages",[System.Convert]::ToBoolean($row.AllowUserEditMessages))
    }

    if($row.AllowUserDeleteMessages -ne $null) {
        $optionalParams.Add("AllowUserDeleteMessages",[System.Convert]::ToBoolean($row.AllowUserDeleteMessages))
    }

    if($row.AllowOwnerDeleteMessages -ne $null) {
        $optionalParams.Add("AllowOwnerDeleteMessages",[System.Convert]::ToBoolean($row.AllowOwnerDeleteMessages))
    }

    if($row.AllowTeamMentions -ne $null) {
        $optionalParams.Add("AllowTeamMentions",[System.Convert]::ToBoolean($row.AllowTeamMentions))
    }

    if($row.AllowChannelMentions -ne $null) {
        $optionalParams.Add("AllowChannelMentions",[System.Convert]::ToBoolean($row.AllowChannelMentions))
    }

    if($row.ShowInTeamsSearchAndSuggestions -ne $null) {
        $optionalParams.Add("ShowInTeamsSearchAndSuggestions",[System.Convert]::ToBoolean($row.ShowInTeamsSearchAndSuggestions))
    }

    Write-Host "You are provisioning a team with group id '$teamGroupId' or named '$teamDisplayName'..." -NoNewline
    Write-Output "You are provisioning a team with group id '$teamGroupId' or named '$teamDisplayName'..." | out-file $global:logFilePath -NoNewline -Append

    $group = $null

    if(![string]::IsNullOrWhiteSpace($teamGroupId)) {
        # provision based on group Id
        $group = New-Team -GroupId $teamGroupId @optionalParams
    }
    else {
        # provision based on name
        $group = New-Team -DisplayName $teamDisplayName -Visibility $visibility @optionalParams
    }

    if($group -ne $null) {
        Write-Host "Done. GroupId: $($group.GroupId)" -BackgroundColor Green
        Write-Output "Done. GroupId: $($group.GroupId)" | out-file $global:logFilePath -Append

        if(![string]::IsNullOrEmpty($channels)) {
            Add-Channels -channels $channels -groupId $group.GroupId
        }

        if(![string]::IsNullOrEmpty($owners)){
            Add-Users -users $owners -groupId $group.GroupId -currentUsername $global:currentUser -role Owner
        }

        if(![string]::IsNullOrEmpty($members)){
            Add-Users -users $members -groupId $group.GroupId -currentUsername $global:currentUser -role Member 
        }
        
        if($removeCreatorFromTeam) {
            try {
                # Remove user running the script who's added as owner. If the user is the only owner then will not be removed.
                Write-Host "Removing provisioning user from owner role on $($group.GroupId)..." -NoNewline
                Write-Output "Removing provisioning user from owner role on $($group.GroupId)..." | out-file $global:logFilePath -NoNewline -Append

                # demote from owner
                Remove-TeamUser -GroupId $($group.GroupId) -User $global:currentUser -Role Owner
                # remove completely
                Remove-TeamUser -GroupId $($group.GroupId) -User $global:currentUser

                Write-Host "Done" -BackgroundColor Green
                Write-Output "Done" | out-file $global:logFilePath -Append
            }
            catch {
                Write-Host "Failed. Possibly because no other owner was specified." -BackgroundColor Red
                Write-Output "Failed. Possibly because no other owner was specified." | out-file $global:logFilePath -Append
            }
        }

        Get-Team -GroupId $($group.GroupId) | fl | out-file $global:logFilePath -Append
          
        # populate log
        "$teamDisplayName `t $visibility `t $($group.GroupId) `t $owners `t $members `t $channels `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 
    }
    else {
        Write-Host "Failed" -BackgroundColor Red
        Write-Output "Failed" | out-file $global:logFilePath -Append
    }
}

function Deprovision-ATeam($groupId, $teamName)
{
    Write-Host "WARNING: You are deprovisioning the team named '$teamName' and a groupId of '$groupId'. Once a team is deprovisioned, ALL associated content such as channels, messages, team mailbox, emails, planner and files will be permanently deleted...." -NoNewline -ForegroundColor Red
    Write-Output "WARNING: You are deprovisioning the team named '$teamName' and a groupId of '$groupId'. Once a team is deprovisioned, ALL associated content such as channels, messages, team mailbox, emails, planner and files will be permanently deleted...." | out-file $global:logFilePath -NoNewline -Append 

    Remove-Team -GroupId $groupId

    Write-Host "Done" -BackgroundColor Green
    Write-Output "Done" | out-file $global:logFilePath -Append
}

function Action-ATeam([PSCustomObject] $row)
{   
    $teamGroupId = $($row.TeamGroupId.Trim())  
    $teamDisplayName = $($row.TeamDisplayName.Trim())
    
    if([string]::IsNullOrWhiteSpace($teamGroupId) -and [string]::IsNullOrWhiteSpace($teamDisplayName)) {
        Write-Host "Either group Id or team display name must be specified..." -NoNewline
        Write-Host "Quitting" -BackgroundColor Red

        return
    }

    Write-Host "`n--------------------------------------------------------------------------------------"
    Write-Host "Checking if a team with name '$teamDisplayName' or id '$teamGroupId' exists..." -NoNewline
    Write-Output "Checking if a team with name '$teamDisplayName' or id '$teamGroupId' exists..." | out-file $global:logFilePath -NoNewline -Append 

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
        Write-Host "No" -BackgroundColor Green
        Write-Output "No" | out-file $global:logFilePath -Append
    }
    else {
        $exists = $true

        Write-Host "...Yes. GroupId: $($team.GroupId)..." -BackgroundColor Red
        Write-Output "...Yes. GroupId: $($team.GroupId)..." | out-file $global:logFilePath -Append 
    }

    switch($action)
    {
        'Provision'{            
            if(!$exists) {
                Provision-ATeam $row
            }
            else {
                Write-Host "Will NOT be reprovisioned" -BackgroundColor Red
                Write-Output "Will NOT be reprovisioned" | out-file $global:logFilePath -Append 
            }
        }
        'Deprovision'{
            
            if(!$exists) {
                Write-Host "No action taken" -BackgroundColor Green
                Write-Output "No action taken" | out-file $global:logFilePath -Append 
            }
            else {
                Deprovision-ATeam $($team.GroupId) $teamDisplayName
            }
        }
        default{
            Write-Host "`nYou are NOT taking any action on a team named '$teamDisplayName'..."
            Write-Output "`nYou are NOT taking any action on a team named '$teamDisplayName'..." | out-file $global:logFilePath -Append
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
    "TeamName `t Type `t GroupId `t Owners `t Members `t Channels `t Date `t Action" | out-file $global:logCSVPath

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
                    TeamDescription=[string]::Empty;
                    MailNickName=[string]::Empty;
                    TeamPrivacy=$teamPrivacy;
                    OwnerUPNs=$owners;
                    MemberUPNs=$members;
                    Channels=$channels;
                    AllowGiphy=$true;
                    GiphyContentRating="Moderate";
                    AllowStickersAndMemes=$true;
                    AllowCustomMemes=$true;
                    AllowGuestCreateUpdateChannels=$false;
                    AllowGuestDeleteChannels=$false;
                    AllowCreateUpdateChannels=$true;
                    AllowDeleteChannels=$false;
                    AllowAddRemoveApps=$true;
                    AllowCreateUpdateRemoveTabs=$true;
                    AllowCreateUpdateRemoveConnectors=$true;
                    AllowUserEditMessages=$true;
                    AllowUserDeleteMessages=$true;
                    AllowOwnerDeleteMessages=$false;
                    AllowTeamMentions=$true;
                    AllowChannelMentions=$true;
                    ShowInTeamsSearchAndSuggestions=$true
                }
    
            Action-ATeam $row | Out-Null
        }
    }

    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append

    Disconnect-MicrosoftTeams
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow