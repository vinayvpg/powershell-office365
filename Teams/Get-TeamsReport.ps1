
[CmdletBinding()]
param(
    
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=1, HelpMessage="O365 Group Id filter")]
    [string] $groupId,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Team display name filter")]
    [string] $teamDisplayName
)

$ErrorActionPreference = "Continue"

function Get-Teams()
{   
    $teams = $null

    $allTeams = Get-Team

    if(![string]::IsNullOrWhiteSpace($groupId)) {
        Write-Host "Finding Teams with GroupId: $groupId..." -ForegroundColor Cyan
        $teams = $allTeams | ? {$_.GroupId -eq $groupId}
    }
    elseif(![string]::IsNullOrWhiteSpace($teamDisplayName)) {
        Write-Host "Finding Teams with Display Name like: $teamDisplayName..." -ForegroundColor Cyan
        $teams = $allTeams | ? {$_.DisplayName -like $teamDisplayName}
    }
    else {
        $teams = $allTeams
    }

    if($teams -ne $null) {

        Write-Host "$($teams.Count) Teams found in the tenant " -BackgroundColor Green

        foreach($team in $teams) {
            Write-Host "----------------------------------------------------------------------"
            Write-Host "Processing team '$($team.DisplayName)'..." -NoNewline

            $displayName = $team.DisplayName
            $groupId = $team.GroupId
            $visibility = $team.Visibility
            $channels = (Get-TeamChannel -GroupId $groupId | % {$_.DisplayName}) -join '; '
            $channelsCount = (Get-TeamChannel -GroupId $groupId).Count
            $owners = (Get-TeamUser -GroupId $groupId -Role Owner | % {$_.User}) -join '; '
            $ownersCount = (Get-TeamUser -GroupId $groupId -Role Owner).Count
            $members = (Get-TeamUser -GroupId $groupId -Role Member | % {$_.User}) -join '; '
            $membersCount = (Get-TeamUser -GroupId $groupId -Role Member).Count

            # populate log
            "$displayName `t $visibility `t $groupId `t $owners `t $ownersCount `t $members `t $membersCount `t $channels `t $channelsCount `t $(Get-Date)" | out-file $global:logCSVPath -Append 
            
            Write-Host "Done" -BackgroundColor Green
        }
    }
    else {
        Write-Host "Did not find any Team in the tenant" -BackgroundColor Red
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

    $global:logCSVPath = $global:logFolderPath + "AllTeams.csv"
    #$global:logFilePath = $global:logFolderPath + "ActionLog.log"

    #log csv
    "TeamName `t Type `t GroupId `t Owners `t OwnersCount `t Members `t MembersCount `t Channels `t ChannelsCount `t RunDate" | out-file $global:logCSVPath

    Get-Teams

    Disconnect-MicrosoftTeams
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow