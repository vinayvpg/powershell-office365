<#  
 .NOTES
 ===========================================================================
 Created On:   8/1/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Provision-UnifiedGroup.ps1
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
    Bulk or individual provisioning and/or deprovisioning of a Office 365 Groups (Unified Group).

.Description  
    Bulk or individual provisioning and/or deprovisioning of a Office 365 Groups (Unified Group). Script prompts for user credential on every run.
    It should be run by a user who has Office 365 group creation rights in Azure AD.
         
.PARAMETER - inputCSVPath (string)
    Path to csv file containing batch of Groups to provision. The csv schema contains following columns - GroupDisplayName,GroupDescription,Alias,GroupPrivacy,OwnerUPNs,MemberUPNs
.PARAMETER - outputLogFolderPath (string)
    Path where the logs should be created
.PARAMETER - groupDisplayName (string < 256 characters required)
    Display name of the Group.
.PARAMETER - groupDescription (string)
    Description of the group. If empty then display name becomes description as well.
.PARAMETER - groupPrivacy (string)
    Privacy setting for the Group. Default is private.
.PARAMETER - owners (string)
    Semicolon separated list of UPNs.
.PARAMETER - members (string)
    Semicolon separated list of UPNs.
.PARAMETER - alias (string)
    Email alias for the group. Should be unique within Azure AD.
.PARAMETER - action (string)
    Provision, deprovision. Default is empty string which means no action.

.USAGE 
    Bulk provision Groups specified in csv in a single batch operation
     
    PS >  Provision-UnifiedGroup.ps1 -inputCSVPath "c:/temp/O365groups.csv" -action "Provision"
.USAGE 
    Provision individual O365 group with default parameters
     
    PS >  Provision-UnifiedGroup.ps1 -groupDisplayName "Company function Group" -action "Provision"
.USAGE 
    Provision individual Group with specific parameters
     
    PS >  Provision-UnifiedGroup.ps1 -groupDisplayName "Company function Group" -groupDescription "Detailed description of what the group is for" -groupPrivacy "Public" -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -alias "uniquealiasforgroup" -action "Provision"
.USAGE 
    Deprovision individual group
     
    PS >  Provision-UnifiedGroup.ps1 -groupDisplayName "Name of group to deprovision" -action "Deprovision"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on Office 365 Groups to be provisioned or deprovisioned")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Group display name")]
    [string] $groupDisplayName,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Group description")]
    [string] $groupDescription,

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Group privacy setting")]
    [ValidateSet('Private','Public')][string] $groupPrivacy="Private",

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Owners (semicolon separated list of UPNs)")]
    [string] $owners,

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Members (semicolon separated list of UPNs)")]
    [string] $members,

    [Parameter(Mandatory=$false, Position=7, HelpMessage="Alias or mail nick name for the group (should be unique in tenant)")]
    [string] $alias,

    [Parameter(Mandatory=$false, Position=8, HelpMessage="Action to take")]
    [ValidateSet('','Provision','Deprovision')] [string] $action = ""
)

$ErrorActionPreference = "Continue"

function Add-Users
{   
    param(   
             $users,$groupId,$currentUsername,$role
          )   
    Process
    {
        
        try{
                $groupUsers = $users -split ";" 
                if($groupUsers)
                {
                    for($j =0; $j -le ($groupUsers.count - 1) ; $j++)
                    {
                        #if($groupUsers[$j] -ne $currentUsername)
                        #{
                            Write-Host "---> Adding '$($groupUsers[$j])' to $groupId as $role ..." -NoNewline
                            Write-Output "---> Adding '$($groupUsers[$j])' to $groupId as $role ..." | out-file $global:logFilePath -NoNewline -Append

                            Add-UnifiedGroupLinks -Identity $groupId -Links $($groupUsers[$j]) -LinkType "Members"

                            # owners have to be added as members first
                            if($role -eq "Owner") {
                                Add-UnifiedGroupLinks -Identity $groupId -Links $($groupUsers[$j]) -LinkType "Owners"
                            }

                            Write-Host "Done" -BackgroundColor Green
                            Write-Output "Done" | out-file $global:logFilePath -Append

                        #}
                    }
                }
            }
        Catch
            {
            }
        }
}

function Provision-AUnifiedGroup([PSCustomObject] $row)
{
    $groupDisplayName = $($row.GroupDisplayName.Trim())
    $accessType = $($row.GroupPrivacy.Trim())

    if([string]::IsNullOrWhiteSpace($accessType)) {
        $accessType = "Private"
    }

    $owners = $($row.OwnerUPNs.Trim())
    $members = $($row.MemberUPNs.Trim())

    $optionalParams = @{}

    $groupDescription = $($row.GroupDescription.Trim())
    if(![string]::IsNullOrEmpty($groupDescription)){
        $optionalParams.Add("Notes", $groupDescription)
    }

    $alias = $($row.Alias.Trim())
    if(![string]::IsNullOrEmpty($alias)){
        $optionalParams.Add("Alias", $alias)
    }

    Write-Host "You are provisioning a group named '$groupDisplayName'..." -NoNewline
    Write-Output "You are provisioning a group named '$groupDisplayName'..." | out-file $global:logFilePath -NoNewline -Append

    $group = New-UnifiedGroup -DisplayName $groupDisplayName -AccessType $accessType @optionalParams

    if($group -ne $null) {
        Write-Host "Done. GroupId: $($group.ExternalDirectoryObjectId)" -BackgroundColor Green
        Write-Output "Done. GroupId: $($group.ExternalDirectoryObjectId)" | out-file $global:logFilePath -Append

        Write-Host "Hiding O365 group from GAL..." -NoNewline
        Write-Output "Hiding O365 group from GAL..." | out-file $global:logFilePath -NoNewline -Append

        Set-UnifiedGroup -Identity $($group.ExternalDirectoryObjectId) -HiddenFromAddressListsEnabled $true

        Write-Host "Done" -BackgroundColor Green
        Write-Output "Done" | out-file $global:logFilePath -Append

        if(![string]::IsNullOrEmpty($owners)){
            Add-Users -users $owners -groupId $($group.ExternalDirectoryObjectId) -currentUsername $global:currentUser -role Owner
        }

        if(![string]::IsNullOrEmpty($members)){
            Add-Users -users $members -groupId $($group.ExternalDirectoryObjectId) -currentUsername $global:currentUser -role Member 
        }
        
        # output all properties to detailed log
        Get-UnifiedGroup -Identity $($group.ExternalDirectoryObjectId) | fl | out-file $global:logFilePath -Append
          
        # populate log
        "$groupDisplayName `t $accessType `t $($group.ExternalDirectoryObjectId) `t $owners `t $members `t $($group.Alias) `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 
    }
    else {
        Write-Host "Failed" -BackgroundColor Red
        Write-Output "Failed" | out-file $global:logFilePath -Append
    }
}

function Deprovision-AUnifiedGroup($groupId, $groupName)
{
    Write-Host "WARNING: You are deprovisioning the group named '$groupName' and a groupId of '$groupId'. Once a group is deprovisioned, ALL associated assets such as team, channels, messages, team mailbox, emails, planner and files will be permanently deleted...." -NoNewline -ForegroundColor Red
    Write-Output "WARNING: You are deprovisioning the group named '$groupName' and a groupId of '$groupId'. Once a group is deprovisioned, ALL associated assets such as team, channels, messages, team mailbox, emails, planner and files will be permanently deleted...." | out-file $global:logFilePath -NoNewline -Append 

    Remove-UnifiedGroup -Identity $groupId -Confirm

    Write-Host "Done" -BackgroundColor Green
    Write-Output "Done" | out-file $global:logFilePath -Append
}

function Action-AUnifiedGroup([PSCustomObject] $row)
{     
    $groupDisplayName = $($row.GroupDisplayName.Trim())

    if([string]::IsNullOrWhiteSpace($groupDisplayName)) {
        Write-Host "Group display name must be specified..." -NoNewline
        Write-Host "Quitting" -BackgroundColor Red

        return
    }

    Write-Host "Checking if a unified group with name '$groupDisplayName' already exists..." -NoNewline
    Write-Output "Checking if a unified group with name '$groupDisplayName' already exists..." | out-file $global:logFilePath -NoNewline -Append 

    # check if Group with the same name already exists
    try {
        $group = Get-UnifiedGroup -Identity $groupDisplayName
    }
    catch {}

    $exists = $false

    if($group -eq $null) {
        Write-Host "...No..." -BackgroundColor Green
        Write-Output "...No..." | out-file $global:logFilePath -Append
    }
    else {
        $exists = $true

        Write-Host "...Yes. GroupId: $($group.ExternalDirectoryObjectId)..." -BackgroundColor Red
        Write-Output "...Yes. GroupId: $($group.ExternalDirectoryObjectId)..." | out-file $global:logFilePath -Append 
    }

    switch($action)
    {
        'Provision'{            
            if(!$exists) {
                Provision-AUnifiedGroup $row
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
                Deprovision-AUnifiedGroup $($group.ExternalDirectoryObjectId) $groupDisplayName
            }
        }
        default{
            Write-Host "`nYou are NOT taking any action on a group named '$groupDisplayName'..."
            Write-Output "`nYou are NOT taking any action on a group named '$groupDisplayName'..." | out-file $global:logFilePath -Append
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
    $global:currentUser = $global:cred.UserName
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

        $global:logCSVPath = $global:logFolderPath + "ProvisioningLog.csv"
        $global:logFilePath = $global:logFolderPath + "ActionLog.log"

        Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

        #log csv
        "GroupName `t Type `t GroupId `t Owners `t Members `t Alias `t Date `t Action" | out-file $global:logCSVPath

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
        
                if([string]::IsNullOrWhiteSpace($groupDisplayName))
                {
                    do {
                        $groupDisplayName = Read-Host "Specify the display name for the group"
                    }
                    until (![string]::IsNullOrWhiteSpace($groupDisplayName))
                }

                $row = @{
                        GroupDisplayName=$groupDisplayName;
                        GroupDescription=$groupDescription;
                        Alias=$alias;
                        GroupPrivacy=$groupPrivacy;
                        OwnerUPNs=$owners;
                        MemberUPNs=$members
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