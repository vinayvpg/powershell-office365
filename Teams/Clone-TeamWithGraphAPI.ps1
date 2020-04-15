<#
.SYNOPSIS
    Bulk or individual cloning of a Microsoft Team in Office 365 using Microsoft Graph API.

.NOTES
 ===========================================================================
 Created On:   10/31/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Clone-TeamWithGraphAPI.ps1
 Version:      1.0.15
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
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on Teams to be provisioned")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Team display name")]
    [string] $teamDisplayName="Vinay Team 1",

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Team description")]
    [string] $teamDescription,

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Template Team display name - New team will be cloned from this team")]
    [string] $templateTeamName="Private Channel Template Team",

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Tenant root site collection url e.g. https://office365ninja.sharepoint.com")]
    [string] $tenantRootUrl="https://office365ninja.sharepoint.com",

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Tenant Id from Azure Enterprise App Registration e.g. 538d1ac4-c163-43ae-b40c-c27b6b8e4f96")]
    [string] $tenantId="538d1ac4-c163-43ae-b40c-c27b6b8e4f96",

    [Parameter(Mandatory=$false, Position=7, HelpMessage="Client Id from Azure Enterprise App Registration e.g. 52614244-27af-4f33-a77a-6fe724d7d638")]
    [string] $clientId="52614244-27af-4f33-a77a-6fe724d7d638",

    [Parameter(Mandatory=$false, Position=8, HelpMessage="Client Secret from Azure Enterprise App Registration e.g. Xj2fC5H=Y.QLN6pWmby:nv8evM:-s45@")]
    [string] $clientSecret="Xj2fC5H=Y.QLN6pWmby:nv8evM:-s45@",

    [Parameter(Mandatory=$false, Position=9, HelpMessage="Redirect URI from Azure Enterprise App Registration e.g. https://office365ninja-apiaccessor-app/login")]
    [string] $redirectUri="https://office365ninja-apiaccessor-app/login",

    [Parameter(Mandatory=$false, Position=10, HelpMessage="Team privacy setting")]
    [ValidateSet('Private','Public','HiddenMembership')][string] $teamPrivacy="Private",

    [Parameter(Mandatory=$false, Position=11, HelpMessage="Owners (semicolon separated list of UPNs)")]
    [string] $owners="AdeleV@office365ninja.onmicrosoft.com",

    [Parameter(Mandatory=$false, Position=12, HelpMessage="Members (semicolon separated list of UPNs) e.g. diegos@office365ninja.onmicrosoft.com;johannal@office365ninja.onmicrosoft.com")]
    [string] $members,

    [Parameter(Mandatory=$false, Position=13, HelpMessage="Action to take")]
    [ValidateSet('','Clone')] [string] $action = "Clone",

    [Parameter(Mandatory=$false, Position=14, HelpMessage="Remove wiki tab(s)?")]
    [switch] $removeWikiTab = $true,

    [Parameter(Mandatory=$false, Position=15, HelpMessage="Prevent script runner from being owner of Team?")]
    [switch] $removeCreatorFromTeam = $false,

    [Parameter(Mandatory=$false, Position=16, HelpMessage="Auto favorite a channel? This setting does not seem to be cloned.")]
    [switch] $autoFavoriteChannel = $true
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

function Add-TeamUsers
{   
    param(   
             $users,$teamId,$role
          )   
    Process
    {
        Write-Host "`n--------------------------------------------------------------------------------------"
     
        $teamusers = $users -split ";" 
        if($teamusers)
        {
            for($j =0; $j -le ($teamusers.count - 1) ; $j++)
            {
                Write-Message "---> Adding '$($teamusers[$j])' to Team '$teamId' as '$role' ..." -NoNewline

                $userId = [string]::Empty

                $user = Get-O365User $($teamusers[$j])

                if($user -ne $null) {
                    $userId = $user.id

                    Write-Message "Found user '$($teamusers[$j])' with id '$userId'...." -NoNewline
                }

                if(![string]::IsNullOrEmpty($userId)) {
                    try {
                        Add-TeamUser -teamId $teamId -userId $userId -role $role

                        Write-Message "Done" -BackgroundColor Green
                    }
                    Catch {
                        Write-Message "Failed. Check that user UPN is valid and the user is not already present in this role." -BackgroundColor Red
                    }
                }
                else {
                    Write-Message "User '$($teamusers[$j])' not found in directory. Not added to Team."
                }
            }
        }
    }
}

function Action-ATeam([PSCustomObject] $row)
{     
    $teamDisplayName = $($row.TeamDisplayName.Trim())
    $teamDescription = $($row.TeamDescription.Trim())
    $templateTeamName = $($row.TemplateTeamName.Trim())
    $visibility = $($row.TeamPrivacy.Trim())
    
    if([string]::IsNullOrWhiteSpace($visibility)) {
        $visibility = "Private"
    }

    $owners = $($row.OwnerUPNs.Trim())
    $members = $($row.MemberUPNs.Trim())

    if([string]::IsNullOrWhiteSpace($teamDisplayName) -or [string]::IsNullOrWhiteSpace($templateTeamName)) {
        Write-Message -msg "Team display name and Template Team Name must be specified..." -NoNewLine 
        Write-Message -msg "Quitting" -BackgroundColor Red

        return
    }

    if([string]::IsNullOrEmpty($global:authCode)) {
        Write-Message -msg "Authentication code could not be retrieved for TenantId: '$global:tenantId', ClientId: '$global:clientId', ClientSecret: '$global:clientSecret'..." -NoNewline
        Write-Message -msg "Quitting" -BackgroundColor Red

        return
    }

    <#
    $allTeams = Get-AllTeams
        
    if(($allTeams -eq $null) -or ($allTeams.Count -eq 0)) {
        Write-Message -msg "No teams were found in this tenant..." -NoNewline
        Write-Message -msg "Quitting" -BackgroundColor Red

        return
    }
    #>

    Write-Host "`n--------------------------------------------------------------------------------------"
    #Write-Message "Found $($allTeams.count) Teams in the tenant..." -ForegroundColor Cyan
                
    Write-Message "Checking if a team with name '$teamDisplayName' exists..." -NoNewline

    #$cloneTeam = $allTeams | ? {$_.displayName -eq $teamDisplayName}
    $cloneTeam = Get-TeamsByDisplayName $teamDisplayName $true

    if($cloneTeam -ne $null) {
        Write-Message -msg "Team with display name '$teamDisplayName' already exists..." -NoNewline
        Write-Message -msg "Quitting" -BackgroundColor Red

        return
    }

    Write-Message -msg "No" -BackgroundColor Green

    Write-Message -msg "Checking if a template team with name '$templateTeamName' exists..." -NoNewline

    #$templateTeam = $allTeams | ? {$_.displayName -eq $templateTeamName}
    $templateTeam = Get-TeamsByDisplayName $templateTeamName $true

    if($templateTeam -eq $null) {
        Write-Message -msg "Template team with display name '$templateTeamName' DOES NOT exist..." -NoNewline
        Write-Message -msg "Quitting" -BackgroundColor Red

        return
    }

    $templateTeamId = $templateTeam.id               #"60410af1-0859-4fb3-b5a3-9ed2462331db"
                    
    Write-Message -msg "Yes. Template Team Id: '$templateTeamId'" -BackgroundColor Green

    $templateTeamDriveRes = Get-TeamDefaultDriveUrl $templateTeamId $false

    $templateTeamDriveId = $templateTeamDriveRes.DriveId
    $templateTeamDriveFullUrl = $templateTeamDriveRes.DriveUrl

    $templateTeamDriveUrl = Get-ServerRelativeUrl $templateTeamDriveFullUrl
            
    if([string]::IsNullOrEmpty($action)) {
        Write-Message -msg "You have chosen to NOT take any provisioning action..." -NoNewline
        Write-Message -msg "Quitting" -BackgroundColor Red

        return
    }

    $cloneTeamId = Clone-Team $($templateTeam.id) $teamDisplayName $teamPrivacy $teamDescription
            
    if([string]::IsNullOrEmpty($cloneTeamId)) {
        Write-Message "Clone team with name '$teamDisplayName' was not created..." -NoNewline
        Write-Message "Quitting" -BackgroundColor Red

        return
    }

    $cloneTeamDriveRes = Get-TeamDefaultDriveUrl $cloneTeamId $false

    $cloneTeamDriveId = $cloneTeamDriveRes.DriveId
    $cloneTeamDriveFullUrl = $cloneTeamDriveRes.DriveUrl

    $cloneTeamDriveUrl = Get-ServerRelativeUrl $cloneTeamDriveFullUrl        

    if([string]::IsNullOrEmpty($cloneTeamDriveFullUrl)) {
        Write-Message "Clone team drive was not provisioned. PLEASE DELETE THIS TEAM and try again later..." -NoNewline -ForegroundColor Red
        Write-Message "Quitting" -BackgroundColor Red

        return
    }

    #log csv
    "$teamDisplayName `t $teamPrivacy `t $cloneTeamId `t $templateTeamName `t $templateTeamId `t $owners `t $members `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append

    Write-Host "`n--------------------------------------------------------------------------------------"
    Write-Message "Getting template team channels..." -NoNewline

    $templateTeamChannels = (Get-TeamChannels $templateTeamId).value | Select-Object id, displayName, membershipType

    Write-Message "Done. $($templateTeamChannels.count) channels found." -BackgroundColor Green

    for($i=0; $i -le ($templateTeamChannels.count - 1); $i++) {
        Write-Host "`n--------------------------------------------------------------------------------------------------"

        $cloneTeamChannelId = [string]::Empty

        Write-Message "Getting channel id for clone team '$cloneTeamId' channel named '$($templateTeamChannels[$i].displayName)'..." -NoNewline

        $cloneTeamChannelId = Get-TeamChannelIdByName $cloneTeamId $($templateTeamChannels[$i].displayName)

        Write-Message "Done. Channel Id: '$cloneTeamChannelId'" -BackgroundColor Green

        if([string]::IsNullOrEmpty($cloneTeamChannelId)) {
            Write-Message "Clone team with name '$teamDisplayName' DOES NOT seem to have a channel named '$($templateTeamChannels[$i].displayName)'..." -NoNewline
            
            Write-Message "'$($templateTeamChannels[$i].displayName)' is a '$($templateTeamChannels[$i].membershipType)' channel..." -BackgroundColor Red

            if($templateTeamChannels[$i].membershipType -eq "private") {
                # create the channel if it is private because clone operation does not create it

                $privateChannelOwnersBody = @(
                        @{
                           "@odata.type" = "#microsoft.graph.aadUserConversationMember";
                           "user@odata.bind" = "https://graph.microsoft.com/beta/users('$($global:myId)')";
                           "roles" = '["owner"]'
                        }
                     )
                
                if(![string]::IsNullOrEmpty($owners)) {

                }

                $cloneTeamAddPrivateChannelBodyJson = @{ "@odata.type" = "#Microsoft.Teams.Core.channel";
                                                            displayName = $($templateTeamChannels[$i].displayName);
                                                            isFavoriteByDefault = $true;
                                                            membershipType = "private";
                                                            members = $privateChannelOwnersBody 
                                                       } | ConvertTo-Json
            
                Write-Host "Add channel Request Body : $cloneTeamAddPrivateChannelBodyJson"
            
                Add-TeamChannel $cloneTeamId $cloneTeamAddPrivateChannelBodyJson

                Write-Message "Done" -BackgroundColor Green
            }
            else {
                Write-Message "Skipped" -BackgroundColor Red
            }

            break
        }
        
        if($autoFavoriteChannel -and ($($templateTeamChannels[$i].displayName) -ne "General")) {
            Write-Message "Setting clone team '$cloneTeamId' channel named '$($templateTeamChannels[$i].displayName)' as a favorite for all members..."

            $cloneTeamChannelBodyJson = @{ isFavoriteByDefault = $true } | ConvertTo-Json
            
            Write-Host "Update Request Body : $cloneTeamChannelBodyJson"
            
            Update-TeamChannel $cloneTeamId $cloneTeamChannelId $cloneTeamChannelBodyJson

            Write-Message "Done" -BackgroundColor Green
        }

        $templateTeamChannelFolderDriveItemId = Get-ChannelFolderDriveItemId -teamId $templateTeamId -channelDisplayName $($templateTeamChannels[$i].displayName)
        $cloneTeamChannelFolderDriveItemId = Get-ChannelFolderDriveItemId -teamId $cloneTeamId -channelDisplayName $($templateTeamChannels[$i].displayName)

        [object[]] $channelFilesColl = Get-ChannelFiles -teamId $templateTeamId -channelDisplayName $($templateTeamChannels[$i].displayName)

        Write-Message "`nFound $($channelFilesColl.count) files/folders in channel '$($templateTeamChannels[$i].displayName)' in template Team '$templateTeamId'" -ForegroundColor Cyan

        for($f=0; $f -le ($channelFilesColl.count - 1); $f++) {
            
            $sourceItemId = $channelFilesColl[$f].id

            if([string]::IsNullOrEmpty($templateTeamDriveId) -or [string]::IsNullOrEmpty($sourceItemId) -or [string]::IsNullOrEmpty($cloneTeamDriveId) -or [string]::IsNullOrEmpty($cloneTeamChannelFolderDriveItemId)) {
                Write-Message "One of the copy parameters is empty. NO COPY. Source Drive: '$templateTeamDriveId' Source Item: '$sourceItemId' Target Drive: '$cloneTeamDriveId' Target Folder: '$cloneTeamChannelFolderDriveItemId'" -ForegroundColor Red
            }
            else {
                Copy-ChannelFiles -sourceDriveId $templateTeamDriveId -sourceDriveItemId $sourceItemId -targetDriveId $cloneTeamDriveId -targetDriveItemId $cloneTeamChannelFolderDriveItemId
            }
        }

        <# Use with PnP only
        #$templateChannelFolderUrl = $templateTeamDriveUrl + "/" + [System.Web.HttpUtility]::UrlPathEncode($templateTeamChannels[$i].displayName)
        $templateChannelFolderUrl = [System.Web.HttpUtility]::UrlDecode($templateTeamDriveUrl) + "/" + $($templateTeamChannels[$i].displayName) # PnP CopyFile seems to require unescaped url path

        #$cloneChannelFolderUrl = $cloneTeamDriveUrl + "/" + [System.Web.HttpUtility]::UrlPathEncode($templateTeamChannels[$i].displayName)
        $cloneChannelFolderUrl = [System.Web.HttpUtility]::UrlDecode($cloneTeamDriveUrl) + "/" + $($templateTeamChannels[$i].displayName)

        Copy-TemplateTeamFiles $templateChannelFolderUrl $cloneChannelFolderUrl $true 
        #>

        Write-Message "`nGetting tabs for template channel named '$($templateTeamChannels[$i].displayName)'..." -NoNewline

        $templateTeamChannelTabs = (Get-TeamChannelTabs $templateTeamId $($templateTeamChannels[$i].id) $true).value

        Write-Message "Done. $($templateTeamChannelTabs.count) tabs found." -BackgroundColor Green

        for($j=0; $j -le ($templateTeamChannelTabs.count - 1); $j++) 
        {
            $templateTabAppId = $templateTeamChannelTabs[$j].teamsAppId
            $templateTabDisplayName = $templateTeamChannelTabs[$j].displayName

            Write-Message "`nGetting tab '$templateTabDisplayName' with AppId: '$templateTabAppId' for clone Team: '$cloneTeamId' Channel: '$cloneTeamChannelId'..." -NoNewline

            $cloneTeamChannelTab = Get-TeamChannelTabByAppIdAndName $cloneTeamId $cloneTeamChannelId $templateTabAppId $templateTabDisplayName

            $tabAction = "Update"

            $entityId = $null
            $contentUrl = $null
            $removeUrl = $null
            $websiteUrl = $null

            if($cloneTeamChannelTab -eq $null) {
                # check if cloning didn't bring over the tab. Create it.
                $tabAction = "Add"

                Write-Message "Not Found" -BackgroundColor Red
            }
            else {
                Write-Message "Found" -BackgroundColor Green
            }

            Write-Message "$tabAction clone team tab named '$templateTabDisplayName'..."

            switch ($templateTabAppId)
            {
                $($global:teamsAppIds.Website) {
                    $contentUrl = $templateTeamChannelTabs[$j].configuration.contentUrl
                    $websiteUrl = $templateTeamChannelTabs[$j].configuration.websiteUrl
 
                    break
                }
                $($global:teamsAppIds.OneNote) {
                
                    $defaultNotebook = Get-TeamDefaultNotebook $cloneTeamId

                    if($defaultNotebook -ne $null) {
                        #$notebookDisplayName = $defaultNotebook.displayName.ToString()
                        $notebookDisplayName = "Notes"
                        $notebookId = $defaultNotebook.id.ToString()
                        $oneNoteWebUrl = [System.Web.HttpUtility]::UrlEncode($defaultNotebook.links.oneNoteWebUrl.href)
                    }

                    $entityId = ([System.Guid]::NewGuid()).ToString() + "_" + $notebookId
                    $contentUrl = "https://www.onenote.com/teams/TabContent?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&notebookSource=New&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$cloneTeamId%2Fnotes%2Fnotebooks%2F$notebookId&oneNoteWebUrl=$oneNoteWebUrl&notebookName=$notebookDisplayName&ui={locale}`&tenantId={tid}"
                    $removeUrl = "https://www.onenote.com/teams/TabRemove?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D&auth_upn=%7Bupn%7D&notebookSource=New&notebookSelfUrl=https%3A%2F%2Fwww.onenote.com%2Fapi%2Fv1.0%2FmyOrganization%2Fgroups%2F$cloneTeamId%2Fnotes%2Fnotebooks%2F$notebookId&oneNoteWebUrl=$oneNoteWebUrl&notebookName=$notebookDisplayName&ui={locale}`&tenantId={tid}"
                    $websiteUrl = "https://www.onenote.com/teams/TabRedirect?redirectUrl=$oneNoteWebUrl"

                    break
                }
                {@($($global:teamsAppIds.Word),$($global:teamsAppIds.PDF),$($global:teamsAppIds.Excel),$($global:teamsAppIds.PowerPoint)) -contains $_} {
                    
                    $templateTabContentUrl = [System.Web.HttpUtility]::UrlPathEncode($templateTeamChannelTabs[$j].configuration.contentUrl)

                    $contentUrl = [System.Web.HttpUtility]::UrlDecode($templateTabContentUrl.Replace($templateTeamDriveFullUrl, $cloneTeamDriveFullUrl)) # requires unescaped url

                    break
                }
                $($global:teamsAppIds.DocLib) {
                    $templateTabContentUrl = [System.Web.HttpUtility]::UrlPathEncode($templateTeamChannelTabs[$j].configuration.contentUrl)

                    # is doc lib tab pointing to default doc lib of Team
                    $n = $templateTabContentUrl.IndexOf("Shared%20Documents")

                    if($n -eq -1) {
                        # some other doc lib. Need to copy the lib first.
                        $templateDocLibRelativeUrl = Get-ServerRelativeUrl $templateTabContentUrl
                        $templateTeamRootRelativeUrl = $templateTeamDriveUrl.Replace("Shared%20Documents", [string]::Empty)
                        $cloneTeamRootRelativeUrl = $cloneTeamDriveUrl.Replace("Shared%20Documents", [string]::Empty)

                        # TBD: Copying doc lib other than 'Shared Documents'
                        #Copy-TemplateTeamFiles $templateDocLibRelativeUrl $cloneTeamRootRelativeUrl $false
                    }

                    $contentUrl = [System.Web.HttpUtility]::UrlDecode($templateTabContentUrl.Replace($templateTeamRootRelativeUrl, $cloneTeamRootRelativeUrl)) # requires unescaped url

                    break
                }
                default {}
            }
                                
            $cloneTabBody = $null
                                
            switch($tabAction) {
                'Add'{
                    $cloneTabBody = Configure-TabBody $tabAction $entityId $contentUrl $removeUrl $websiteUrl $templateTabDisplayName $templateTabAppId

                    Write-Message "New tab body: $cloneTabBody"

                    Add-TeamChannelTab $cloneTeamId $cloneTeamChannelId $cloneTabBody
                }
                'Update' {
                    $cloneTabBody = Configure-TabBody $tabAction $entityId $contentUrl $removeUrl $websiteUrl ([string]::Empty) $null

                    Write-Message "New tab body: $cloneTabBody"

                    Update-TeamChannelTab $cloneTeamId $cloneTeamChannelId $($cloneTeamChannelTab.id) $cloneTabBody
                }
                default {}
            }

            Write-Message "Done" -BackgroundColor Green
        }

        # delete wiki tabs if any
        if($removeWikiTab) {
            Write-Message "`nGetting 'Wiki' tabs for clone team channel named '$($templateTeamChannels[$i].displayName)'..." -NoNewline

            [System.Array] $cloneTeamChannelTabs = (Get-TeamChannelTabs $cloneTeamId $cloneTeamChannelId $false).value
            Write-Message "Found '$($cloneTeamChannelTabs.count)' total tabs" -ForegroundColor Cyan

            #for($p = 0; $p -le ($cloneTeamChannelTabs.count - 1); $p++) {
            #    Write-Host "writing $($cloneTeamChannelTabs[$p].displayName) $($cloneTeamChannelTabs[$p].teamsAppId)"
            #}
            
            $cloneTeamWikiTabs = $cloneTeamChannelTabs | ?{$_.teamsAppId -eq $($global:teamsAppIds.Wiki)}

            Write-Message "Found '$($cloneTeamWikiTabs.count)' Wiki tabs" -ForegroundColor Cyan

            for($k = 0; $k -le ($cloneTeamWikiTabs.count - 1); $k++) {
                Write-Message "`nRemoving 'Wiki' tab with id '$($cloneTeamWikiTabs[$k].id)'..." -NoNewLine

                Delete-TeamChannelTab -teamId $cloneTeamId -channelId $cloneTeamChannelId -tabId $($cloneTeamWikiTabs[$k].id)

                Write-Message "Done." -BackgroundColor Green
            }
        }
    }
          
    # add owners/members
    if(![string]::IsNullOrEmpty($owners)){
        Add-TeamUsers -users $owners -teamId $cloneTeamId -role Owner
    }

    if(![string]::IsNullOrEmpty($members)){
        Add-TeamUsers -users $members -teamId $cloneTeamId -role Member 
    }

    # remove script runner from owner role
    if($removeCreatorFromTeam) {
        Remove-TeamUser -teamId $cloneTeamId -role Owner -userId $global:myId
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

<# module loader - if using PnP and/or Teams
function CheckAndLoadRequiredModules() {
    $res = [PSCustomObject]@{SPOPnPModuleSuccess = $false;TeamsModuleSuccess = $false;SkypeModuleSuccess = $false}

    if(!(Get-Module -ListAvailable | ? {$_.Name -like "SharePointPnPPowerShellOnline"})) {
        Write-Host "Installing SharePointPnPPowerShellOnline module from https://www.powershellgallery.com/packages/SharePointPnPPowerShellOnline/3.13.1909.0..." -NoNewline

        Install-Module SharePointPnPPowerShellOnline -RequiredVersion 3.13.1909.0 -AllowClobber -Force

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
#>

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

function Wait-Condition {
    [OutputType('void')]
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [scriptblock]$condition,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [int]$checkEvery = 5, ## seconds

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [int]$timeout = 60 ## seconds
    )

    $ErrorActionPreference = 'Stop'

    try {
        $timer = [Diagnostics.Stopwatch]::StartNew()

        # Keep in the loop while the condition is NOT met
        while (-not (& $condition)) {
            Write-Verbose -Message "Waiting for condition '$condition' to be met..."

            # If the timer has waited greater than or equal to the timeout, throw an exception exiting the loop
            if ($timer.Elapsed.TotalSeconds -ge $timeout) {
                throw "Timeout exceeded. Giving up..."
            }

            # Stop the loop every $checkEvery seconds
            Start-Sleep -Seconds $checkEvery
        }
    } 
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    } 
    finally {
        $timer.Stop()
    }
}

function Call-TeamsAsyncEndpoint($endpoint) {
    $asyncOpJSONResponse = $null

    Write-Message "`nEndpoint: $endpoint" -ForegroundColor Magenta

    try {
        $asyncOpJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $endpoint -Method Get -ErrorAction Stop
    }
    catch {
        Write-Message $_ -ForegroundColor Red
    }

    return $asyncOpJSONResponse 
}

function Clone-Team($templateTeamId, $cloneTeamDisplayName, $cloneTeamVisibility, $cloneTeamDescription) {
    $cloneTeamId = [string]::Empty
    $cloneTeamBody = $null

    Write-Host "`n-------------------------------------------------------------------"
    Write-Message "Cloning a new team with name : '$cloneTeamDisplayName', visibility : '$cloneTeamVisibility', description : '$cloneTeamDescription' from a template Team with id '$templateTeamId'..."

    $apiCloneTeamEndpoint = $global:graphAPIv1Endpoint + "/teams/" + $templateTeamId + "/clone"

    if([string]::IsNullOrEmpty($cloneTeamDescription)) {
        $cloneTeamDescription = $cloneTeamDisplayName
    }

    $cloneTeamBody = @{  
            displayName = $cloneTeamDisplayName;
            description = $cloneTeamDescription;
            mailNickname = "msteam_" + [System.Guid]::NewGuid().ToString();        # has to be specified but is ignored
            partsToClone = "channels,tabs,settings,apps";
            visibility = $cloneTeamVisibility
    }
    
    $cloneTeamJSON = $cloneTeamBody | ConvertTo-Json
    
    Write-Message "`nEndpoint: $apiCloneTeamEndpoint" -ForegroundColor Magenta

    try {

        $teamCloneResponse = Invoke-WebRequest -Headers @{Authorization = "Bearer $global:access_token"} -Uri $apiCloneTeamEndpoint -Method Post -ContentType "application/json" -Body $cloneTeamJSON -ErrorAction STOP
    }
    catch {
        Write-Message "Cloning request failed. Exception: $($_.Exception.Message)"
    }

    try {
        # long running async operation returns 202. Retrieve "Location" from response header and call that endpoint periodically to verify that operation succeeded
        $statusCode = $teamCloneResponse.StatusCode
        
        Write-Message "Response status code received: $statusCode" -ForegroundColor Magenta

        if ($statusCode -eq 202) {
            $locationHeader = $teamCloneResponse.Headers.Location
            
            if($locationHeader -ne $null) {
                $apiTeamsAsyncOperationEndPoint = $global:graphAPIv1Endpoint + $locationHeader

                #Write-Message "Pinging '$apiTeamsAsyncOperationEndPoint'..." -ForegroundColor Magenta

                Wait-Condition -condition {(Call-TeamsAsyncEndpoint $apiTeamsAsyncOperationEndPoint).status -eq "succeeded"} -verbose -checkEvery 30 -timeout 600

                #Write-Message "Done" -BackgroundColor Green

                $res = Call-TeamsAsyncEndpoint $apiTeamsAsyncOperationEndPoint

                $cloneTeamId = $res.targetResourceId

                Write-Message "Clone operation succeeded. Result: $res"
            }
        }
    }
    catch {
        Write-Message "Clone operation failed. Exception: $($_.Exception.Message)"
    }

    return $cloneTeamId
}

function Get-O365User {
    param (
        [string] $upn, 
        [switch] $me
    )

    if($me){
        $apiGetUserEndpoint = $global:graphAPIv1Endpoint + "/me"
    }
    else {
        $apiGetUserEndpoint = $global:graphAPIv1Endpoint + "/users/$upn"
    }
        
    Write-Message "`nEndpoint: $apiGetUserEndpoint" -ForegroundColor Magenta

    $userJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetUserEndpoint -Method Get

    return $userJSONResponse
}

function Add-TeamUser($teamId, $userId, $role) {
    $apiAddTeamUserEndpoint = $global:graphAPIv1Endpoint + "/groups/$teamId/members/`$ref"

    switch($role) {
        'Member'{
            $apiAddTeamUserEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/members/`$ref"
        }
        'Owner' {
            $apiAddTeamUserEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/owners/`$ref"
        }
        default {
            $apiAddTeamUserEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/members/`$ref"
        }
    }

    Write-Message "`nEndpoint: $apiAddTeamUserEndpoint" -ForegroundColor Magenta

    Write-Message "Adding '$userId' to '$teamId' as '$role'..." -NoNewline

    $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$userId" } | ConvertTo-Json

    try {

        $userAddJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiAddTeamUserEndpoint -Body $body -Method POST -ErrorAction Stop
        
        Write-Message "Done" -BackgroundColor Green
    }
    catch {
        Write-Message "Failed: $($_.Exception.Message)" -BackgroundColor Red
    }

    return $userAddJSONResponse
}

function Remove-TeamUser($teamId, $userId, $role) {
    $apiRemoveTeamUserEndpoint = $global:graphAPIv1Endpoint + "/groups/$teamId/members/$userId/`$ref"

    switch($role) {
        'Member'{
            $apiRemoveTeamUserEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/members/$userId/`$ref"
        }
        'Owner' {
            $apiRemoveTeamUserEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/owners/$userId/`$ref"
        }
        default {
            $apiRemoveTeamUserEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/members/$userId/`$ref"
        }
    }

    Write-Message "`nEndpoint: $apiRemoveTeamUserEndpoint" -ForegroundColor Magenta

    Write-Message "Removing '$userId' from '$teamId' as '$role'..." -NoNewline

    try {

        $userRemoveJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -Uri $apiRemoveTeamUserEndpoint -Method DELETE -ErrorAction Stop
        
        Write-Message "Done" -BackgroundColor Green
    }
    catch {
        Write-Message "Failed: $($_.Exception.Message)" -BackgroundColor Red
    }

    return $userRemoveJSONResponse
}

function Add-TeamChannel($teamId, $body) {
    $apiAddChannelEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels"

    Write-Message "`nEndpoint: $apiAddChannelEndpoint" -ForegroundColor Magenta

    $channelAddJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiAddChannelEndpoint -Body $body -Method POST

    return $channelUpdateJSONResponse
}

function Update-TeamChannel($teamId, $channelId, $body) {
    $apiUpdateChannelEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels/$channelId"

    Write-Message "`nEndpoint: $apiUpdateChannelEndpoint" -ForegroundColor Magenta

    $channelUpdateJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiUpdateChannelEndpoint -Body $body -Method Patch

    return $channelUpdateJSONResponse
}

function Get-TeamChannels($teamId) {
    # membershipType only available in beta endpoint

    $apiGetChannelsEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels"

    Write-Message "`nEndpoint: $apiGetChannelsEndpoint" -ForegroundColor Magenta

    $channelsJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetChannelsEndpoint -Method Get

    return $channelsJSONResponse | Select-Object Value
}

function Get-TeamChannelIdByName($teamId, $channelDisplayName) {
    $channelId = [string]::Empty

    $apiGetAChannelEndpoint = $global:graphAPIv1Endpoint + "/teams/$teamId/channels?`$filter=displayName eq '$channelDisplayName'"

    Write-Message "`nEndpoint: $apiGetAChannelEndpoint" -ForegroundColor Magenta

    $channelJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetAChannelEndpoint -Method Get

    $val = ($channelJSONResponse | Select-Object Value).value

    if ($val -ne $null) {
        if($val.length -gt 0) {
            $channelId = $val[0].id
        }
    }
    
    return $channelId
}

function Get-TeamChannelTabs($teamId, $channelId, $excludeWikiTab) {
    # needs beta endpoint to return teamsAppId on tab
    $apiGetChannelTabsEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels/$channelId/tabs"

    if($excludeWikiTab) {
        $apiGetChannelTabsEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels/$channelId/tabs?`$filter=displayName ne 'Wiki'"
    }

    Write-Message "`nEndpoint: $apiGetChannelTabsEndpoint" -ForegroundColor Magenta

    $channelTabsJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetChannelTabsEndpoint -Method Get

    return $channelTabsJSONResponse | Select-Object Value
}

function Get-TeamChannelTabByAppIdAndName($teamId, $channelId, $appId, $tabDisplayName) {
    $channelTab = $null

    $encodedTabDisplayName = [System.Web.HttpUtility]::UrlEncode($tabDisplayName)

    # filter clause requires beta endpoint
    $apiGetChannelTabEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels/$channelId/tabs?`$filter=teamsAppId eq '$appId' and displayName eq '$encodedTabDisplayName'"

    Write-Message "`nEndpoint: $apiGetChannelTabEndpoint" -ForegroundColor Magenta

    $channelTabJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetChannelTabEndpoint -Method Get

    $val = ($channelTabJSONResponse | Select-Object Value).value

    if($val -ne $null) {
        if($val.length -gt 0) {
            $channelTab = $val[0]
        }
    }

    return $channelTab
}

function Add-TeamChannelTab($teamId, $channelId, $body) {
    $apiAddChannelTabEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels/$channelId/tabs"

    Write-Message "`nEndpoint: $apiAddChannelTabEndpoint" -ForegroundColor Magenta

    $channelTabAddJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiAddChannelTabEndpoint -Body $body -Method POST

    return $channelTabAddJSONResponse
}

function Update-TeamChannelTab($teamId, $channelId, $tabId, $body) {
    $apiUpdateChannelTabEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels/$channelId/tabs/$tabId"

    Write-Message "`nEndpoint: $apiUpdateChannelTabEndpoint" -ForegroundColor Magenta

    $channelTabUpdateJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiUpdateChannelTabEndpoint -Body $body -Method Patch

    return $channelTabUpdateJSONResponse
}

function Delete-TeamChannelTab($teamId, $channelId, $tabId) {
    $apiDeleteChannelTabEndpoint = $global:graphAPIBetaEndpoint + "/teams/$teamId/channels/$channelId/tabs/$tabId"

    Write-Message "`nEndpoint: $apiDeleteChannelTabEndpoint" -ForegroundColor Magenta

    $channelTabDeleteJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiDeleteChannelTabEndpoint -Method Delete

    return $channelTabDeleteJSONResponse
}

function Configure-TabBody($act, $entityId, $contentUrl, $removeUrl, $websiteUrl, $displayName, $teamsAppId) {
    $body = $null

    $configObj = @{ 
                entityId = $entityId; 
                contentUrl = $contentUrl;
                removeUrl = $removeUrl;
                websiteUrl = $websiteUrl
            }
    
    switch($act) {
        'Add' {
            $body = @{ displayName = $displayName; "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/$teamsAppId"; configuration = $configObj }
        }
        'Update' {
            if(![string]::IsNullOrWhiteSpace($displayName)) {
                $body = @{ displayName = $displayName; configuration = $configObj }
            }
            else {
                $body = @{ configuration = $configObj }
            }
        }
        default {}
    }

    return $body | ConvertTo-Json
}

function Get-TeamDefaultNotebook($teamId) {
    Write-Message "`nRetrieving default notebook for Team '$teamId'..." -NoNewline
    
    $defaultNotebook = $null

    $apiGetNotebookEndpoint = $global:graphAPIv1Endpoint + "/groups/$teamId/onenote/notebooks"

    Write-Message "`nEndpoint: $apiGetNotebookEndpoint" -ForegroundColor Magenta

    $notebooksJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetNotebookEndpoint -Method Get

    $val = ($notebooksJSONResponse | Select-Object Value).value
    
    if($val -ne $null) {
        if($val.length -gt 0) {
            $defaultNotebook = $val[0]
        }
    }

    Write-Message "Done" -BackgroundColor Green

    return $defaultNotebook
}

function Add-ChannelFolderDriveItem($teamId, $channelDisplayName) {
    Write-Message "`nProvisioning drive item for channel folder '$channelDisplayName' in Team '$teamId' drive..." -NoNewLine

    $driveItemId = [string]::Empty

    $apiPostChannelDriveItemEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/drive/root/children"
    
    $body = ConvertTo-Json @{ name = $channelDisplayName; folder = @{}; "@microsoft.graph.conflictBehavior" = "rename" }

    Write-Message "`nEndpoint: $apiPostChannelDriveItemEndpoint" -ForegroundColor Magenta

    try {
        $addChannelDriveItemJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiPostChannelDriveItemEndpoint -Method POST -Body $body -ErrorAction Stop
        
        $driveItemId = $addChannelDriveItemJSONResponse.id
    }
    catch {
       $statusCode = $_.Exception.Response.StatusCode.value__
       
       Write-Message "Exception: $($_.Exception.Message) StatusCode: $statusCode" -ForegroundColor Red
    }

    Write-Message "Done. Drive item id: $driveItemId" -BackgroundColor Green

    return $driveItemId
}

function Get-ChannelFolderDriveItemId($teamId, $channelDisplayName) {
    Write-Message "`nRetrieving drive item id for channel folder named '$channelDisplayName' in Team '$teamId'..." -NoNewLine

    $driveItemId = [string]::Empty

    $encodedchannelDisplayName = [System.Web.HttpUtility]::UrlPathEncode($channelDisplayName)

    $apiGetChannelDriveItemEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/drive/root:/" + $encodedchannelDisplayName
    
    Write-Message "`nEndpoint: $apiGetChannelDriveItemEndpoint" -ForegroundColor Magenta

    try {
        $getChannelDriveItemJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetChannelDriveItemEndpoint -Method Get -ErrorAction Stop
        $driveItemId = $getChannelDriveItemJSONResponse.id
    }
    catch {
       $statusCode = $_.Exception.Response.StatusCode.value__
       
       Write-Message "Exception: $($_.Exception.Message) StatusCode: $statusCode" -ForegroundColor Red
       
       if($statusCode -eq "404") {
            # channel folder is NOT created until you click the "Files" tab. Mitigate by adding folder with exact same name.
            $driveItemId = Add-ChannelFolderDriveItem $teamId $channelDisplayName
       }
    }

    Write-Message "Done. Drive item id: $driveItemId" -BackgroundColor Green

    return $driveItemId
}

function Get-ChannelFiles($teamId, $channelDisplayName){
    Write-Message "`nRetrieving all files and folders in channel '$channelDisplayName' for Team '$teamId'..." -NoNewLine

    $encodedchannelDisplayName = [System.Web.HttpUtility]::UrlPathEncode($channelDisplayName)

    $apiGetChannelFilesEndpoint = $global:graphAPIBetaEndpoint + "/groups/$teamId/drive/root:/" + $encodedchannelDisplayName + ":/children"

    Write-Message "`nEndpoint: $apiGetChannelFilesEndpoint" -ForegroundColor Magenta

    $getChannelFilesJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"; Accept = "application/json"} -Uri $apiGetChannelFilesEndpoint -Method Get
    
    $val = ($getChannelFilesJSONResponse | Select-Object Value).value

    Write-Message "Done" -BackgroundColor Green

    return $val
}

function Copy-ChannelFiles($sourceDriveId, $sourceDriveItemId, $targetDriveId, $targetDriveItemId) {
    Write-Message "`nCopying from drive '$sourceDriveId', file/folder '$sourceDriveItemId' to drive '$targetDriveId' folder '$targetDriveItemId'..." -NoNewLine

    $apiCopyChannelFilesEndpoint = $global:graphAPIBetaEndpoint + "/drives/$sourceDriveId/items/$sourceDriveItemId/copy"

    $parentRef = @{ 
            driveId = $targetDriveId; 
            id = $targetDriveItemId
        }

    $body = @{ parentReference = $parentRef } | ConvertTo-Json

    Write-Message "`nEndpoint: $apiCopyChannelFilesEndpoint" -ForegroundColor Magenta

    $copyChannelFilesJSONResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -ContentType "application/json" -Uri $apiCopyChannelFilesEndpoint -Body $body -Method POST

    Write-Message "Done" -BackgroundColor Green

    return $copyChannelFilesJSONResponse
}

function Get-TeamDefaultDriveUrl($teamId, $serverRelativeUrl) {
    Write-Host "`n--------------------------------------------------------------------------"
    Write-Message "Retrieving default root drive (Shared Documents library path) for Team '$teamId'..."

    $res = [PSCustomObject]@{DriveUrl = [string]::Empty;DriveId = [string]::Empty}

    $apiGetTeamDefaultDriveEndpoint = $global:graphAPIv1Endpoint + "/groups/$teamId/drive"
    
    Wait-Condition -checkEvery 30 -timeout 180 -condition {(Call-TeamsAsyncEndpoint $apiGetTeamDefaultDriveEndpoint).webUrl -ne $null} -verbose

    $defaultDriveJSONResponse = Call-TeamsAsyncEndpoint $apiGetTeamDefaultDriveEndpoint

    if($defaultDriveJSONResponse -ne $null) {
        $res.DriveId = $defaultDriveJSONResponse.id

        if($serverRelativeUrl) {
            $res.DriveUrl = Get-ServerRelativeUrl($defaultDriveJSONResponse.webUrl)
        }
        else {
            $res.DriveUrl = $defaultDriveJSONResponse.webUrl
        }
    }

    Write-Message "Done. Default Drive Url: $($res.DriveUrl), Id: $($res.DriveId)" -BackgroundColor Green

    return $res
}

function Get-ServerRelativeUrl($fullUrl) {
    [RegEx] $serverUrlPattern = '([a-zA-Z]{3,})://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?'

    $serverRelativeUrl = [string]::Empty

    if(![string]::IsNullOrWhiteSpace($fullUrl)) {
        $serverRelativeUrl = $serverUrlPattern.Replace($fullUrl, [string]::Empty)
    }

    return $serverRelativeUrl
}

<# if using PnP
function Copy-TemplateTeamFiles($sourceFolderUrl, $targetFolderUrl, $skipSourceFolder) {

    Write-Message "Copying files from '$sourceFolderUrl' to '$targetFolderUrl'..." -NoNewline

    $optionalParams = @{}

    if($skipSourceFolder -ne $null) {
        $optionalParams.Add("SkipSourceFolderName", $skipSourceFolder)
    }

    Copy-PnPFile -SourceUrl $sourceFolderUrl -TargetUrl $targetFolderUrl -OverwriteIfAlreadyExists @optionalParams -force | Out-Null

    Write-Message "Done" -BackgroundColor Green
}
#>

function Get-TeamsByDisplayName($teamName, $returnFirst) {
    $team = $null

    # Need beta endpoint for the filter expression
    $apiGetTeamsEndpoint = $global:graphAPIBetaEndpoint + "/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '$teamName'"

    Write-Message "`nEndpoint: $apiGetTeamsEndpoint" -ForegroundColor Magenta

    $getTeamsResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -Uri $apiGetTeamsEndpoint -Method Get

    $allTeams = ($getTeamsResponse | select-object Value).Value | Select-Object id, displayName, visibility, mailNickName

    $team = $allTeams

    $cnt = $allTeams.count

    if($cnt -gt 0) {
        if($returnFirst) {
            Write-Message "Found $cnt Teams with name '$teamName'. Returning first one." -ForegrounColor Red

            $team = $allTeams[0]
        }
    }

    return $team
}

function Get-AllTeams() {
    Write-Host "`n--------------------------------------------------------------------------"
    Write-Message "Retrieving all Teams in the '$global:tenantRootUrl' tenant..." -NoNewLine

    # Need beta endpoint for the filter expression
    $apiGetAllTeamsEndpoint = $global:graphAPIBetaEndpoint + "/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"

    Write-Message "`nEndpoint: $apiGetAllTeamsEndpoint" -ForegroundColor Magenta

    $getAllTeamsResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $global:access_token"} -Uri $apiGetAllTeamsEndpoint -Method Get

    $allTeams = ($getAllTeamsResponse | select-object Value).Value | Select-Object id, displayName, visibility, mailNickName

    Write-Message "Done" -BackgroundColor Green

    return $allTeams
}

function Get-AuthorizationCode() {
    Add-Type -AssemblyName System.Windows.Forms

    Write-Host "`n--------------------------------------------------------------------------"
    Write-Message "Get authentication code..." -NoNewLine

    $authCodeEndpoint = "https://login.microsoftonline.com/$global:tenantIDEncoded/oauth2/authorize?response_type=code&redirect_uri=$global:redirectUriEncoded&client_id=$global:clientID&scope=$global:scopeEncoded"

    Write-Message "`nEndpoint: $authCodeEndpoint" -ForegroundColor Magenta

    $form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
    $web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($authCodeEndpoint -f ($scope -join "%20")) }
    
    # execute inline script block on sign in completion
    $DocComp  = {
        $global:uri = $web.Url.AbsoluteUri        
        if ($global:uri -match "error=[^&]*|code=[^&]*") { $form.Close() }
    }

    $web.ScriptErrorsSuppressed = $true
    $web.Add_DocumentCompleted($DocComp)
    $form.Controls.Add($web)
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null

    $queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)

    $output = @{}
    foreach($key in $queryOutput.Keys){
        $output["$key"] = $queryOutput[$key]
    }

    $output

    # Extract authorization code from the returned URI
    $regex = '(?<=code=)(.*)(?=&)'
    $global:authCode = ($global:uri | Select-string -pattern $regex).Matches[0].Value

    Write-Message "Done" -BackgroundColor Green
}

function Get-AccessToken() {
    Write-Host "`n--------------------------------------------------------------------------"
    Write-Message "Get accesss token..." -NoNewLine

    # For access token following query string needs to go in request body
    $body = "grant_type=authorization_code&redirect_uri=$global:redirectUri&client_id=$global:clientId&client_secret=$global:clientSecretEncoded&code=$global:authCode&resource=$global:resource"

    $accessTokenEndpoint = "https://login.microsoftonline.com/$global:tenantIDEncoded/oauth2/token"

    Write-Message "`nEndpoint: $accessTokenEndpoint" -ForegroundColor Magenta

    $tokenResponse = Invoke-RestMethod $accessTokenEndpoint `
                    -Method Post -ContentType "application/x-www-form-urlencoded" `
                    -Body $body `
                    -ErrorAction STOP
    
    Write-Message "Done" -BackgroundColor Green

    return $tokenResponse.access_token
}

#------------------ main script --------------------------------------------------

Write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

$global:graphAPIv1Endpoint = "https://graph.microsoft.com/v1.0"
$global:graphAPIBetaEndpoint = "https://graph.microsoft.com/beta"

$global:teamsAppIds = @{
    Website = "com.microsoft.teamspace.tab.web"; # website
    Planner = "com.microsoft.teamspace.tab.planner"; # planner
    Stream = "com.microsoftstream.embed.skypeteamstab"; # stream
    Forms = "81fef3a6-72aa-4648-a763-de824aeafb7d"; # forms
    Word = "com.microsoft.teamspace.tab.file.staticviewer.word"; # word
    Excel = "com.microsoft.teamspace.tab.file.staticviewer.excel"; # excel
    PowerPoint = "com.microsoft.teamspace.tab.file.staticviewer.powerpoint"; # powerpoint
    Pdf = "com.microsoft.teamspace.tab.file.staticviewer.pdf"; # pdf
    Wiki = "com.microsoft.teamspace.tab.wiki"; # wiki
    DocLib = "com.microsoft.teamspace.tab.files.sharepoint"; # sharepoint document library
    OneNote = "0d820ecd-def2-4297-adad-78056cde7c78"; # oneNote
    PowerBI = "com.microsoft.teamspace.tab.powerbi"; # power BI
    SPOPage = "2a527703-1f6f-4559-a332-d8a7d288cd88" # sharepoint page or list
}

# Root resource URI
$global:resource = "https://graph.microsoft.com"

$psVersionCheck = CheckPowerShellVersion

$psExecutionPolicyLevelCheck = CheckExecutionPolicy 

<# if using Teams module or PnP
#$modulesCheck = CheckAndLoadRequiredModules

#if($psVersionCheck.Success -and $psExecutionPolicyLevelCheck.Success -and $modulesCheck.TeamsModuleSuccess -and $modulesCheck.SkypeModuleSuccess -and $modulesCheck.SPOPnPModuleSuccess) {
#>

if($psVersionCheck.Success -and $psExecutionPolicyLevelCheck.Success) {    
    # Root tenant url
    if([string]::IsNullOrWhiteSpace($tenantRootUrl))
    {
        do {
            $tenantRootUrl = Read-Host "Specify the tenant root url https://tenantname.sharepoint.com"
        }
        until (![string]::IsNullOrWhiteSpace($tenantRootUrl))
    }
    $global:tenantRootUrl = $tenantRootUrl

    # Attributes of enterprise app that exposes the graph API
    if([string]::IsNullOrWhiteSpace($tenantId))
    {
        do {
            $tenantId = Read-Host "Specify the tenant id for the Azure Enterprise app through which to access Graph API"
        }
        until (![string]::IsNullOrWhiteSpace($tenantId))
    }
    $global:tenantId = $tenantId

    if([string]::IsNullOrWhiteSpace($clientId))
    {
        do {
            $clientId = Read-Host "Specify the client id for the Azure Enterprise app through which to access Graph API"
        }
        until (![string]::IsNullOrWhiteSpace($clientId))
    }
    $global:clientid = $clientId

    if([string]::IsNullOrWhiteSpace($clientSecret))
    {
        do {
            $clientSecret = Read-Host "Specify the client secret for the Azure Enterprise app through which to access Graph API"
        }
        until (![string]::IsNullOrWhiteSpace($clientSecret))
    }
    $global:clientSecret = $clientSecret

    if([string]::IsNullOrWhiteSpace($redirectUri))
    {
        do {
            $redirectUri = Read-Host "Specify the redirect url for the Azure Enterprise app through which to access Graph API"
        }
        until (![string]::IsNullOrWhiteSpace($redirectUri))
    }
    $global:redirectUri = $redirectUri

    # UrlEncode the attributes 
    Add-Type -AssemblyName System.Web

    $global:tenantIDEncoded = [System.Web.HttpUtility]::UrlEncode($global:tenantId)
    $global:clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($global:clientid)
    $global:clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($global:clientSecret)
    $global:redirectUriEncoded = [System.Web.HttpUtility]::UrlEncode($global:redirectUri)
    $global:resourceEncoded = [System.Web.HttpUtility]::UrlEncode($global:resource)
    $global:scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com/.") # default acquire all API permissions exposed by app

    <# if using PnP or Teams Modules
    #$global:cred = Get-Credential -Message "Please enter your organizational credential for Office 365"
    #$global:currentUser = $global:cred.UserName

    
    # Connect Microsoft Teams
    Connect-MicrosoftTeams -Credential $global:cred

    # Connect PnP shell
    #Connect-PnPOnline -Url $global:tenantRootUrl -Credentials $global:cred
    #>

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

    $global:logCSVPath = $global:logFolderPath + "QuickLog.csv"
    $global:logFilePath = $global:logFolderPath + "ActionLog.log"

    #log csv
    "TeamName `t Type `t TeamId `t TemplateTeamName `t TemplateTeamId `t Owners `t Members `t Date `t Action" | out-file $global:logCSVPath

    Write-Output "Logging started - $(Get-Date)`n" | out-file $global:logFilePath

    Get-AuthorizationCode

    if([string]::IsNullOrEmpty($global:authCode)) {
        Write-Message -msg "Authentication code could not be retrieved for TenantId: '$global:tenantId', ClientId: '$global:clientId', ClientSecret: '$global:clientSecret'..." -NoNewline
        Write-Message -msg "Quitting" -BackgroundColor Red

        return
    }

    $global:access_token = Get-AccessToken

    Write-Message "`nAccess Token: $global:access_token" -ForegroundColor Cyan

    Write-Message "Getting Office 365 user id for logged in user..." -NoNewLine

    $me = Get-O365User -me

    if($me -ne $null) {
        $global:myId = $me.id
        Write-Message "UPN: '$($me.userPrincipalName)' id:'$global:myId'..." -NoNewLine
    }

    Write-Message "Done" -BackgroundColor Green

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
            
            if([string]::IsNullOrWhiteSpace($teamDisplayName))
            {
                do {
                    $teamDisplayName = Read-Host "Specify the display name for the Team"
                }
                until (![string]::IsNullOrWhiteSpace($teamDisplayName))
            }

            if([string]::IsNullOrWhiteSpace($templateTeamName))
            {
                do {
                    $templateTeamName = Read-Host "Specify the display name of the Template Team to clone"
                }
                until (![string]::IsNullOrWhiteSpace($templateTeamName))
            }

            $row = @{
                    TeamDisplayName=$teamDisplayName;
                    TeamDescription=$teamDescription;
                    TeamPrivacy=$teamPrivacy;
                    TemplateTeamName=$templateTeamName;
                    OwnerUPNs=$owners;
                    MemberUPNs=$members
                }
    
            Action-ATeam $row | Out-Null
        }
    }

    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append

    #Disconnect-PnPOnline

    #Disconnect-MicrosoftTeams
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow