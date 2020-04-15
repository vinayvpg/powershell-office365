<#
 .NOTES
 ===========================================================================
 Created On:   8/1/2019
 Author:       Vinay Prabhugaonkar
 E-Mail:       vinay.prabhugaonkar@sparkhound.com
 Filename:     Update-Team.ps1
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
    Bulk or individual update of a Microsoft Team in Office 365.

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
    Update. Default is empty string which means no action.

.USAGE 
    Bulk update Teams specified in csv in a single batch operation
     
    PS >  Update-Team.ps1 -inputCSVPath "c:/temp/teams.csv" -action "Update"
.USAGE 
    Update individual Team from existing O365 group
     
    PS >  Update-Team.ps1 -groupId 'xxxxxxx-xxxx-xxxx-xxxxxxxx' -action "Update"
.USAGE 
    Update individual Team with specific parameters
     
    PS >  Update-Team.ps1 -teamDisplayName 'My team' -teamPrivacy "Public" -owners "abc@contoso.com;xyz@contoso.com" -members "joe@contoso.com;jane@contoso.com" -channels "channel 1;channel 2" -action "Update"
#>


[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv containing information on Teams to be updated")]
    [string] $inputCSVPath,
    
    [Parameter(Mandatory=$false, Position=1, HelpMessage="Path to folder where log will be created")]
    [string] $outputLogFolderPath,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="O365 Group Id")]
    [string] $groupId,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="Team display name")]
    [string] $teamDisplayName="Murphy Canada",

    [Parameter(Mandatory=$false, Position=4, HelpMessage="Team description")]
    [string] $teamDescription="Murphy Canada Team. This is a private team for collaboration amongst Murphy's Canada employees only.",

    [Parameter(Mandatory=$false, Position=5, HelpMessage="Team mail nickname. Must be unique within the tenant.")]
    [string] $teamMailNickName,

    [Parameter(Mandatory=$false, Position=6, HelpMessage="Team privacy setting")]
    [ValidateSet('Private','Public','HiddenMembership')][string] $teamPrivacy,

    [Parameter(Mandatory=$false, Position=7, HelpMessage="Owners (semicolon separated list of UPNs)")]
    [string] $owners="sean_petersen@murphyoilcorp.com;Kathy_Pfaffinger@murphyoilcorp.com",

    [Parameter(Mandatory=$false, Position=8, HelpMessage="Members (semicolon separated list of UPNs)")]
    [string] $members="Coralee_Cryne@murphyoilcorp.com;Dennis_Enyedy@murphyoilcorp.com;Joy_Mortenson@murphyoilcorp.com;Dragan_Kalajdzic@murphyoilcorp.com;Kelven_Pang@murphyoilcorp.com;lily_wang@murphyoilcorp.com;Kathy_Pfaffinger@murphyoilcorp.com;Kevin_Kienzle@contractor.murphyoilcorp.com;Les_Lambert@murphyoilcorp.com;Helen_Holder@murphyoilcorp.com;Mark_Friesen@murphyoilcorp.com;Dejan_Timotijevic@murphyoilcorp.com;Adam_McConnell@murphyoilcorp.com;Jennifer_Wilkinson@murphyoilcorp.com;Craig_Sinclair@murphyoilcorp.com;Natasha_Khalil@contractor.murphyoilcorp.com;Lama_Amro@murphyoilcorp.com;Chris_Flannery@murphyoilcorp.com;Jill_Aspin@murphyoilcorp.com;Robert_Bak@murphyoilcorp.com;Jason_Baird@murphyoilcorp.com;Dave_Elmer@murphyoilcorp.com;Michael_Bartoszewski@murphyoilcorp.com;Sarah_Li@murphyoilcorp.com;laura_naaykens@murphyoilcorp.com;Adam_Bryson@murphyoilcorp.com;chris_phone@contractor.murphyoilcorp.com;kelly_mcmullen@murphyoilcorp.com;greg_blayone@contractor.murphyoilcorp.com;Jesse_Frederick@contractor.murphyoilcorp.com;dennis_dubiel@contractor.murphyoilcorp.com;Millen_Knuuttila@contractor.murphyoilcorp.com;Cobus_duPlessis@contractor.murphyoilcorp.com;Bill_Winter@contractor.murphyoilcorp.com;corey_laas@murphyoilcorp.com;Dale_Dubrule@contractor.murphyoilcorp.com;Jorge_Osycka@murphyoilcorp.com;adam_shultz@contractor.murphyoilcorp.com;John_Farion@contractor.murphyoilcorp.com;michelle_mattie@murphyoilcorp.com;Steve_Porter@contractor.murphyoilcorp.com;bill_williams@contractor.murphyoilcorp.com;Darius_Wurz@contractor.murphyoilcorp.com;Craig_Garland@murphyoilcorp.com;nadine_yuen@contractor.murphyoilcorp.com;brenda_ison@contractor.murphyoilcorp.com;miroslav_kurcik@contractor.murphyoilcorp.com;deon_guillet@murphyoilcorp.com;Derek_Greene@murphyoilcorp.com;Cynthia_Moises@contractor.murphyoilcorp.com;James_Scott@murphyoilcorp.com;Rodney_Patterson@contractor.murphyoilcorp.com;Gaye_Marshall@murphyoilcorp.com;Katie_vanKampen@murphyoilcorp.com;Ryan_Gillen@murphyoilcorp.com;Arlene_Gibb@murphyoilcorp.com;Erika_Neaves@murphyoilcorp.com;Randy_Baker@contractor.murphyoilcorp.com;Scott_Restoule@murphyoilcorp.com;Justin_Wong@murphyoilcorp.com;Amber_Poole@murphyoilcorp.com;jennifer_mariani@murphyoilcorp.com;Jeff_Good@murphyoilcorp.com;Joseph_Manalo@murphyoilcorp.com;Dallas_Turcotte@murphyoilcorp.com;Tim_Call@contractor.murphyoilcorp.com;rob_clunk@murphyoilcorp.com;Simon_Janzen@contractor.murphyoilcorp.com;Shannon_Thiessen@contractor.murphyoilcorp.com;aaron_bull@contractor.murphyoilcorp.com;Sean_Ryder@murphyoilcorp.com;Shane_Kremsater@murphyoilcorp.com;Heather_StMartin@murphyoilcorp.com;Mario_Santerre@murphyoilcorp.com;Jay_Thebeau@murphyoilcorp.com;Dwayne_Davie@contractor.murphyoilcorp.com;Cody_Lightbody@contractor.murphyoilcorp.com;Dionne_Paul@murphyoilcorp.com;Sarah_MacFarlane@murphyoilcorp.com;Ryan_Hasell@murphyoilcorp.com;Brent_Collyer@murphyoilcorp.com;laura_metcalfe@contractor.murphyoilcorp.com;Dean_Eastly@contractor.murphyoilcorp.com;Jarred_Anstett@murphyoilcorp.com;Jason_Smith@murphyoilcorp.com;Logan_Barclay@contractor.murphyoilcorp.com;jacquie_mccarroll@contractor.murphyoilcorp.com;Tim_Bergen@contractor.murphyoilcorp.com;Dale_Kramer@contractor.murphyoilcorp.com;Conner_Krieger@contractor.murphyoilcorp.com;Bret_McGhie@contractor.murphyoilcorp.com;Shyam_Sundar@contractor.murphyoilcorp.com;Andrew_Cook@murphyoilcorp.com;david_will@murphyoilcorp.com;Gihan_Rajapaksha@murphyoilcorp.com;travis_tiffin@contractor.murphyoilcorp.com;wacey_arthur@contractor.murphyoilcorp.com;Lorne_Ewert@contractor.murphyoilcorp.com;Trevor_Want@contractor.murphyoilcorp.com;Colin_Krieger@contractor.murphyoilcorp.com;Curtis_Szafron@murphyoilcorp.com;Gerry_Kienzle@murphyoilcorp.com;Debbi_McDonald@murphyoilcorp.com;Logan_Marlow@murphyoilcorp.com;Howard_Ly@murphyoilcorp.com;MyAnh_Chau@murphyoilcorp.com;Ben_Duffy@murphyoilcorp.com;jacqueline_james@murphyoilcorp.com;Nick_Choy@murphyoilcorp.com;Jas_Paul@murphyoilcorp.com;Daniel_Sarrosa@murphyoilcorp.com;Brian_Fong@murphyoilcorp.com;doug_loewen@murphyoilcorp.com;Jamila_Mahjor@murphyoilcorp.com;lynne_myron@murphyoilcorp.com;Brian_Ho@murphyoilcorp.com;ken_sturgeon@contractor.murphyoilcorp.com;doug_frith@contractor.murphyoilcorp.com;Rob_Ismond@murphyoilcorp.com;Shara_McFarland@murphyoilcorp.com;Matt_Desroches@murphyoilcorp.com;Craig_Zenner@contractor.murphyoilcorp.com;Eric_Sultanian@murphyoilcorp.com;Thanh_Hwang@murphyoilcorp.com;Rob_Nelson@murphyoilcorp.com;Jordan_Klebanowski@murphyoilcorp.com;Murray_Roth@murphyoilcorp.com;Chris_Young@murphyoilcorp.com;Judy_Musselman@contractor.murphyoilcorp.com;Aaron_Cho@murphyoilcorp.com;Jason_Hogberg@contractor.murphyoilcorp.com;angela_toth@murphyoilcorp.com;Scott_Levins@murphyoilcorp.com;michael_delesky@murphyoilcorp.com;Roman_Nelson@contractor.murphyoilcorp.com;travis_jessen@contractor.murphyoilcorp.com;Jeff_Rayner@contractor.murphyoilcorp.com;George_Katsimihas@murphyoilcorp.com;Jennifer_Kha@murphyoilcorp.com;Tyson_Trail@murphyoilcorp.com;Mike_Pacholek@murphyoilcorp.com;Melanie_Currie@murphyoilcorp.com;Suzy_Chen@murphyoilcorp.com;Charlene_Zhang@murphyoilcorp.com;joanna_ng@murphyoilcorp.com;Kelly_Szautner@murphyoilcorp.com;Christiane_Martin@murphyoilcorp.com;sop_kaybob@contractor.murphyoilcorp.com;Grant_Mclean@murphyoilcorp.com;sop_montney@contractor.murphyoilcorp.com;Lulu_Chen@murphyoilcorp.com;Cheryl_Alhashwa@murphyoilcorp.com;John_Michaud@murphyoilcorp.com;Hayley_Sartin@murphyoilcorp.com;ashley_trieu@murphyoilcorp.com;Dan_Barnard@contractor.murphyoilcorp.com;Nate_Kreiger@murphyoilcorp.com;Evan_Boire@murphyoilcorp.com;Ransis_Kais@murphyoilcorp.com;Katrina_Chambers@contractor.murphyoilcorp.com;bryce_nicholson@murphyoilcorp.com;Chance_Rich@murphyoilcorp.com;Rob_Lanctot@murphyoilcorp.com;Ai_Chee@murphyoilcorp.com;Todd_Tarala@murphyoilcorp.com;Calgary_Reception@murphyoilcorp.com;Anna_Steininger@murphyoilcorp.com;andrea_vigueras@murphyoilcorp.com;Blaine_Ham@murphyoilcorp.com;Stephanie_Neilson@murphyoilcorp.com;Tyler_Heffernan@murphyoilcorp.com;Cameron_Prefontaine@murphyoilcorp.com;Brett_Frostad@murphyoilcorp.com;Nathalie_Stock@murphyoilcorp.com;Richard_Dunn@murphyoilcorp.com;Joanna_Clarke@murphyoilcorp.com;Caley_Millar@murphyoilcorp.com;Carolyn_Murphy@murphyoilcorp.com;myron_hirniak@contractor.murphyoilcorp.com;Al_Frandsen@contractor.murphyoilcorp.com;Troy_Morrison@contractor.murphyoilcorp.com;Chris_Ragot@contractor.murphyoilcorp.com;Darryl_Snethun@contractor.murphyoilcorp.com;Peter_Lewis@contractor.murphyoilcorp.com;Tom_Whitford@contractor.murphyoilcorp.com;jeff_larson@contractor.murphyoilcorp.com;kevin_swain@contractor.murphyoilcorp.com;Todd_Margareeth@contractor.murphyoilcorp.com;waireka_morris@murphyoilcorp.com;dan_meier@murphyoilcorp.com;Pete_Magnuson@contractor.murphyoilcorp.com;Conrad_Kilba@contractor.murphyoilcorp.com;Adam_Johnson@murphyoilcorp.com;Elie_Meyer@murphyoilcorp.com;CDN_DORA@MurphyOilCorp.com;Chrystal_Matthews@murphyoilcorp.com;Madison_Cruse@contractor.murphyoilcorp.com;Kenny_Yong@murphyoilcorp.com;lana_makonin@contractor.murphyoilcorp.com;grace_poletto@contractor.murphyoilcorp.com;Daniel_Mora@murphyoilcorp.com;Chantale_Wold@murphyoilcorp.com;Filip_Burcevski@murphyoilcorp.com;Mark_Leger@contractor.murphyoilcorp.com;Alessandro_Vena@murphyoilcorp.com;Steven_Kutarna@contractor.murphyoilcorp.com;Rob_Rattie@murphyoilcorp.com;Gilbert_Cheng@contractor.murphyoilcorp.com",

    [Parameter(Mandatory=$false, Position=9, HelpMessage="Channels to provision (semicolon separated list)")]
    [string] $channels,

    [Parameter(Mandatory=$false, Position=10, HelpMessage="Team setting")]
    [boolean] $allowGiphy,

    [Parameter(Mandatory=$false, Position=11, HelpMessage="Team setting")]
    [string] $giphyContentRating,

    [Parameter(Mandatory=$false, Position=12, HelpMessage="Team setting")]
    [boolean] $allowStickersAndMemes,
        
    [Parameter(Mandatory=$false, Position=13, HelpMessage="Team setting")]
    [boolean] $allowCustomMemes,
        
    [Parameter(Mandatory=$false, Position=14, HelpMessage="Team setting")]
    [boolean] $allowGuestCreateUpdateChannels ,
        
    [Parameter(Mandatory=$false, Position=15, HelpMessage="Team setting")]
    [boolean] $allowGuestDeleteChannels,
        
    [Parameter(Mandatory=$false, Position=16, HelpMessage="Team setting")]
    [boolean] $allowCreateUpdateChannels=$false,
        
    [Parameter(Mandatory=$false, Position=17, HelpMessage="Team setting")]
    [boolean] $allowDeleteChannels=$false,
        
    [Parameter(Mandatory=$false, Position=18, HelpMessage="Team setting")]
    [boolean] $allowAddRemoveApps,
        
    [Parameter(Mandatory=$false, Position=19, HelpMessage="Team setting")]
    [boolean] $allowCreateUpdateRemoveTabs,
        
    [Parameter(Mandatory=$false, Position=20, HelpMessage="Team setting")]
    [boolean] $allowCreateUpdateRemoveConnectors,
        
    [Parameter(Mandatory=$false, Position=21, HelpMessage="Team setting")]
    [boolean] $allowUserEditMessages,
        
    [Parameter(Mandatory=$false, Position=22, HelpMessage="Team setting")]
    [boolean] $allowUserDeleteMessages,
        
    [Parameter(Mandatory=$false, Position=23, HelpMessage="Team setting")]
    [boolean] $allowOwnerDeleteMessages,
        
    [Parameter(Mandatory=$false, Position=24, HelpMessage="Team setting")]
    [boolean] $allowTeamMentions,
        
    [Parameter(Mandatory=$false, Position=25, HelpMessage="Team setting")]
    [boolean] $allowChannelMentions,

    [Parameter(Mandatory=$false, Position=26, HelpMessage="Team setting")]
    [boolean] $showInTeamsSearchAndSuggestions,

    [Parameter(Mandatory=$false, Position=27, HelpMessage="Action to take")]
    [ValidateSet('','Update')] [string] $action = "Update"
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
             $users,$groupId,$role
          )   
    Process
    {
        $teamusers = $users -split ";" 
        if($teamusers)
        {
            for($j =0; $j -le ($teamusers.count - 1) ; $j++)
            {
                Write-Host "---> Adding '$($teamusers[$j])' to $groupId as $role ..." -NoNewline
                Write-Output "---> Adding '$($teamusers[$j])' to $groupId as $role ..." | out-file $global:logFilePath -NoNewline -Append

                try {
                    Add-TeamUser -GroupId $groupId -User $($teamusers[$j]) -Role $role -Verbose

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

function Update-ATeam([PSCustomObject] $row)
{
    $teamGroupId = $($row.TeamGroupId.Trim())

    $channels = $($row.Channels.Trim())
    $owners = $($row.OwnerUPNs.Trim())
    $members = $($row.MemberUPNs.Trim())

    $optionalParams = @{}

    $teamDisplayName = $($row.TeamDisplayName.Trim())
    if(![string]::IsNullOrEmpty($teamDisplayName)){
        $optionalParams.Add("DisplayName", $teamDisplayName)
    }

    $teamDescription = $($row.TeamDescription.Trim())
    if(![string]::IsNullOrEmpty($teamDescription)){
        $optionalParams.Add("Description", $teamDescription)
    }

    $visibility = $($row.TeamPrivacy.Trim())
    if(![string]::IsNullOrWhiteSpace($visibility)) {
        $optionalParams.Add("Visibility", $visibility)
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

    if(![string]::IsNullOrWhiteSpace($teamGroupId)) {
        # update based on group Id

        try {

            Write-Host "You are updating a team with group id '$teamGroupId' or named '$teamDisplayName'..." -NoNewline
            Write-Output "You are updating a team with group id '$teamGroupId' or named '$teamDisplayName'..." | out-file $global:logFilePath -NoNewline -Append

            Set-Team -GroupId $teamGroupId @optionalParams -ErrorAction Stop

            Write-Host "Done" -BackgroundColor Green
            Write-Output "Done" | out-file $global:logFilePath -Append

            if(![string]::IsNullOrEmpty($channels)) {
                Add-Channels -channels $channels -groupId $teamGroupId
            }

            if(![string]::IsNullOrEmpty($owners)){
                Add-Users -users $owners -groupId $teamGroupId -role Owner
            }

            if(![string]::IsNullOrEmpty($members)){
                Add-Users -users $members -groupId $teamGroupId -role Member 
            }
        
            Get-Team -GroupId $teamGroupId | fl | out-file $global:logFilePath -Append
          
            # populate log
            "$teamDisplayName `t $visibility `t $teamGroupId `t $owners `t $members `t $channels `t $(Get-Date) `t $action" | out-file $global:logCSVPath -Append 
        }
        catch {
            Write-Host "Failed" -BackgroundColor Red
            Write-Output "Failed" | out-file $global:logFilePath -Append
        }
    }
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
        Write-Host "No" -BackgroundColor Red
        Write-Output "No" | out-file $global:logFilePath -Append
    }
    else {
        $exists = $true

        Write-Host "...Yes. GroupId: $($team.GroupId)..." -BackgroundColor Green
        Write-Output "...Yes. GroupId: $($team.GroupId)..." | out-file $global:logFilePath -Append

        if([string]::IsNullOrEmpty($teamGroupId)) {
            # team discovered from display name, update group id to pass on
            Write-Host "Updating TeamGroupId of team '$teamDisplayName'..." -NoNewline

            $row.TeamGroupId = $($team.GroupId)

            Write-Host "Done" -BackgroundColor Green
        }
    }

    switch($action)
    {
        'Update'{            
            if($exists) {
                Update-ATeam $row
            }
            else {
                Write-Host "Will NOT be updated" -BackgroundColor Red
                Write-Output "Will NOT be updated" | out-file $global:logFilePath -Append 
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
                    TeamDescription=$teamDescription;
                    MailNickName=$teamMailNickName;
                    TeamPrivacy=$teamPrivacy;
                    OwnerUPNs=$owners;
                    MemberUPNs=$members;
                    Channels=$channels;
                    AllowGiphy=$allowGiphy;
                    GiphyContentRating=$giphyContentRating;
                    AllowStickersAndMemes=$allowStickersAndMemes;
                    AllowCustomMemes=$allowCustomMemes;
                    AllowGuestCreateUpdateChannels=$allowGuestCreateUpdateChannels;
                    AllowGuestDeleteChannels=$allowGuestDeleteChannels;
                    AllowCreateUpdateChannels=$allowCreateUpdateChannels;
                    AllowDeleteChannels=$allowDeleteChannels;
                    AllowAddRemoveApps=$allowAddRemoveApps;
                    AllowCreateUpdateRemoveTabs=$allowCreateUpdateRemoveTabs;
                    AllowCreateUpdateRemoveConnectors=$allowCreateUpdateRemoveConnectors;
                    AllowUserEditMessages=$allowUserEditMessages;
                    AllowUserDeleteMessages=$allowUserDeleteMessages;
                    AllowOwnerDeleteMessages=$allowOwnerDeleteMessages;
                    AllowTeamMentions=$allowTeamMentions;
                    AllowChannelMentions=$allowChannelMentions;
                    ShowInTeamsSearchAndSuggestions=$showInTeamsSearchAndSuggestions
                }
    
            Action-ATeam $row | Out-Null
        }
    }

    Write-Output "Logging ended - $(Get-Date)`n" | out-file $global:logFilePath -Append

    Disconnect-MicrosoftTeams
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow