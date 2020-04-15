function Export-TeamsList
{   
     param (   
           $ExportPath
           )   
    process{
        Connect-PnPOnline -Scopes "Group.Read.All","User.ReadBasic.All"
        $accesstoken =Get-PnPAccessToken
        $group = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri  "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c+eq+`'Unified`')" -Method Get
        $TeamsList = @()
        do
        {
        foreach($value in $group.value)
        {
            "Group Name: " + $value.displayName + " Group Type: " + $value.groupTypes
            if($value.groupTypes -eq "Unified")
            {
                 $id= $value.id
                 Try
                 {
                 $team = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri  https://graph.microsoft.com/beta/Groups/$id/channels -Method Get
                 #"Channel count for " + $value.displayName + " is " + $team.value.id.count
                 }
                 Catch
                 {
                 #"Could not get channels for " + $value.displayName + ". " + $_.Exception.Message
                 $team = $null
                 }
                 If($team.value.id.count -ge 1)
                 {
                     $Owner = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri  https://graph.microsoft.com/v1.0/Groups/$id/owners -Method Get
                     $Members = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri  https://graph.microsoft.com/v1.0/Groups/$id/Members -Method Get
                     $Teams = "" | Select "TeamsName","TeamType","ChannelCount","Channels","Owners","MembersCount","Members"
                     $Teams.TeamsName = $value.displayname
                     $Teams.TeamType = $value.visibility
                     $Teams.ChannelCount = $team.value.id.count
                     $Teams.Channels = $team.value.displayName -join ";"
                     $Teams.Owners = $Owner.value.userPrincipalName -join ";"
                     $Teams.MembersCount = $Members.value.userPrincipalName.count
                     $Teams.Members = $Members.value.userPrincipalName -join ";"
                     $TeamsList+= $Teams
                     $Teams=$null
                 }
             }
        }
        if ($group.'@odata.nextLink' -eq $null )
        {
        break
        }
        else
        {
        $group = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri $group.'@odata.nextLink' -Method Get
        }
        }while($true);
        $TeamsList |Export-csv $ExportPath -NoTypeInformation
    }
}
Export-TeamsList -ExportPath "C:\Users\prabhvx\OneDrive - Murphy Oil\Desktop\mocconnect\reports\teamsmembers.csv"
