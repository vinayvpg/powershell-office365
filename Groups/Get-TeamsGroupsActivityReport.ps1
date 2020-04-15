# A script to check the activity of Office 365 Groups and Teams and report the groups and teams that might be deleted because they're not used.
# We check the group mailbox to see what the last time a conversation item was added to the Inbox folder. 
# Another check sees whether a low number of items exist in the mailbox, which would show that it's not being used.
# We also check the group document library in SharePoint Online to see whether it exists or has been used in the last 90 days.
# And we check Teams compliance items to figure out if any chatting is happening.

$spoTenantAdminUrl = "https://murphyoil-admin.sharepoint.com"

function CheckAndLoadRequiredModules($cred) {
    $res = [PSCustomObject]@{TeamsModuleSuccess = $false;ExchangeModuleSuccess = $false;SPOModuleSuccess = $false}

    $moduleCheckSuccess = $false

    try {
        Write-Host "Loading Exchange Online Powershell module and connecting..." -ForegroundColor Cyan -NoNewline
        
        # Connect to Exchange Online
        $global:exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri “https://outlook.office365.com/powershell-liveid/” -Credential $cred -Authentication “Basic” –AllowRedirection

        Import-PSSession $global:exchangeSession -AllowClobber

        Write-Host "Done" -BackgroundColor Green

        $res.ExchangeModuleSuccess = $true
    }
    catch {
        Write-Host "Could not load exchange online powershell module. Verify credentials." -BackgroundColor Red
    }

    try {
        Write-Host "Loading SharePoint Online Powershell module and connecting..." -ForegroundColor Cyan -NoNewline

        Connect-SPOService -Url $spoTenantAdminUrl -Credential $cred

        Write-Host "Done" -BackgroundColor Green

        $res.SPOModuleSuccess = $true
    }
    catch {

        Write-Host "SharePoint Online Powershell module is not installed. This module is required for SharePoint activity check. Install from https://www.microsoft.com/en-us/download/details.aspx?id=35588" -BackgroundColor Red

        $res.SPOModuleSuccess = $false
    }

    
    if(!(Get-InstalledModule -Name MicrosoftTeams)) {
        Write-Host "Installing MicrosoftTeams module from https://www.powershellgallery.com/packages/MicrosoftTeams/1.0.1..." -NoNewline

        Install-Module MicrosoftTeams -Force

        Write-Host "Done" -BackgroundColor Green
    }

    try {
        Write-Host "Loading MicrosoftTeams module and connecting..." -ForegroundColor Cyan -NoNewline
        
        Import-Module MicrosoftTeams -Force

        Connect-MicrosoftTeams -Credential $cred

        Write-Host "Done" -BackgroundColor Green

        $res.TeamsModuleSuccess = $true
    }
    catch{
        Write-Host "Failed" -BackgroundColor Red

        $res.TeamsModuleSuccess = $false
    }
    
    $moduleCheckSuccess = $res.TeamsModuleSuccess -and $res.SPOModuleSuccess -and $res.ExchangeModuleSuccess

    return $moduleCheckSuccess
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

    if($modulesCheck) {
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

        $global:logCSVPath = $global:logFolderPath + "GroupActivityReport.csv"
        $global:logFilePath = $global:logFolderPath + "GroupActivityReport.html"

        # Check that we are connected to Exchange Online
        $OrgName = (Get-OrganizationConfig).Name
       
        # OK, we seem to be fully connected to both Exchange Online and SharePoint Online...
        Write-Host "Checking for Obsolete Office 365 Groups in the tenant:" $OrgName

        # Setup some stuff we use
        $WarningDate = (Get-Date).AddDays(-90)
        $WarningEmailDate = (Get-Date).AddDays(-365)
        $Today = (Get-Date)
        $Date = $Today.ToShortDateString()
        $TeamsGroups = 0
        $TeamsEnabled = $False
        $ObsoleteSPOGroups = 0
        $ObsoleteEmailGroups = 0
        $Report = @()

        $htmlhead="<html>
	           <style>
	           BODY{font-family: Arial; font-size: 8pt;}
	           H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	           H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	           H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	           TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	           TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	           TD{border: 1px solid #969595; padding: 5px; }
	           td.pass{background: #B7EB83;}
	           td.warn{background: #FFF275;}
	           td.fail{background: #FF2626; color: #ffffff;}
	           td.info{background: #85D4FF;}
	           </style>
	           <body>
                   <div align=center>
                   <p><h1>Office 365 Groups and Teams Activity Report</h1></p>
                   <p><h3>Generated: " + $date + "</h3></p></div>"
		
        # Get a list of all Office 365 Groups in the tenant
        Write-Host "Extracting list of Office 365 Groups for checking..."
        $Groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName

        # Get a list of all teams in the tenant
        $TeamsList = @{}
        Get-Team | ForEach { $TeamsList.Add($_.GroupId, $_.DisplayName) }

        Write-Host "Processing" $Groups.Count "groups"

        # Progress bar
        $ProgDelta = 100/($Groups.count)
        $CheckCount = 0
        $GroupNumber = 0

        ForEach ($G in $Groups) {
            Write-Host "---------------------------------------------------------------------------------------------------------------------------"
            Write-Host "Processing group '$($G.DisplayName)'...." -ForegroundColor Magenta
           
           $GroupNumber++
           $GroupStatus = $G.DisplayName + " ["+ $GroupNumber +"/" + $Groups.Count + "]"

           Write-Progress -Activity "Checking group" -Status $GroupStatus -PercentComplete $CheckCount

           $CheckCount += $ProgDelta
           $ObsoleteReportLine = $G.DisplayName
           $SPOStatus = "Normal"
           $SPOActivity = "Document library in use"
           $NumberWarnings = 0
           $NumberofChats = 0
           $TeamChatData = $Null
           $TeamsEnabled = $False
           $LastItemAddedtoTeams = "No chats"
           $MailboxStatus = $Null

            # Check who manages the group
          $ManagedBy = $G.ManagedBy
          If ([string]::IsNullOrWhiteSpace($ManagedBy) -and [string]::IsNullOrEmpty($ManagedBy)) {
             $ManagedBy = "No owners"
             Write-Host $G.DisplayName "has no group owners!" -ForegroundColor Red
            }
          Else {
             $ManagedBy = (Get-Mailbox -Identity $G.ManagedBy[0]).DisplayName
            }
  
            # Fetch information about activity in the Inbox folder of the group mailbox  
           $Data = (Get-MailboxFolderStatistics -Identity $G.Alias -IncludeOldestAndNewestITems -FolderScope Inbox)
           $LastConversation = $Data.NewestItemReceivedDate
           $NumberConversations = $Data.ItemsInFolder
           $MailboxStatus = "Normal"
  
           If ($Data.NewestItemReceivedDate -le $WarningEmailDate) {
              Write-Host "Last conversation item created in" $G.DisplayName "was" $Data.NewestItemReceivedDate "----> Obsolete?"

              $ObsoleteReportLine = $ObsoleteReportLine + " Last conversation dated: " + $Data.NewestItemReceivedDate + "."
              $MailboxStatus = "Group Inbox Not Recently Used"
              $ObsoleteEmailGroups++
              $NumberWarnings++ 
            }
           Else {
            # Some conversations exist - but if there are fewer than 20, we should flag this...
              If ($Data.ItemsInFolder -lt 20) {
                   $ObsoleteReportLine = $ObsoleteReportLine + " Only " + $Data.ItemsInFolder + " conversation item(s) found."
                   $MailboxStatus = "Low number of conversations"
                   $NumberWarnings++
                }
            }

            # Loop to check SharePoint document library
           If ($G.SharePointDocumentsUrl -ne $Null) {
              $SPOSite = (Get-SPOSite -Identity $G.SharePointDocumentsUrl.replace("/Shared Documents", ""))
              $AuditCheck = $G.SharePointDocumentsUrl + "/*"
              $AuditRecs = 0
              $AuditRecs = (Search-UnifiedAuditLog -RecordType SharePointFileOperation -StartDate $WarningDate -EndDate $Today -ObjectId $AuditCheck -SessionCommand ReturnNextPreviewPage)
              If ($AuditRecs -eq $null) {
                 #Write-Host "No audit records found for" $SPOSite.Title "-> Potentially obsolete!"
                 $ObsoleteSPOGroups++   
                 $ObsoleteReportLine = $ObsoleteReportLine + " No SPO activity detected in the last 90 days."  
                }          
            }
            Else {
                # The SharePoint document library URL is blank, so the document library was never created for this group
                 #Write-Host "SharePoint has never been used for the group" $G.DisplayName 
                $ObsoleteSPOGroups++  
                $ObsoleteReportLine = $ObsoleteReportLine + " SPO document library never created." 
            }

            # Report to the screen what we found - but only if something was found...   
            If ($ObsoleteReportLine -ne $G.DisplayName) {
                Write-Host $ObsoleteReportLine 
            }

            # Generate the number of warnings to decide how obsolete the group might be...   
            If ($AuditRecs -eq $Null) {
               $SPOActivity = "No SPO activity detected in the last 90 days"
               $NumberWarnings++ 
            }

            If ($G.SharePointDocumentsUrl -eq $Null) {
               $SPOStatus = "Document library never created"
               $NumberWarnings++ 
            }
  
            $Status = "Pass"
            If ($NumberWarnings -eq 1)
               {
               $Status = "Warning"
            }
            If ($NumberWarnings -gt 1)
               {
               $Status = "Fail"
            } 

            # If Team-Enabled, we can find the date of the last chat compliance record
            If ($TeamsList.ContainsKey($G.ExternalDirectoryObjectId) -eq $True) {
                  $TeamsEnabled = $True
                  $TeamChatData = (Get-MailboxFolderStatistics -Identity $G.Alias -IncludeOldestAndNewestItems -FolderScope ConversationHistory)
                  If ($TeamChatData.ItemsInFolder[1] -ne 0) {
                      $LastItemAddedtoTeams = $TeamChatData.NewestItemReceivedDate[1]
                      $NumberofChats = $TeamChatData.ItemsInFolder[1] 
                      If ($TeamChatData.NewestItemReceivedDate -le $WarningEmailDate) {
                        Write-Host "Team-enabled group" $G.DisplayName "has only" $TeamChatData.ItemsInFolder[1] "compliance record(s)" }
                    }
            }

            # Generate a line for this group for our report
            $ReportLine = [PSCustomObject][Ordered]@{
                  GroupName           = $G.DisplayName
                  ManagedBy           = $ManagedBy
                  Members             = $G.GroupMemberCount
                  ExternalGuests      = $G.GroupExternalMemberCount
                  Description         = $G.Notes
                  MailboxStatus       = $MailboxStatus
                  TeamEnabled         = $TeamsEnabled
                  LastChat            = $LastItemAddedtoTeams
                  NumberChats         = $NumberofChats
                  LastConversation    = $LastConversation
                  NumberConversations = $NumberConversations
                  SPOActivity         = $SPOActivity
                  SPOStatus           = $SPOStatus
                  NumberWarnings      = $NumberWarnings
                  Status              = $Status}
        
           $Report += $ReportLine
        }

        # Create the HTML report
        $PercentTeams = ($TeamsList.Count/$Groups.Count)
        $htmlbody = $Report | ConvertTo-Html -Fragment
        $htmltail = "<p>Report created for: " + $OrgName + "
                     </p>
                     <p>Number of groups scanned: " + $Groups.Count + "</p>" +
                     "<p>Number of potentially obsolete groups (based on document library activity): " + $ObsoleteSPOGroups + "</p>" +
                     "<p>Number of potentially obsolete groups (based on conversation activity): " + $ObsoleteEmailGroups + "<p>"+
                     "<p>Number of Teams-enabled groups    : " + $TeamsList.Count + "</p>" +
                     "<p>Percentage of Teams-enabled groups: " + ($PercentTeams).tostring("P") + "</body></html>"	

        $htmlreport = $htmlhead + $htmlbody + $htmltail

        $htmlreport | Out-File $global:logFilePath  -Encoding UTF8

        # Summary note 
        Write-Host $ObsoleteSPOGroups "obsolete group document libraries and" $ObsoleteEmailGroups "obsolete email groups found out of" $Groups.Count "checked"
        Write-Host "Summary report available in" $global:logFilePath "and CSV file saved in" $global:logCSVPath

        # Create the csv report
        $Report | Export-CSV -NoTypeInformation $global:logCSVPath
    }

    if($global:exchangeSession -ne $null){
        Remove-PSSession -Session $global:exchangeSession
    }
}

Write-Host "`nDone - $(Get-Date)" -ForegroundColor Yellow