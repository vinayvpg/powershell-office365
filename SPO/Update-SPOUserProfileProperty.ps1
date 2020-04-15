[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv report file")]
    [string] $reportCSVPath,

    [Parameter(Mandatory=$false, Position=1, HelpMessage="Tenant admin site url https://tenantname-admin.sharepoint.com")]
    [string] $tenantAdminSiteUrl=“https://murphyoil-admin.sharepoint.com”,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Current login name of SPUser to migrate")]
    [string] $userAlias = "Mohktar_Samon",

    [Parameter(Mandatory=$false, Position=3, HelpMessage="New login name of SPUser after migration")]
    [switch] $byAlias = $false
)
# Global Variables
$sCSOMPath = “C:\Users\prabhvx\OneDrive - Murphy Oil\SPLibs\” # Path to CSOM DLLs
$UserProfilePrefix = “i:0#.f|membership|” # Claims membership prefix

# Adding the Client OM Assemblies
$sCSOMRuntimePath=$sCSOMPath + “Microsoft.SharePoint.Client.Runtime.dll”
$sCSOMUserProfilesPath=$sCSOMPath + “Microsoft.SharePoint.Client.UserProfiles.dll”
$sCSOMPath=$sCSOMPath + “Microsoft.SharePoint.Client.dll”

Add-Type -Path $sCSOMPath
Add-Type -Path $sCSOMRuntimePath
Add-Type -Path $sCSOMUserProfilesPath

$ErrorActionPreference = "Continue"

# Clear the screen
Cls

function UpdateUserProfilePropertiesInBatch($usersBatch, $peopleManager, $spoCtx, $batchNum) {
    $usersBatch | ForEach-Object {
        <#
        # Output the current progress
        $count = $count + 1
        $percent = ($count * 100 / $total) -as [int]
        Write-Progress -Activity “Processing employees” -Status “Processing $($_.DisplayName)…” -PercentComplete $percent
        #>

        # Get the UserName, SF Location (CustomAttribute1)
        $userName = $_.UserPrincipalName
        $extensionAttribute1 = $_.CustomAttribute1
        $extensionAttribute2 = $_.CustomAttribute2
        $office = $_.Office

        Write-Host "`n---> Processing UserName: $userName, Attribute1: $extensionAttribute1, Attribute2: $extensionAttribute2, Office: $office ..." -NoNewline
    
        if (![string]::IsNullOrWhiteSpace($extensionAttribute1)) {
            try {
                # Update the property
                $peopleManager.SetSingleValueProfileProperty($UserProfilePrefix + $userName, “SPS-Location”, $extensionAttribute1)
                #$peopleManager.GetUserProfilePropertyFor($UserProfilePrefix + $userName, "FirstName")

                Write-Host "Done" -BackgroundColor Green

                "$userName `t $extensionAttribute1 `t $extensionAttribute2 `t $office" | out-file $global:reportPath -Append 

            }
            catch {
                Write-Host “Could not set property for $userName" -ForegroundColor Red
            }
        }
    }

    try
    {                     
        Write-Host "Committing changes for batch $batchNum..." -NoNewline -BackgroundColor Cyan

        $spoCtx.ExecuteQuery()
        
        Write-Host "Done" -BackgroundColor Green              
    }
    catch [Microsoft.SharePoint.Client.ServerException]
    {
        $exceptionMsg = $_.Exception.Message
        Write-Host $exceptionMsg -BackgroundColor Red
    }
}

#---------------------------- main script------------------------------------------

Write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

$timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

if([string]::IsNullOrWhiteSpace($reportCSVPath))
{
    $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $reportCSVPath = $currentDir + "\SetSPOUPProperty-Report-" + $timestamp + ".csv"
    Write-Host "You did not specify a path to the report csv file. The report will be created at '$reportCSVPath'" -ForegroundColor Cyan
}

$global:reportPath = $reportCSVPath

# Read the user’s O365 credentials
$credential = Get-Credential -Message "Please enter your organizational tenant admin credential for Office 365"
$adminUserName = $credential.UserName
$adminPassword = $credential.Password #| ConvertTo-SecureString -AsPlainText -Force # Tenant Administrator password

#$credential = New-Object System.Management.Automation.PsCredential($sUserName,$sPassword)

# Connect to Exchange Online
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri “https://outlook.office365.com/powershell-liveid/” -Credential $credential -Authentication “Basic” –AllowRedirection
Import-PSSession $exchangeSession -AllowClobber

$users = New-Object System.Collections.ArrayList

if($byAlias -eq $true) {
    $obj = Get-Mailbox -ResultSize unlimited | ? { $_.Alias -eq $userAlias }
}
else {
    $obj = Get-Mailbox -ResultSize unlimited -Filter { (CustomAttribute1 -ne $null) -and (CustomAttribute2 -eq "Active") }
}

if($obj -ne $null) {
    if($obj.GetType().BaseType.Name -eq "Object") {
        $users.Add($obj)
    }

    if($obj.GetType().BaseType.Name -eq "Array") {
        $users.AddRange($obj)
    }
}
# Output the number of users
Write-Host “Found $($users.Count) employees to process” -ForegroundColor Cyan

# Remove the connection to exchange online
Remove-PSSession $exchangeSession

# Set counts
$count = 0
$total = $users.Count
$batchSize = 500

if($total -gt 0) {
    #Write CSV - TAB Separated File Header
    "UserPrincipalName `t SPS-Location `t UserStatus `t Office" | out-file $global:reportPath

    # SPO Client Object Model Context
    $spoCtx = New-Object Microsoft.SharePoint.Client.ClientContext($tenantAdminSiteUrl)
    $spoCtx.RequestTimeout = 900000 # 15 min (seems to have no effect)
    $spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($adminUserName, $adminPassword)
    $spoCtx.Credentials = $spoCredentials

    $peopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($spoCtx)

    $loopNum = [Math]::Floor($total/$batchSize)

    $startAt = 0
    $usersBatch = New-Object System.Collections.ArrayList

    for($i = 1; $i -le $loopNum; $i++) {
        Write-Host "Starting At: $startAt"
        for($j = $startAt; $j -lt ($batchSize + $startAt); $j++) {
            $usersBatch.Add($users.Item($j)) > $null
        }

        UpdateUserProfilePropertiesInBatch -usersBatch $usersBatch -peopleManager $peopleManager -spoCtx $spoCtx -batchNum $i

        $usersBatch.Clear()

        if($i -lt $loopNum) {
            $startAt += $batchSize
        }
        else {
            $startAt = $batchSize * $loopNum
        }
    }

    # remaining items
    Write-Host "Start at: $startAt"
    for($k = $startAt; $k -lt $total; $k++) {
        $usersBatch.Add($users.Item($k)) > $null
    }

    UpdateUserProfilePropertiesInBatch -usersBatch $usersBatch -peopleManager $peopleManager -spoCtx $spoCtx -batchNum $($loopNum + 1)

    # Dispose of the SharePoint Online context
    $spoCtx.Dispose()
}

Write-host "`nDone - $(Get-Date)" -ForegroundColor Yellow