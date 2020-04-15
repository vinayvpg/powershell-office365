[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to csv report file")]
    [string] $reportCSVPath,

    [Parameter(Mandatory=$false, Position=1, HelpMessage="Tenant admin site url https://tenantname-admin.sharepoint.com")]
    [string] $tenantAdminSiteUrl=“https://murphyoil-admin.sharepoint.com”,

    [Parameter(Mandatory=$false, Position=2, HelpMessage="Current login name of SPUser to migrate")]
    [string] $userAlias,

    [Parameter(Mandatory=$false, Position=3, HelpMessage="New login name of SPUser after migration")]
    [switch] $byAlias = $false
)

$ErrorActionPreference = "Continue"

Cls

function GetMailboxPropertiesInBatch($usersBatch) {
    $usersBatch | ForEach-Object {
        $userName = $_.Name
        $userPrincipalName = $_.UserPrincipalName
        $litigationHoldDate = $_.LitigationHoldDate
        $litigationHoldOwner = $_.LitigationHoldOwner
        $litigationHoldDuration = $_.LitigationHoldDuration

        Write-Host "`n---> Processing UserName: $userName..." -NoNewline

        "$userName `t $userPrincipalName `t $litigationHoldDate `t $litigationHoldOwner `t $litigationHoldDuration" | out-file $global:reportPath -Append 

        Write-Host "Done" -BackgroundColor Green
    }
}

#---------------------------- main script------------------------------------------

Write-host "Start - $(Get-Date)`n" -ForegroundColor Yellow

$timestamp = Get-Date -Format s | % { $_ -replace ":", "-" }

if([string]::IsNullOrWhiteSpace($reportCSVPath))
{
    $currentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
    $reportCSVPath = $currentDir + "\GetMailbox-Report-" + $timestamp + ".csv"
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
    $obj = Get-Mailbox -ResultSize unlimited -Filter { (LitigationHoldEnabled -eq $true) }
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
Write-Host “Found $($users.Count) users to process” -ForegroundColor Cyan

# Remove the connection to exchange online
Remove-PSSession $exchangeSession

# Set counts
$count = 0
$total = $users.Count
$batchSize = 500

if($total -gt 0) {
    #Write CSV - TAB Separated File Header
    "Name `t UserPrincipalName `t LitigationHoldDate `t LitigationHoldOwner `t LitigationHoldDuration" | out-file $global:reportPath

    $loopNum = [Math]::Floor($total/$batchSize)

    $startAt = 0
    $usersBatch = New-Object System.Collections.ArrayList

    for($i = 1; $i -le $loopNum; $i++) {
        Write-Host "Starting At: $startAt"
        for($j = $startAt; $j -lt ($batchSize + $startAt); $j++) {
            $usersBatch.Add($users.Item($j)) > $null
        }

        GetMailboxPropertiesInBatch -usersBatch $usersBatch

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

    GetMailboxPropertiesInBatch -usersBatch $usersBatch
}

Write-host "`nDone - $(Get-Date)" -ForegroundColor Yellow