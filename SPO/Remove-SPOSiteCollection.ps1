# Connecting to your tenant 
#$conn = Connect-PnPOnline -Url https://murphyoil-admin.sharepoint.com -Credentials (Get-Credential) -ReturnConnection

$siteCollUrl = "https://murphyoil.sharepoint.com/sites/MOCDIMAdoption"

$conn = Connect-SPOService -Url https://murphyoil-admin.sharepoint.com -Credential (Get-Credential)

Write-Host "Deleting site collection $siteCollUrl..." -NoNewline

# Remove site collection to recycle bin
Remove-SPOSite -Identity $siteCollUrl -NoWait

Write-Host "Done" -BackgroundColor Green -ForegroundColor White

Write-Host "Purging site collection $siteCollUrl from recycle bin..." -NoNewline

# Purge from recycle bin
Remove-SPODeletedSite -Identity $siteCollUrl -NoWait

Write-Host "Done" -BackgroundColor Green -ForegroundColor White