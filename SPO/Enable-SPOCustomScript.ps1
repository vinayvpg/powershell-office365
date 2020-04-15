# Connecting to your tenant 
#$conn = Connect-PnPOnline -Url https://murphyoil-admin.sharepoint.com -Credentials (Get-Credential) -ReturnConnection

$adminUrl = "https://murphyoil-admin.sharepoint.com"

$siteCollUrl = "https://murphyoil.sharepoint.com/sites/MOCLegacyDocs"

$conn = Connect-SPOService -Url $adminUrl -Credential (Get-Credential)

# allow custom scripts
Set-SPOSite $siteCollUrl -DenyAddAndCustomizePages 0