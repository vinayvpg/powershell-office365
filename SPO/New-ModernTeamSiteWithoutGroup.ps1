# Connecting to your tenant 
$conn = Connect-PnPOnline -Url https://murphyoil-admin.sharepoint.com -Credentials (Get-Credential) -ReturnConnection

# Creating the Modern Site
New-PnPTenantSite -Title "Land Management - QLS" -Url "https://murphyoil.sharepoint.com/sites/MOCDocs_LAN_QLS" -Description "QLS" -Owner "vinay_prabhugaonkar@contractor.murphyoilcorp.com" -Lcid 1033 -Template "STS#3" -TimeZone 11 -Connection $conn -Wait