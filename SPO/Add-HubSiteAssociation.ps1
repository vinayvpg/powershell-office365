$conn = Connect-PnPOnline -Url https://murphyoil.sharepoint.com/sites/MOCReportingCoE_Root -Credentials (Get-Credential) -ReturnConnection

$hubSiteUrl = "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_Root"

Register-PnPHubSite -Site $hubSiteUrl -Connection $conn

$childrenOfHub = @(
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_CrossFunctional",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_DCI",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_EXEC",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_FIN",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_GPRO",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_HR",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_HSE",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_IT",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_PM",
                "https://murphyoil.sharepoint.com/sites/MOCReportingCoE_PRA"
            )

foreach($child in $childrenOfHub) {
    Write-Host "Associating $child to $hubSiteUrl..." -NoNewline
    Add-PnPHubSiteAssociation -Site $child -HubSite $hubSiteUrl -Connection $conn
    Write-Host "Done" -BackgroundColor Green -ForegroundColor White
}