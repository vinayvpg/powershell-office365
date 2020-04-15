Connect-SPOService -Url https://murphyoil-admin.sharepoint.com -Credential (Get-Credential)

$ExtUser = Get-SPOExternalUser -Filter "vinay.prabhugaonkar@sparkhound.com" -SiteUrl https://murphyoil.sharepoint.com/sites/OPS_EXT_MurphyPetrobras

Remove-SPOExternalUser -UniqueIDs @($ExtUser.UniqueId)