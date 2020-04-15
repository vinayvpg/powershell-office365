import-module MSOnline

$msolcred = get-credential
connect-msolservice -credential $msolcred

$clientId = "96c1f36f-e3d9-4b37-9c7f-721cd794f7ba" #m-files connector app principal

$bytes = New-Object Byte[] 32
$rand = [System.Security.Cryptography.RandomNumberGenerator]::Create()
$rand.GetBytes($bytes)
$rand.Dispose()

$newClientSecret = [System.Convert]::ToBase64String($bytes)

$dtStart = [System.DateTime]::Now
$dtEnd = $dtStart.AddYears(3)

New-MsolServicePrincipalCredential -AppPrincipalId $clientId -Type Symmetric -Usage Sign -Value $newClientSecret -StartDate $dtStart –EndDate $dtEnd

New-MsolServicePrincipalCredential -AppPrincipalId $clientId -Type Symmetric -Usage Verify -Value $newClientSecret -StartDate $dtStart –EndDate $dtEnd

New-MsolServicePrincipalCredential -AppPrincipalId $clientId -Type Password -Usage Verify -Value $newClientSecret -StartDate $dtStart –EndDate $dtEnd

$newClientSecret
