[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0, HelpMessage="Full path to certificate")]
    [string] $certPath="C:\Users\prabhvx\OneDrive - Murphy Oil\Desktop\Azure AD App Registrations\postman_cert_base64Encoded.cer"
)

$ErrorActionPreference = "Continue"

$cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cer.Import($certPath)
$bin = $cer.GetRawCertData()
$base64Value = [System.Convert]::ToBase64String($bin)

Write-Host "base64Value... $base64Value"

$bin = $cer.GetCertHash()
$base64Thumbprint = [System.Convert]::ToBase64String($bin)
$keyid = [System.Guid]::NewGuid().ToString()


Write-Host "base64Thumbprint... $base64Thumbprint"