# sign_runner.ps1 — sign an existing CustomsAppify runner exe.
#
# Usage (Windows PowerShell, Admin recommended for -Trust):
#   powershell -ExecutionPolicy Bypass -File scripts\sign_runner.ps1 `
#     -ExePath C:\CustomsAppify\customsappify-runner.exe -Trust
#
# -Trust: also installs the (self-signed) cert into CurrentUser Root and
#         TrustedPublisher so Authenticode validates and AppLocker
#         trusted-publisher rules pass on this machine.

param(
  [string]$ExePath = "C:\CustomsAppify\customsappify-runner.exe",
  [string]$Thumbprint = "",
  [switch]$Trust
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $ExePath)) {
  throw "exe not found: $ExePath"
}

# Reuse an existing code-signing cert or create a self-signed one.
if (-not $Thumbprint) {
  $cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.EnhancedKeyUsageList -match 'Code Signing' } |
    Select-Object -First 1
  if (-not $cert) {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert `
      -Subject "CN=CustomsAppify Runner" `
      -CertStoreLocation Cert:\CurrentUser\My `
      -NotAfter (Get-Date).AddYears(3)
    Write-Host "created self-signed code-signing cert: $($cert.Thumbprint)"
  }
  $Thumbprint = $cert.Thumbprint
}

if ($Trust) {
  $tmp = Join-Path $env:TEMP "customsappify-runner-cert.cer"
  Export-Certificate -Cert (Get-Item "Cert:\CurrentUser\My\$Thumbprint") -FilePath $tmp | Out-Null
  Import-Certificate -FilePath $tmp -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
  Import-Certificate -FilePath $tmp -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
  Write-Host "trust installed (Root + TrustedPublisher, CurrentUser)"
}

try {
  $sig = Set-AuthenticodeSignature -FilePath $ExePath `
    -Certificate (Get-Item "Cert:\CurrentUser\My\$Thumbprint") `
    -HashAlgorithm SHA256 `
    -TimestampServer http://timestamp.digicert.com
} catch {
  $sig = Set-AuthenticodeSignature -FilePath $ExePath `
    -Certificate (Get-Item "Cert:\CurrentUser\My\$Thumbprint") `
    -HashAlgorithm SHA256
  Write-Host "signed without timestamp (timestamp server unreachable)"
}

Write-Host "signature status: $($sig.Status)"
Get-AuthenticodeSignature $ExePath | Format-List Status, StatusMessage, @{n='Signer';e={$_.SignerCertificate.Subject}}

if ($sig.Status -ne 'Valid') {
  Write-Host ""
  Write-Host "NOTE: not 'Valid' = cert not trusted on this machine yet. Run with"
  Write-Host "  -Trust to install it, or use a CA-issued cert (for Smart App Control"
  Write-Host "  / enterprise WDAC, submit the signed build to Microsoft WDSI:"
  Write-Host "  https://www.microsoft.com/wdsi/filesubmission)"
}
