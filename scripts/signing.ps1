# signing.ps1 — in-house PKI code signing (better-than-Microsoft edition)
#
# Fully self-hosted, cross-platform (Windows/macOS/Linux), zero per-signature
# cost, your own publisher identity. Two-tier CA (offline root + issuing CA),
# serial database, revocation workflow, trust installer, verifier.
#
# NOT a Smart App Control bypass: SAC only trusts Microsoft-rooted CAs. This
# PKI gives you the same guarantees as a commercial OV cert everywhere else
# (AppLocker Trusted Publishers, enterprise WDAC policy, internal distribution,
# code provenance) — with your name as publisher and no CA fees.
#
# Requires: openssl on PATH (all platforms). Windows signing prefers
# osslsigncode (brew/apt/winget) then signtool, then PowerShell fallback.
#
# Usage:
#   pwsh scripts/signing.ps1 new-ca -RootDir .\pki\root -IssuingDir .\pki\issuing -SubjectRoot "CN=CustomsAppify Root CA" -SubjectIssuing "CN=CustomsAppify Code Signing CA"
#   pwsh scripts/signing.ps1 issue -CaDir .\pki\issuing -Subject "CN=CustomsAppify Runner" -OutPfx .\certs\runner.pfx -Password change-me
#   pwsh scripts/signing.ps1 sign -File .\dist\runner.exe -Pfx .\certs\runner.pfx -Password change-me
#   pwsh scripts/signing.ps1 trust-install -CaCrt .\pki\root\ca.crt
#   pwsh scripts/signing.ps1 verify -File .\dist\runner.exe
#   pwsh scripts/signing.ps1 revoke -CaDir .\pki\issuing -Serial 1001

param(
    [Parameter(Position = 0, Mandatory = $true)]
    [ValidateSet("new-ca", "issue", "sign", "verify", "trust-install", "revoke")]
    [string]$Command,

    [string]$RootDir = "./pki/root",
    [string]$IssuingDir = "./pki/issuing",
    [string]$SubjectRoot = "O=Nyixo, CN=Nyixo Root CA",
    [string]$SubjectIssuing = "O=Nyixo, CN=Nyixo Code Signing CA",
    [string]$CaDir = "./pki/issuing",
    [string]$Subject = "O=Nyixo, CN=CustomsAppify Runner",
    [string]$OutPfx = "",
    [string]$Password = "change-me",
    [string]$File = "",
    [string]$Pfx = "",
    [string]$CaCrt = "",
    [string]$Serial = "",
    [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"

function Assert-OpenSsl {
    $o = Get-Command openssl -ErrorAction SilentlyContinue
    if (-not $o) { throw "openssl not found on PATH (winget install openssl / brew install openssl / apt install openssl)" }
    return $o.Source
}

function Invoke-Ossl {
    param([string[]]$Args)
    $ca = Assert-OpenSsl
    & $ca @Args
    if ($LASTEXITCODE -ne 0) { throw "openssl failed: $($Args -join ' ')" }
}

function New-OpenSslConf {
    param([string]$Path, [string]$CaDir)
    $db = Join-Path $CaDir "index.txt"
    $serial = Join-Path $CaDir "serial"
    $crl = Join-Path $CaDir "crl.pem"
    $newcerts = Join-Path $CaDir "newcerts"
    New-Item -ItemType Directory -Force -Path $newcerts | Out-Null
    if (-not (Test-Path $db)) { New-Item -ItemType File -Path $db | Out-Null }
    if (-not (Test-Path $serial)) { "1000" | Set-Content $serial }
    @"
[ ca ]
default_ca = CA_default
[ CA_default ]
dir = $CaDir
database = $db
new_certs_dir = $newcerts
certificate = $CaDir/ca.crt
private_key = $CaDir/ca.key
serial = $serial
crlnumber = $CaDir/crlnumber
crl = $crl
default_md = sha256
policy = policy_loose
default_days = 825
copy_extensions = copy
[ policy_loose ]
countryName = optional
stateOrProvinceName = optional
localityName = optional
organizationName = optional
organizationalUnitName = optional
commonName = supplied
emailAddress = optional
[ v3_leaf ]
basicConstraints = CA:FALSE
keyUsage = digitalSignature
extendedKeyUsage = codeSigning
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid,issuer
"@ | Set-Content $Path
}

function New-Ca {
    New-Item -ItemType Directory -Force -Path $RootDir | Out-Null
    New-Item -ItemType Directory -Force -Path $IssuingDir | Out-Null

    # ---- offline root CA (keep private key cold/offline) ----
    $rootKey = Join-Path $RootDir "ca.key"
    $rootCrt = Join-Path $RootDir "ca.crt"
    if (-not (Test-Path $rootKey)) { Invoke-Ossl genrsa -out $rootKey 4096 }
    Invoke-Ossl req -x509 -new -key $rootKey -sha256 -days 7300 -subj $SubjectRoot -out $rootCrt

    # ---- issuing CA ----
    $issKey = Join-Path $IssuingDir "ca.key"
    $issCsr = Join-Path $IssuingDir "ca.csr"
    $issCrt = Join-Path $IssuingDir "ca.crt"
    $issConf = Join-Path $IssuingDir "openssl.cnf"
    if (-not (Test-Path $issKey)) { Invoke-Ossl genrsa -out $issKey 4096 }
    if (-not (Test-Path $issCsr)) { Invoke-Ossl req -new -key $issKey -subj $SubjectIssuing -out $issCsr }
    @"
basicConstraints = CA:TRUE, pathlen:0
keyUsage = keyCertSign, cRLSign, digitalSignature
subjectKeyIdentifier = hash
authorityKeyIdentifier = keyid,issuer
"@ | Set-Content (Join-Path $IssuingDir "ca_ext.cnf")
    Invoke-Ossl x509 -req -in $issCsr -CA $rootCrt -CAkey $rootKey -CAcreateserial `
        -days 3650 -sha256 -extfile (Join-Path $IssuingDir "ca_ext.cnf") -out $issCrt
    New-OpenSslConf -Path $issConf -CaDir $IssuingDir

    Write-Host "ROOT CA (keep offline):  $rootCrt"
    Write-Host "ISSUING CA:               $issCrt"
    Write-Host "Trust this on machines:   $rootCrt"
}

function New-Issue {
    $conf = Join-Path $CaDir "openssl.cnf"
    if (-not (Test-Path $conf)) { throw "issuing CA not found in $CaDir (run new-ca first)" }
    New-Item -ItemType Directory -Force -Path (Split-Path $OutPfx) | Out-Null
    if (-not $OutPfx) { $OutPfx = Join-Path $CaDir "runner.pfx" }

    $key = Join-Path $CaDir "runner.key"
    $csr = Join-Path $CaDir "runner.csr"
    $crt = Join-Path $CaDir "runner.crt"
    Invoke-Ossl genrsa -out $key 4096
    Invoke-Ossl req -new -key $key -subj $Subject -out $csr
    Invoke-Ossl ca -batch -config $conf -in $csr -out $crt

    $rootCrt = Join-Path (Split-Path $CaDir -Parent) "root/ca.crt"
    if (Test-Path $rootCrt) {
        Invoke-Ossl pkcs12 -export -out $OutPfx -inkey $key -in $crt -certfile $rootCrt -passout "pass:$Password"
    } else {
        Invoke-Ossl pkcs12 -export -out $OutPfx -inkey $key -in $crt -passout "pass:$Password"
    }
    Write-Host "issued: $OutPfx (use with 'sign')"
}

function Invoke-Sign {
    if (-not $File -or -not (Test-Path $File)) { throw "File not found: $File" }
    if (-not $Pfx -or -not (Test-Path $Pfx)) { throw "PFX not found: $Pfx" }

    # Cross-platform primary: osslsigncode (works on all OSes).
    $ossl = Get-Command osslsigncode -ErrorAction SilentlyContinue
    if ($ossl) {
        & $ossl.Source sign -pkcs12 $Pfx -pass $Password -in $File -out "$File.signed" -t $TimestampUrl
        if ($LASTEXITCODE -ne 0) { throw "osslsigncode failed (exit $LASTEXITCODE)" }
        Move-Item -Force "$File.signed" $File
        Write-Host "signed with osslsigncode"
        return
    }

    if ($IsWindows -or $env:OS -eq "Windows_NT") {
        $signtool = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin\*\x64\signtool.exe" -ErrorAction SilentlyContinue | Select-Object -Last 1
        if ($signtool) {
            & $signtool.FullName sign /fd SHA256 /f $Pfx /p $Password /tr $TimestampUrl /td SHA256 $File
            if ($LASTEXITCODE -ne 0) { throw "signtool failed (exit $LASTEXITCODE)" }
            return
        }
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Pfx, $Password)
        $sig = Set-AuthenticodeSignature -FilePath $File -Certificate $cert -HashAlgorithm SHA256 -TimestampServer $TimestampUrl
        if ($sig.Status -ne "Valid") { $sig = Set-AuthenticodeSignature -FilePath $File -Certificate $cert -HashAlgorithm SHA256 }
        Write-Host "signature status: $($sig.Status)"
        return
    }

    if ($IsMacOS) {
        security import $Pfx -k "$env:HOME/Library/Keychains/login.keychain-db" -P $Password | Out-Null
        $identity = "CN=" + ((Invoke-Ossl pkcs12 -in $Pfx -passin "pass:$Password" -nokeys -clcerts 2>$null |
            Select-String "subject=" | Select-Object -First 1) -replace '.*subject= ?','' -replace ',.*$','')
        & codesign --force --sign $identity --options runtime --timestamp $File
        if ($LASTEXITCODE -ne 0) { throw "codesign failed (exit $LASTEXITCODE)" }
        Write-Host "macOS codesign OK ($identity)"
        return
    }

    $sum = Get-FileHash -Algorithm SHA256 $File
    "$($sum.Hash)  $File" | Set-Content "$File.sha256"
    Write-Host "Linux: wrote $File.sha256 (no Authenticode needed)."
}

function Invoke-Verify {
    if (-not $File -or -not (Test-Path $File)) { throw "File not found: $File" }
    $ossl = Get-Command osslsigncode -ErrorAction SilentlyContinue
    if ($ossl) {
        & $ossl.Source verify -in $File
    } elseif ($IsWindows -or $env:OS -eq "Windows_NT") {
        $sig = Get-AuthenticodeSignature $File
        $sig | Format-List Status, StatusMessage, @{n='Signer';e={$_.SignerCertificate.Subject}}
    } else {
        Get-FileHash -Algorithm SHA256 $File
    }
}

function Invoke-TrustInstall {
    if (-not $CaCrt) { $CaCrt = Join-Path $RootDir "ca.crt" }
    if (-not (Test-Path $CaCrt)) { throw "CA cert not found: $CaCrt" }
    if ($IsWindows -or $env:OS -eq "Windows_NT") {
        Import-Certificate -FilePath $CaCrt -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
        Import-Certificate -FilePath $CaCrt -CertStoreLocation Cert:\LocalMachine\TrustedPublisher | Out-Null
        Write-Host "trust installed (LocalMachine Root + TrustedPublisher)"
    } elseif ($IsMacOS) {
        # shell out with sudo; user must enter password
        sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain $CaCrt
        Write-Host "trust installed (System keychain); note: notarization still required for public macOS distribution"
    } else {
        sudo cp $CaCrt /usr/local/share/ca-certificates/customsappify-ca.crt
        sudo update-ca-certificates
        Write-Host "trust installed (Linux CA store)"
    }
}

function Invoke-Revoke {
    $conf = Join-Path $CaDir "openssl.cnf"
    if (-not (Test-Path $conf)) { throw "issuing CA not found in $CaDir" }
    if (-not $Serial) { throw "-Serial required (see $CaDir/index.txt)" }
    Invoke-Ossl ca -config $conf -revoke (Join-Path $CaDir "newcerts\$Serial.pem")
    Invoke-Ossl ca -config $conf -gencrl -out (Join-Path $CaDir "crl.pem")
    Write-Host "revoked serial $Serial; CRL at $CaDir/crl.pem (distribute with updates)"
}

switch ($Command) {
    "new-ca"      { New-Ca }
    "issue"       { New-Issue }
    "sign"        { Invoke-Sign }
    "verify"      { Invoke-Verify }
    "trust-install" { Invoke-TrustInstall }
    "revoke"      { Invoke-Revoke }
}
