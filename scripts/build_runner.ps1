# Build the encrypted (compiled) runner binary with Nuitka on Windows.
#
# Usage (in PowerShell on the Windows PC):
#   cd <repo>
#   powershell -ExecutionPolicy Bypass -File scripts\build_runner.ps1 -Version 1.0.9
#
# Requirements: Python 3.11 on PATH (python --version). The script makes its own
# venv and installs requirements.txt + nuitka.

param(
  [Parameter(Mandatory = $true)]
  [string]$Version,
  [switch]$Sign,
  [string]$SignThumbprint = "",
  [switch]$TrustSelfSigned
)

$ErrorActionPreference = "Stop"

$Platform = "windows-x86_64"
# Works both in the repo (script lives in scripts\) and in the staged
# MacShare folder (script lives next to src\).
$Root = $PSScriptRoot
if (-not (Test-Path (Join-Path $Root "src\web\runner")) -and (Test-Path (Join-Path $Root "..\web\runner"))) {
  $Root = [string](Resolve-Path (Join-Path $Root ".."))
}
$Work     = Join-Path $Root ".runner-build"
$Publish  = Join-Path $Root "web\static\runner"
$Name     = "customsappify-runner-$Version-$Platform.exe"
$Venv     = Join-Path $Work "venv"

Write-Host "==> staging source at $Work"
if (Test-Path "$Work\src") { Remove-Item -Recurse -Force "$Work\src" }
New-Item -ItemType Directory -Force -Path "$Work\src\web" | Out-Null
# Source may be laid out as <root>\web\... (repo) or pre-staged as <root>\src\web\...
if (Test-Path "$Root\src\web\runner") {
  $SrcRoot = Join-Path $Root "src"
} else {
  $SrcRoot = $Root
}
Copy-Item -Recurse -Force "$SrcRoot\web\runner"  "$Work\src\web\"
Copy-Item -Recurse -Force "$SrcRoot\web\modules" "$Work\src\web\"
Copy-Item -Force "$SrcRoot\web\__init__.py"      "$Work\src\web\"
Copy-Item -Force "$SrcRoot\requirements.txt"     "$Work\src\"

Write-Host "==> stamping RUNNER_VERSION=$Version"
$client = Join-Path $Work "src\web\runner\client.py"
$text = [IO.File]::ReadAllText($client)
$text = [regex]::Replace($text, '^RUNNER_VERSION = ".*?"', "RUNNER_VERSION = `"$Version`"", [Text.RegularExpressions.RegexOptions]::Multiline)
[IO.File]::WriteAllText($client, $text)

# Find a working Python interpreter - use whatever is installed on this PC.
# The Microsoft Store stub in WindowsApps is fine IF the Store Python app is
# installed; if not, it just opens the Store page. So probe each candidate by
# actually running it instead of trusting the filename.
$Py = $null
$candidates = @()
foreach ($cand in @("python", "python3", "py")) {
  $cmd = Get-Command $cand -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source -and (Test-Path $cmd.Source)) {
    $candidates += $cmd.Source
  }
}
$candidates += @(
  "C:\Python311\python.exe", "C:\Program Files\Python311\python.exe",
  "C:\Program Files (x86)\Python311\python.exe",
  "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
  "C:\Python312\python.exe", "C:\Program Files\Python312\python.exe",
  "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
  "C:\Python310\python.exe", "C:\Program Files\Python310\python.exe",
  "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe"
)
# Probe every candidate once, remember the ones that actually run, keyed by
# their real version (install dirs are named Python311 - no dot - so we cannot
# filter on the path string). Skip Python 3.13+ - pandas and other wheels are
# not published for them yet, so pip falls back to a source build that needs
# MSVC/meson and fails.
$working = @{}   # version string -> path
foreach ($c in $candidates) {
  try {
    $probe = & $c -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
    if ($LASTEXITCODE -eq 0 -and $probe) {
      $working["$($probe.Trim())"] = $c
    }
  } catch {
    # Non-runnable interpreter (e.g. Store stub without the app) - skip it.
    continue
  }
}
foreach ($ver in @("3.11", "3.12", "3.10")) {
  if ($working.ContainsKey($ver)) {
    $Py = $working[$ver]
    Write-Host "==> found working Python: $Py ($ver)"
    break
  }
}
if (-not $Py) {
  throw "No usable Python 3.10-3.12 interpreter found on this PC (found only Python 3.13+, which has no pandas wheels, or non-runnable Store stubs). Install Python 3.11 from https://www.python.org/downloads/ (tick 'Add Python to PATH') and re-run this script."
}

# Recreate the venv if it is missing or was left over from an earlier run that
# picked a wrong Python version (e.g. 3.13+, which has no pandas wheels).
$NeedVenv = $true
if (Test-Path "$Venv\Scripts\python.exe") {
  $venvVer = & "$Venv\Scripts\python.exe" -c "import sys; print(sys.version_info[:2])" 2>$null
  if ($LASTEXITCODE -eq 0 -and $venvVer -and "$venvVer".Trim() -in @("(3, 11)","(3, 12)","(3, 10)")) {
    $NeedVenv = $false
  } else {
    Write-Host "==> discarding old venv (was Python $venvVer)"
    Remove-Item -Recurse -Force "$Venv"
  }
}
if ($NeedVenv) {
  Write-Host "==> creating build venv at $Venv"
  & $Py -m venv "$Venv"
  if ($LASTEXITCODE -ne 0) { throw "venv creation failed (exit $LASTEXITCODE)" }
  & "$Venv\Scripts\python.exe" -m pip install -q -U pip
  & "$Venv\Scripts\python.exe" -m pip install -q -r "$Work\src\requirements.txt"
  if ($LASTEXITCODE -ne 0) { throw "requirements install failed (exit $LASTEXITCODE)" }
} else {
  # Top up deps added since the venv was created (fast when already present).
  & "$Venv\Scripts\python.exe" -m pip install -q -r "$Work\src\requirements.txt"
}

# Nuitka is required; ordered-set/zstandard only speed up the onefile build and
# are not installable on machines without MSVC, so they are best-effort.
& "$Venv\Scripts\python.exe" -m pip install -q -U nuitka
if ($LASTEXITCODE -ne 0) { throw "nuitka install failed (exit $LASTEXITCODE)" }
try {
  & "$Venv\Scripts\python.exe" -m pip install -q -U ordered-set zstandard
} catch {
  Write-Host "==> ordered-set/zstandard unavailable; continuing without them"
}

Write-Host "==> compiling with Nuitka (this takes several minutes)"
Push-Location "$Work\src"
& "$Venv\Scripts\python.exe" -m nuitka `
  --onefile --standalone --assume-yes-for-downloads --jobs=1 --low-memory `
  --include-package=websockets --include-package=websockets.asyncio `
  --include-package=pandas --include-package=dateutil `
  --include-package=selenium --include-package=webdriver_manager `
  --include-package=PIL --include-package=pikepdf `
  --include-package=undetected_chromedriver --include-package=numpy `
  --nofollow-import-to=pymupdf --nofollow-import-to=pymupdf.mupdf `
  --nofollow-import-to=web.modules --nofollow-import-to=web.modules.* `
  --no-deployment-flag=excluded-module-usage `
  --output-dir="$Work\dist" --output-filename="customsappify-runner.exe" `
  web\runner\client.py
$rc = $LASTEXITCODE
Pop-Location
if ($rc -ne 0) { throw "Nuitka build failed (exit $rc)" }

if (-not (Test-Path $Publish)) { New-Item -ItemType Directory -Force -Path $Publish | Out-Null }
Move-Item -Force "$Work\dist\customsappify-runner.exe" (Join-Path $Publish $Name)
Write-Host "==> published $(Join-Path $Publish $Name)"

# ------------------------------------------------------- optional signing
if ($Sign) {
  $exe = Join-Path $Publish $Name
  if (-not $SignThumbprint) {
    $cert = Get-ChildItem Cert:\CurrentUser\My |
      Where-Object { $_.EnhancedKeyUsageList -match 'Code Signing' } |
      Select-Object -First 1
    if (-not $cert) {
      $cert = New-SelfSignedCertificate -Type CodeSigningCert `
        -Subject "CN=CustomsAppify Runner" `
        -CertStoreLocation Cert:\CurrentUser\My `
        -NotAfter (Get-Date).AddYears(3)
      Write-Host "==> created self-signed code-signing cert: $($cert.Thumbprint)"
    }
    $SignThumbprint = $cert.Thumbprint
  }

  if ($TrustSelfSigned) {
    $tmp = Join-Path $env:TEMP "customsappify-runner-cert.cer"
    Export-Certificate -Cert (Get-Item "Cert:\CurrentUser\My\$SignThumbprint") -FilePath $tmp | Out-Null
    Import-Certificate -FilePath $tmp -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    Import-Certificate -FilePath $tmp -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
    Write-Host "==> installed trust (Root + TrustedPublisher, CurrentUser)"
  }

  try {
    $sig = Set-AuthenticodeSignature -FilePath $exe `
      -Certificate (Get-Item "Cert:\CurrentUser\My\$SignThumbprint") `
      -HashAlgorithm SHA256 `
      -TimestampServer http://timestamp.digicert.com
  } catch {
    $sig = Set-AuthenticodeSignature -FilePath $exe `
      -Certificate (Get-Item "Cert:\CurrentUser\My\$SignThumbprint") `
      -HashAlgorithm SHA256
    Write-Host "==> signed without timestamp (timestamp server unreachable)"
  }
  Write-Host "==> signature status: $($sig.Status)"
  if ($sig.Status -ne 'Valid') {
    Write-Host "==> NOTE: signature not 'Valid'. Self-signed certs are only trusted"
    Write-Host "    on machines where it is installed (see -TrustSelfSigned) or after"
    Write-Host "    an enterprise policy trusts the publisher. For Smart App Control,"
    Write-Host "    a CA-issued code-signing cert + WDSI submission is required."
  }
}

Get-ChildItem (Join-Path $Publish $Name) | Format-List Name, Length
