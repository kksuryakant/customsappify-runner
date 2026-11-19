# CustomsAppify Runner

Open-source local automation runner for the CustomsAppify web platform
(https://customsappify.nyixo.com).

The runner connects to the CustomsAppify server over WebSocket, receives
automation jobs, executes them locally with Selenium/Chrome, streams logs and
progress, and uploads result artifacts. Server code is fetched fresh per job
(sha256-verified bundle) so the binary stays small and logic ships instantly.

## Build

Windows (PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build_runner.ps1 -Version 1.0.17
```

macOS/Linux:

```bash
./scripts/build_runner.sh 1.0.17 macos-arm64
```

Requires Python 3.10–3.12; uses Nuitka `--onefile --standalone`.

## Signing

Signing is handled by **SignPath Foundation** (free for qualifying open-source
projects, Sectigo certificate — fully trusted by Windows/Smart App Control).
Signing runs in GitHub Actions on `runner-v*` tags. See
`.github/workflows/sign-runner.yml`.

Local development signing (self-hosted CA, not for public distribution):
`scripts/signing.ps1 new-ca / issue / sign`.

## Verify

```powershell
Get-AuthenticodeSignature .\customsappify-runner.exe   # Status: Valid
```

Or with osslsigncode: `osslsigncode verify -in file.exe`.

## Run

```powershell
.\customsappify-runner.exe --server https://customsappify.nyixo.com --token <RUNNER_TOKEN>
```

## License

MIT — see LICENSE. Copyright (c) 2026 Nyixo.
