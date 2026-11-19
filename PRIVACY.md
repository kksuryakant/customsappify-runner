# Privacy Policy — CustomsAppify Runner

Last updated: 2026-11-19

CustomsAppify Runner is an open-source local client for the CustomsAppify web
platform. This policy describes what the runner itself does with data.

## What the runner does

- Connects, on your instruction, to the CustomsAppify server URL and runner
  token you provide.
- Receives automation jobs from that server and executes them **on your own
  machine** (including opening a local browser, e.g. Chrome, for automation).
- Streams job logs and progress back to that server.
- Uploads result artifacts (e.g. generated Excel/PDF files) to that server.

## Data the runner sends

- The runner token you supply, the machine name, and a machine identifier
  (used to register with the server).
- Job logs and output artifacts, sent only to the server you configured.
- No telemetry, analytics, or usage statistics are sent to the project
  maintainers or any third party by the runner itself. The runner contains no
  tracking SDKs.

## Data the runner does NOT collect

- The runner does not collect personal information, browsing history, or files
  outside of what the job explicitly asks it to process.
- The runner does not transmit data to any third party other than the server
  you explicitly configure.

## Your control

- You choose which server and token the runner uses. You may stop the runner at
  any time.
- Job data and artifacts are governed by the privacy policy and terms of the
  server you connect to.

## Contact

Questions about this policy: open an issue in this repository.

## License

Code: MIT. This privacy statement covers the runner component only.
