"""Local runner client.

Runs on each user's machine, connects to the CustomsAppify server over a
websocket, and executes assigned automation modules *locally* — Selenium and
Chrome run here, never on the server.

Usage:
    python -m web.runner.client --server https://app.example.com --token RUNNER_TOKEN
"""

from .client import RunnerClient, main

__all__ = ["RunnerClient", "main"]
