"""Runner client: connects to the server and executes assigned jobs locally.

Selenium/Chrome run *on this machine*. The server only stores job metadata and
relays log/progress events.
"""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import logging
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.request
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

logger = logging.getLogger("runner")

# Repo root (web/ -> ..) so `web.modules.impl.*` import cleanly on the runner.
_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

# OS family the server uses to pick the right binary for self-update.
_PLATFORM = {"darwin": "macos", "win32": "windows", "linux": "linux"}.get(
    sys.platform, sys.platform
)

# ------------------------------------------------------------- self-update
# Local build version, bumped with every runner release. The server's
# /api/runner/version returns the latest published build; the runner downloads
# it and re-execs itself when they differ.
RUNNER_VERSION = "1.0.16"


def _self_update(server_url: str) -> bool:
    """Download the newest runner binary from the server and relaunch.

    Returns True when a relaunch is scheduled (caller must exit).
    """
    try:
        api = server_url.rstrip("/") + "/api/runner/version?platform=" + _PLATFORM
        with urllib.request.urlopen(api, timeout=15) as resp:
            payload = json.loads(resp.read().decode("utf-8", "ignore"))
        latest = str(payload.get("version", "")).strip()
        url = str(payload.get("download_url", "")).strip()
        if not latest or not url or latest == RUNNER_VERSION:
            return False

        logger.info("Update available: %s -> %s. Downloading %s ...", RUNNER_VERSION, latest, url)
        exe = Path(sys.argv[0]).resolve()
        if not exe.is_file():
            logger.warning("Self-update skipped: cannot resolve running binary path (%s).", exe)
            return False

        tmp_dir = Path(tempfile.mkdtemp(prefix="runner-update-"))
        archive = tmp_dir / ("runner.zip" if url.endswith(".zip") else "runner.bin")
        with urllib.request.urlopen(url, timeout=300) as resp, open(archive, "wb") as out:
            shutil.copyfileobj(resp, out)

        new_exe = exe
        if archive.suffix == ".zip":
            import zipfile

            with zipfile.ZipFile(archive) as zf:
                names = [n for n in zf.namelist() if not n.endswith("/") and "/" not in n.rstrip("/")]
                target = names[0] if names else None
                if not target:
                    logger.warning("Update archive empty; skipping.")
                    return False
                new_exe = tmp_dir / target
                with zf.open(target) as src, open(new_exe, "wb") as out:
                    shutil.copyfileobj(src, out)
        else:
            new_exe = archive

        new_exe.chmod(0o755)

        # Swap in place. Windows cannot overwrite a running .exe, so it stages
        # the new file next to itself and a wrapper finishes the swap on exit.
        if os.name == "nt":
            staged = exe.with_suffix(".new.exe")
            shutil.move(str(new_exe), str(staged))
            # Quote each original argv so paths with spaces survive; embed them
            # literally rather than using %* (which would also pick up the PID
            # argument and launch the new binary with a stray leading arg).
            return_args = " ".join(f'"{a}"' for a in sys.argv[1:])
            bat = tmp_dir / "apply-update.bat"
            bat.write_text(
                '@echo off\r\n'
                ':wait\r\n'
                'tasklist /fi "PID eq %1" | find "%1" >nul && (timeout /t 1 >nul & goto wait)\r\n'
                f'move /y "{staged}" "{exe}" >nul\r\n'
                f'start "" "{exe}" {return_args}\r\n'
                f'del "%~f0"\r\n',
                encoding="ascii",
            )
            subprocess.Popen(
                ["cmd", "/c", str(bat), str(os.getpid())],
                creationflags=subprocess.DETACHED_PROCESS,
                close_fds=True,
            )
            logger.info("Update staged; restarting after exit.")
            return True

        backup = exe.with_suffix(".old")
        shutil.move(str(exe), str(backup))
        shutil.move(str(new_exe), str(exe))
        exe.chmod(0o755)
        logger.info("Updated to %s. Relaunching ...", latest)
        os.execv(str(exe), [str(exe), *sys.argv[1:]])
        return True
    except Exception as exc:
        logger.warning("Self-update failed: %s", exc)
        return False


class RunnerClient:
    """Connects to the server, executes assigned modules on this machine."""

    def __init__(self, server_url: str, token: str, name: Optional[str] = None) -> None:
        self.server_url = server_url.rstrip("/")
        self.token = token
        self.name = name or socket.gethostname()
        self.machine_id = str(uuid.getnode())
        self.websocket: Optional[Any] = None
        self._current_job_id: Optional[int] = None
        self._bundle_dir: Optional[Path] = None
        self._lock = threading.Lock()
        # Admin-managed proxies fetched per job from the server, used to route all
        # runner HTTP (inputs, artifacts, credentials) and Selenium traffic so the
        # host IP is never exposed.
        self._proxy_proxies: List[Dict[str, Any]] = []
        self._proxy_opener: Optional[Any] = None

    # ------------------------------------------------------------------ ws url
    def _ws_url(self) -> str:
        base = self.server_url
        if base.startswith("https://"):
            base = "wss://" + base[len("https://"):]
        elif base.startswith("http://"):
            base = "ws://" + base[len("http://"):]
        return f"{base}/api/runner/ws?token={self.token}"

    # -------------------------------------------------------------- lifecycle
    async def run(self) -> None:
        import websockets

        if _self_update(self.server_url):
            return

        url = self._ws_url()
        backoff = 2
        while True:
            try:
                logger.info("Connecting to %s ...", url)
                async with websockets.connect(
                    url,
                    ping_interval=30,
                    ping_timeout=90,
                    max_size=512 * 1024 * 1024,
                    open_timeout=30,
                ) as websocket:
                    self.websocket = websocket
                    backoff = 2
                    await self._announce()
                    await self._receive_loop(websocket)
            except asyncio.CancelledError:
                logger.info("Runner shutting down.")
                return
            except Exception as exc:
                self.websocket = None
                logger.warning("Connection lost: %s. Reconnecting in %ds", exc, backoff)
                await asyncio.sleep(backoff)
                backoff = min(backoff * 2, 30)

    async def _announce(self) -> None:
        await self._send({
            "type": "announce",
            "name": self.name,
            "machine_id": self.machine_id,
            "platform": sys.platform,
        })

    @staticmethod
    def _print_banner(name: str) -> None:
        """First-run success block so the user knows everything worked."""
        line = "=" * 60
        logger.info(line)
        logger.info("RUNNER ONLINE - '%s'", name)
        logger.info("Connected to CustomsAppify server.")
        logger.info("This machine will now receive and execute jobs.")
        logger.info("Chrome will open here automatically when a job runs.")
        logger.info("Keep this window open. Press Ctrl+C to stop.")
        logger.info(line)

    async def _receive_loop(self, websocket) -> None:
        last_update_check = time.monotonic()
        last_heartbeat_log = time.monotonic()
        welcomed = False
        async for raw in websocket:
            try:
                message = json.loads(raw)
            except json.JSONDecodeError:
                continue
            msg_type = message.get("type")
            if msg_type == "job":
                asyncio.create_task(self._handle_job(message))
            elif msg_type == "cancel":
                self._cancel_job(message.get("job_id"))
            elif msg_type == "welcome":
                welcomed = True
                # Use the name we announced, not the server's stale stored value.
                self._print_banner(self.name)

            # Reassure the user the runner is alive (silent otherwise).
            if welcomed and time.monotonic() - last_heartbeat_log > 60:
                last_heartbeat_log = time.monotonic()
                job = self._current_job_id
                logger.info(
                    "Runner active - connected to %s, %s. Keep this window open.",
                    self.server_url,
                    f"executing job #{job}" if job is not None else "waiting for jobs",
                )

            # Re-check for runner updates roughly every 10 minutes.
            if time.monotonic() - last_update_check > 600:
                last_update_check = time.monotonic()
                if _self_update(self.server_url):
                    raise asyncio.CancelledError()

    async def _send(self, message: Dict[str, Any]) -> None:
        if self.websocket is None:
            return
        try:
            await self.websocket.send(json.dumps(message))
        except Exception as exc:
            logger.debug("Send failed: %s", exc)

    # ------------------------------------------------------------------- jobs
    async def _handle_job(self, message: Dict[str, Any]) -> None:
        job_id = int(message["job_id"])
        module_key = str(message["module_key"])
        params = message.get("params", {}) or {}
        title = message.get("title", module_key)

        with self._lock:
            if self._current_job_id is not None:
                await self._send({
                    "type": "status", "job_id": job_id, "status": "failed",
                    "error": "Runner is already executing another job",
                })
                return
            self._current_job_id = job_id

        logger.info(">>> Job #%s received: %s (%s)", job_id, title, module_key)
        await self._send({"type": "status", "job_id": job_id, "status": "running"})
        await self._send({"type": "progress", "job_id": job_id, "progress": 1,
                          "label": "Preparing local environment"})

        work_dir: Optional[Path] = None
        try:
            work_dir = Path(tempfile.mkdtemp(prefix=f"appify_job{job_id}_"))
            bundle_dir = await self._fetch_bundle(work_dir)
            self._activate_bundle(bundle_dir)
            uploaded = await self._download_inputs(job_id, params, work_dir)
            credentials = await self._fetch_credentials(job_id)
            self._fetch_proxies()
            await self._execute_module(job_id, module_key, title, params, work_dir,
                                       uploaded, credentials)
        except Exception as exc:
            logger.exception("Job %s failed", job_id)
            await self._send({"type": "log", "job_id": job_id, "level": "ERROR",
                              "message": f"Job failed: {exc}"})
            await self._send({"type": "status", "job_id": job_id, "status": "failed",
                              "error": str(exc)})
        finally:
            self._current_job_id = None
            if work_dir:
                await self._upload_artifacts(job_id, work_dir, params)
                self._deactivate_bundle()
                self._cleanup(work_dir)

    def _cancel_job(self, job_id: Optional[int]) -> None:
        with self._lock:
            if job_id is not None and self._current_job_id == int(job_id):
                logger.info("Cancellation requested for job %s (best-effort)", job_id)

    # ---------------------------------------------------------- feature bundle
    async def _fetch_bundle(self, work_dir: Path) -> Path:
        """Download current runner code from the server into this job's temp dir.

        No persistent copy: every job gets fresh code, and the work directory is
        deleted after artifacts are uploaded.
        """
        import hashlib
        import io
        import urllib.request
        import zipfile

        url = f"{self.server_url}/api/runner/bundle"
        request = urllib.request.Request(
            url, headers={"Authorization": f"Bearer {self.token}"}
        )
        with urllib.request.urlopen(request, timeout=300) as response:
            payload = response.read()
            expected = (response.headers.get("X-Bundle-Sha256") or "").strip()
        digest = hashlib.sha256(payload).hexdigest()
        if expected and digest != expected:
            raise RuntimeError(
                f"Runner code bundle checksum mismatch (got {digest}, expected {expected})"
            )

        bundle_dir = (work_dir / "bundle").resolve()
        bundle_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(payload)) as archive:
            for info in archive.infolist():
                target = (bundle_dir / info.filename).resolve()
                if target != bundle_dir and bundle_dir not in target.parents:
                    raise RuntimeError(f"Unsafe path in runner code bundle: {info.filename}")
            archive.extractall(bundle_dir)
        logger.info("Fetched runner code bundle (%d bytes, sha256=%s)", len(payload), digest[:12])
        return bundle_dir

    def _activate_bundle(self, bundle_dir: Path) -> None:
        """Put the freshly downloaded code ahead of the bundled fallback."""
        import web

        root = str(bundle_dir)
        if root not in sys.path:
            sys.path.insert(0, root)
        package_dir = str(bundle_dir / "web")
        try:
            if package_dir not in web.__path__:
                web.__path__.insert(0, package_dir)
        except AttributeError:
            web.__path__ = [package_dir]
        self._purge_module_cache()
        self._bundle_dir = bundle_dir

    def _deactivate_bundle(self) -> None:
        """Forget downloaded modules so the next job cannot reuse stale code."""
        self._purge_module_cache()
        if self._bundle_dir is not None:
            root = str(self._bundle_dir)
            try:
                sys.path.remove(root)
            except ValueError:
                pass
            try:
                import web

                package_dir = str(self._bundle_dir / "web")
                if package_dir in web.__path__:
                    web.__path__.remove(package_dir)
            except Exception:
                pass
            self._bundle_dir = None

    @staticmethod
    def _purge_module_cache() -> None:
        for name in list(sys.modules):
            if name == "web.modules" or name.startswith("web.modules."):
                sys.modules.pop(name, None)

    # --------------------------------------------------------------- inputs
    async def _download_inputs(self, job_id: int, params: Dict[str, Any],
                               work_dir: Path) -> Dict[str, List[Path]]:
        """Download files the user uploaded for this job from the server."""
        import urllib.request

        listing = await self._http_get(f"/api/jobs/{job_id}/inputs")
        uploaded: Dict[str, List[Path]] = {}
        inputs_dir = work_dir / "inputs"
        inputs_dir.mkdir(parents=True, exist_ok=True)

        for entry in listing:
            field_key = entry.get("field")
            filename = entry.get("filename")
            file_id = entry.get("id")
            if not filename or file_id is None:
                continue
            dest = inputs_dir / f"{file_id}_{Path(filename).name}"
            url = f"{self.server_url}/api/jobs/{job_id}/inputs/{file_id}"
            logger.info("Downloading input %s", filename)
            req = urllib.request.Request(
                url, headers={"Authorization": f"Bearer {self.token}"}
            )
            with urllib.request.urlopen(req, timeout=300) as response, dest.open("wb") as handle:
                handle.write(response.read())
            uploaded.setdefault(str(field_key), []).append(dest)

        # Resolve folder params (server-side absolute paths are not trusted; a
        # folder param is only honoured when it exists on this machine).
        resolved: Dict[str, Any] = {}
        for key, value in params.items():
            resolved[key] = value
        return uploaded

    async def _fetch_credentials(self, job_id: int) -> Dict[str, str]:
        payload = await self._http_get(f"/api/jobs/{job_id}/credentials")
        return payload if isinstance(payload, dict) else {}

    async def _http_get(self, path: str) -> Any:
        import urllib.request

        url = f"{self.server_url}{path}"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {self.token}"})
        with urllib.request.urlopen(req, timeout=120) as response:
            return json.loads(response.read().decode("utf-8"))

    # --------------------------------------------------------------- proxies
    def _fetch_proxies(self) -> None:
        """Fetch the admin-managed proxy pool for this runner (best-effort).

        A failure to fetch leaves the previous pool intact; jobs still run, they
        just use whatever proxies were last known.
        """
        try:
            url = f"{self.server_url}/api/runner/proxies?token={self.token}"
            with urllib.request.urlopen(url, timeout=30) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except Exception as exc:
            logger.warning("Proxy fetch failed: %s (using last known pool)", exc)
            return
        proxies = payload.get("proxies") if isinstance(payload, dict) else None
        if not isinstance(proxies, list):
            return
        self._proxy_proxies = [p for p in proxies if isinstance(p, dict)]
        self._proxy_opener = self._build_proxy_opener(self._proxy_proxies)
        logger.info("Loaded %d proxy(ies) for runner traffic", len(self._proxy_proxies))

    @staticmethod
    def _build_proxy_opener(proxies: List[Dict[str, Any]]) -> Optional[Any]:
        """urllib opener that routes through the first usable pool proxy."""
        import urllib.request

        for entry in proxies:
            url = str(entry.get("url") or "").strip()
            if not url:
                continue
            try:
                handler = urllib.request.ProxyHandler({
                    "http": url,
                    "https": url,
                })
                return urllib.request.build_opener(handler)
            except Exception as exc:  # pragma: no cover - defensive
                logger.warning("Ignoring unusable proxy %s: %s", url, exc)
        return None

    def _urlopen(self, request: Any, timeout: int):
        """urlopen through the pool proxy when one is configured, else direct."""
        if self._proxy_opener is not None:
            return self._proxy_opener.open(request, timeout=timeout)
        return urllib.request.urlopen(request, timeout=timeout)

    # ------------------------------------------------------------- execution
    async def _execute_module(self, job_id: int, module_key: str, title: str,
                              params: Dict[str, Any], work_dir: Path,
                              uploaded: Dict[str, List[Path]],
                              credentials: Dict[str, str]) -> None:
        # Deliberately dynamic: module code comes from the per-job bundle in
        # the temp directory, never from the compiled bootstrap executable.
        import importlib

        base = importlib.import_module("web.modules.base")
        registry = importlib.import_module("web.modules.registry")
        ModuleContext = base.ModuleContext
        ModuleError = base.ModuleError
        implementation = registry.get_implementation(module_key)
        if implementation is None:
            raise ModuleError(
                f"Module '{module_key}' has no runner implementation yet. "
                "It is listed in the catalog but not ported to the web runner."
            )

        loop = asyncio.get_running_loop()
        send_lock = asyncio.Lock()

        async def emit_log(message: str, level: str = "INFO") -> None:
            async with send_lock:
                await self._send({
                    "type": "log", "job_id": job_id, "level": level, "message": message,
                })

        async def emit_progress(value: int, label: str = "") -> None:
            async with send_lock:
                await self._send({
                    "type": "progress", "job_id": job_id, "progress": value, "label": label,
                })

        def sync_log(message: str) -> None:
            asyncio.run_coroutine_threadsafe(emit_log(message), loop)

        def sync_progress(value: int, label: str = "") -> None:
            asyncio.run_coroutine_threadsafe(emit_progress(value, label), loop)

        await emit_log(f"Starting module '{module_key}' on {socket.gethostname()}")

        def run_blocking() -> None:
            ctx = ModuleContext(
                job_id=job_id,
                title=title,
                params=params,
                work_dir=work_dir,
                uploaded=uploaded,
                credentials=credentials,
                log=sync_log,
                progress=sync_progress,
                proxies=list(self._proxy_proxies),
            )
            implementation.execute(ctx)

        await asyncio.to_thread(run_blocking)

        await emit_log("Module execution finished")
        await self._send({"type": "status", "job_id": job_id, "status": "completed"})

    # ------------------------------------------------------------- artifacts
    async def _upload_artifacts(self, job_id: int, work_dir: Path,
                                params: Dict[str, Any]) -> None:
        import urllib.request

        candidates: List[Path] = []
        artifacts_dir = params.get("_artifacts_dir")
        if artifacts_dir:
            target = Path(artifacts_dir)
            candidates.extend(sorted(target.glob("*")) if target.is_dir() else [])
        else:
            inputs_dir = work_dir / "inputs"
            candidates.extend(
                sorted(
                    p for p in work_dir.rglob("*")
                    if p.is_file() and inputs_dir not in p.parents
                )
            )

        for path in candidates:
            if not path.is_file() or path.stat().st_size == 0:
                continue
            try:
                await self._upload_file(job_id, path)
            except Exception as exc:
                logger.warning("Failed to upload %s: %s", path.name, exc)

    async def _upload_file(self, job_id: int, path: Path) -> None:
        import urllib.request

        boundary = uuid.uuid4().hex
        filename = path.name
        with path.open("rb") as handle:
            payload = handle.read()
        parts = []
        parts.append(
            f"--{boundary}\r\n"
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f"Content-Type: application/octet-stream\r\n\r\n"
            .encode("utf-8")
        )
        parts.append(payload)
        parts.append(f"\r\n--{boundary}--\r\n".encode("utf-8"))
        body = b"".join(parts)

        url = f"{self.server_url}/api/jobs/{job_id}/artifacts"
        request = urllib.request.Request(
            url,
            data=body,
            headers={
                "Content-Type": f"multipart/form-data; boundary={boundary}",
                "Authorization": f"Bearer {self.token}",
            },
            method="POST",
        )
        with urllib.request.urlopen(request, timeout=600) as response:
            response.read()

    def _cleanup(self, work_dir: Path) -> None:
        def _remove() -> None:
            time.sleep(5)
            shutil.rmtree(work_dir, ignore_errors=True)

        threading.Thread(target=_remove, daemon=True).start()


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="CustomsAppify local runner - executes automation jobs on this machine.",
    )
    parser.add_argument("--server", required=True,
                        help="Server base URL, e.g. https://app.example.com")
    parser.add_argument("--token", required=True, help="Runner token issued by the server")
    parser.add_argument("--name", default=None, help="Friendly runner name")
    parser.add_argument("--log-level", default="INFO")

    args = parser.parse_args(argv)
    logging.basicConfig(
        level=args.log_level.upper(),
        format="[%(asctime)s] [%(name)s] %(levelname)s %(message)s",
        datefmt="%H:%M:%S",
    )

    client = RunnerClient(args.server, args.token, args.name)
    try:
        asyncio.run(client.run())
    except KeyboardInterrupt:
        logger.info("Interrupted by user")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
