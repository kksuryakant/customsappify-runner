"""Shared browser factory for runner-side modules.

Every driver produced here runs on the *user's local machine* (the runner host),
never on the server. Headless mode is the default for server-driven jobs, but a
runner may expose a visible browser when a flow needs manual interaction such as
OTP entry.
"""

from __future__ import annotations

import logging
import os
from typing import Optional

logger = logging.getLogger(__name__)


def _env_proxy() -> Optional[str]:
    """Proxy from APPIFY_PROXY / HTTPS_PROXY / HTTP_PROXY (legacy/env fallback)."""
    proxy = (
        os.environ.get("APPIFY_PROXY")
        or os.environ.get("HTTPS_PROXY")
        or os.environ.get("HTTP_PROXY")
    )
    return proxy.strip() if proxy else None


def build_chrome_options(headless: bool = True, proxy: Optional[str] = None, no_proxy: bool = False):
    from selenium.webdriver.chrome.options import Options

    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-plugins")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--lang=en-US,en")
    chosen = None if no_proxy else (proxy or _env_proxy())
    if chosen:
        if not chosen.startswith("socks"):
            # Chrome wants a scheme; bare host:port is treated as http.
            if "://" not in chosen:
                chosen = "http://" + chosen
        options.add_argument(f"--proxy-server={chosen}")
        logger.info("Browser using proxy: %s", chosen)
    prefs = {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "download.prompt_for_download": False,
    }
    options.add_experimental_option("prefs", prefs)
    return options


def create_driver(headless: bool = True, max_retries: int = 3, user_data_dir: Optional[str] = None, proxy: Optional[str] = None, no_proxy: bool = False):
    """Create a Selenium Chrome driver, preferring undetected-chromedriver.

    Falls back to webdriver-manager then plain selenium. The driver is local to
    whichever machine the runner is executing on. Pass ``user_data_dir`` to reuse
    a persistent Chrome profile (e.g. to keep an anti-bot session cookie).
    ``proxy`` is a full ``scheme://[user:pass@]host:port`` URL from the platform
    pool; falls back to APPIFY_PROXY / HTTP(S)_PROXY env otherwise.
    """
    import time

    options = build_chrome_options(headless=headless, proxy=proxy, no_proxy=no_proxy)
    errors: list[str] = []

    # 1) undetected-chromedriver (best stealth; mirrors the desktop app)
    try:
        import undetected_chromedriver as uc

        # Pin the driver to the installed Chrome major. Without this UC fetches
        # the newest driver (e.g. 156) and fails on a Chrome one major behind
        # (155) with "This version of ChromeDriver only supports Chrome version".
        uc_kwargs = {}
        chrome_major = _installed_chrome_version()
        if chrome_major and chrome_major.isdigit():
            uc_kwargs["version_main"] = int(chrome_major)
        driver = uc.Chrome(
            options=options, use_subprocess=True, user_data_dir=user_data_dir, **uc_kwargs
        )
        _patch_webdriver_flag(driver)
        logger.info("Browser initialised via undetected-chromedriver (headless=%s)", headless)
        return driver
    except Exception as exc:  # pragma: no cover - environment dependent
        errors.append(f"undetected-chromedriver: {exc}")
        logger.debug("undetected-chromedriver unavailable: %s", exc)

    # 2) webdriver-manager
    if user_data_dir:
        options.add_argument(f"--user-data-dir={user_data_dir}")
    for attempt in range(1, max_retries + 1):
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from webdriver_manager.chrome import ChromeDriverManager

            driver = webdriver.Chrome(
                service=Service(_matching_chromedriver(ChromeDriverManager)), options=options
            )
            _patch_webdriver_flag(driver)
            logger.info("Browser initialised via webdriver-manager (headless=%s)", headless)
            return driver
        except Exception as exc:  # pragma: no cover - environment dependent
            errors.append(f"webdriver-manager attempt {attempt}: {exc}")
            if attempt < max_retries:
                time.sleep(2 * attempt)

    # 3) plain selenium with a bundled/system chromedriver
    try:
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service

        service = None
        for candidate in ("/opt/homebrew/bin/chromedriver", "/usr/local/bin/chromedriver",
                          "/usr/bin/chromedriver"):
            if os.path.exists(candidate):
                service = Service(candidate)
                break
        driver = webdriver.Chrome(service=service, options=options)
        _patch_webdriver_flag(driver)
        logger.info("Browser initialised via system chromedriver (headless=%s)", headless)
        return driver
    except Exception as exc:  # pragma: no cover - environment dependent
        errors.append(f"system chromedriver: {exc}")

    raise RuntimeError(
        "Could not start a local Chrome browser on this machine. Tried: " + "; ".join(errors)
    )


def _installed_chrome_version() -> Optional[str]:
    """Major version of the locally installed Chrome, or None if unknown.

    webdriver-manager defaults to the *latest* chromedriver, which breaks when
    the installed Chrome is one major behind (e.g. driver 156 vs Chrome 155) and
    the runner cannot reach the internet to fetch a matching driver. Detecting
    the installed Chrome lets us request the matching driver instead.
    """
    override = (os.environ.get("APPIFY_CHROME_VERSION") or "").strip()
    if override:
        return override
    import re
    import subprocess as _sp

    def _major(text: str) -> Optional[str]:
        m = re.search(r"(\d+)\.\d+\.\d+\.\d+", text or "")
        return m.group(1) if m else None

    # Windows: registry holds the version without launching Chrome.
    if os.name == "nt":
        try:
            import winreg  # type: ignore

            for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                for key in (
                    r"SOFTWARE\Google\Chrome\BLBeacon",
                    r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon",
                ):
                    try:
                        with winreg.OpenKey(root, key) as h:
                            v, _ = winreg.QueryValueEx(h, "version")
                            major = _major(str(v))
                            if major:
                                return major
                    except OSError:
                        continue
        except Exception:  # noqa: BLE001
            pass

    # Fallback: ask the Chrome binary directly.
    for binary in (
        "google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    ):
        try:
            out = _sp.run([binary, "--version"], capture_output=True, text=True, timeout=10)
            major = _major(out.stdout or out.stderr)
            if major:
                return major
        except Exception:  # noqa: BLE001
            continue
    return None


def _matching_chromedriver(manager_cls) -> str:
    """Resolve a chromedriver path matching the installed Chrome version.

    Pins webdriver-manager to the detected Chrome major so a runner with Chrome
    155 gets driver 155 rather than the newest (156). Falls back to the default
    latest-driver behaviour when the Chrome version cannot be detected.
    """
    version = _installed_chrome_version()
    if version:
        try:
            path = manager_cls(driver_version=version).install()
            logger.info("Resolved chromedriver for Chrome %s", version)
            return path
        except Exception as exc:  # noqa: BLE001
            logger.warning("chromedriver %s unavailable (%s); using latest", version, exc)
    return manager_cls().install()


def _patch_webdriver_flag(driver) -> None:
    try:
        driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
    except Exception:
        pass
    _spoof_headless_user_agent(driver)


def _spoof_headless_user_agent(driver) -> None:
    """Hide the ``HeadlessChrome`` UA token.

    Sites (e.g. maersk.com) serve an empty "no results" page when the
    User-Agent contains ``HeadlessChrome``. Overriding it to a normal Chrome UA
    keeps headless jobs functioning. The version is taken from the live session
    so the string stays consistent with the actual browser.
    """
    try:
        version = (driver.capabilities or {}).get("browserVersion") or "131.0.0.0"
        user_agent = (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            f"Chrome/{version} Safari/537.36"
        )
        driver.execute_cdp_cmd(
            "Network.setUserAgentOverride",
            {"userAgent": user_agent, "platform": "MacIntel", "acceptLanguage": "en-US,en"},
        )
        driver.execute_cdp_cmd("Network.enable", {})
    except Exception:
        pass


def safe_quit(driver: Optional[object]) -> None:
    if driver is None:
        return
    try:
        driver.quit()
    except Exception:
        pass
