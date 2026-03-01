from __future__ import annotations

import importlib.metadata
import json
import logging
import os
import socket
import subprocess
import sys
import tempfile
import webbrowser
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlparse
from urllib.request import Request, urlopen

from config.update_config import (
    APPINSTALLER_URL,
    APP_NAME,
    NETWORK_TIMEOUT_SECONDS,
    PACKAGE_NAME_HINTS,
    REQUIRE_MSIX_FOR_AUTOUPDATE,
    SUPPORT_URL,
    VERSION_FILE_NAME,
)

__version__ = "0.0.0-dev"

_LOGGER = logging.getLogger("update_service")
_LOGGER.setLevel(logging.INFO)
_LOGGER.propagate = False
_FILE_HANDLER: logging.Handler | None = None


def set_update_log_dir(log_dir: str | Path | None) -> None:
    global _FILE_HANDLER

    if log_dir is None:
        return

    target_dir = Path(log_dir)
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
    except OSError:
        return

    log_file = target_dir / "update.log"
    if _FILE_HANDLER is not None:
        existing = getattr(_FILE_HANDLER, "baseFilename", "")
        if existing and Path(existing) == log_file:
            return
        _LOGGER.removeHandler(_FILE_HANDLER)
        try:
            _FILE_HANDLER.close()
        except Exception:
            pass
        _FILE_HANDLER = None

    try:
        handler = logging.FileHandler(log_file, encoding="utf-8")
    except OSError:
        return

    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    _LOGGER.addHandler(handler)
    _FILE_HANDLER = handler


def get_app_version() -> str:
    package = _detect_msix_package()
    if package and package.get("Version"):
        return str(package["Version"])

    version_file = _read_version_file()
    if version_file:
        return version_file

    metadata_version = _read_metadata_version()
    if metadata_version:
        return metadata_version

    return __version__


def get_installation_type() -> str:
    package = _detect_msix_package()
    if package:
        return "msix"

    runtime_path = _runtime_path()
    if runtime_path.exists():
        return "portable"

    return "unknown"


def get_runtime_update_context() -> dict[str, Any]:
    package = _detect_msix_package()
    installation_type = "msix" if package else get_installation_type()
    version = str(package.get("Version")) if package and package.get("Version") else get_app_version()
    package_family_name = str(package.get("PackageFamilyName") or "").strip() if package else ""
    install_location = str(package.get("InstallLocation") or "").strip() if package else ""

    return {
        "app_name": APP_NAME,
        "installation_type": installation_type,
        "version": version,
        "package_family_name": package_family_name,
        "package_name": str(package.get("Name") or "").strip() if package else "",
        "install_location": install_location,
        "appinstaller_url": APPINSTALLER_URL,
        "support_url": SUPPORT_URL,
        "requires_msix_for_autoupdate": bool(REQUIRE_MSIX_FOR_AUTOUPDATE),
        "runtime_path": str(_runtime_path()),
        "is_msix": installation_type == "msix",
    }


def check_update_source_reachable(timeout_seconds: float = NETWORK_TIMEOUT_SECONDS) -> dict[str, Any]:
    if not APPINSTALLER_URL:
        return {"ok": False, "detail": "Keine AppInstaller-URL konfiguriert."}

    parsed = urlparse(APPINSTALLER_URL)
    if not parsed.scheme or not parsed.netloc:
        return {"ok": False, "detail": "Die AppInstaller-URL ist ungueltig."}

    request = Request(APPINSTALLER_URL, method="HEAD")
    try:
        with urlopen(request, timeout=timeout_seconds) as response:
            return {"ok": True, "detail": f"HTTP {getattr(response, 'status', 200)}"}
    except HTTPError as exc:
        if exc.code in {200, 204, 301, 302, 303, 307, 308, 403, 405}:
            return {"ok": True, "detail": f"HTTP {exc.code}"}
    except URLError as exc:
        _LOGGER.warning("HEAD request to AppInstaller URL failed: %s", exc)
    except Exception:
        _LOGGER.exception("Unexpected error during update source check.")

    host = parsed.hostname
    if not host:
        return {"ok": False, "detail": "Host konnte nicht ermittelt werden."}

    try:
        socket.getaddrinfo(host, parsed.port or 443)
        return {"ok": True, "detail": "DNS-Aufloesung erfolgreich"}
    except OSError as exc:
        _LOGGER.warning("DNS check for AppInstaller host failed: %s", exc)
        return {"ok": False, "detail": "Keine Netzwerkverbindung oder Host nicht erreichbar."}


def build_ms_appinstaller_uri() -> str:
    return f"ms-appinstaller:?source={quote(APPINSTALLER_URL, safe=':/?=&%')}"


def trigger_update_installation(*, prefer_appinstaller: bool = True) -> dict[str, Any]:
    if not APPINSTALLER_URL:
        return {"ok": False, "via": "none", "detail": "Keine AppInstaller-URL konfiguriert."}

    connectivity = check_update_source_reachable()
    if not connectivity.get("ok"):
        return {"ok": False, "via": "none", "detail": "Kein Update moeglich: Keine Internetverbindung."}

    protocol_enabled = _is_ms_appinstaller_protocol_enabled()

    if prefer_appinstaller and protocol_enabled:
        uri = build_ms_appinstaller_uri()
        try:
            os.startfile(uri)
            _LOGGER.info("Triggered App Installer via protocol URI.")
            return {"ok": True, "via": "ms-appinstaller", "detail": "App Installer wurde geoeffnet."}
        except Exception as exc:
            _LOGGER.warning("ms-appinstaller protocol failed: %s", exc)
    elif prefer_appinstaller:
        _LOGGER.info("ms-appinstaller protocol skipped because it is disabled on this system.")

    try:
        os.startfile(APPINSTALLER_URL)
        _LOGGER.info("Opened AppInstaller URL directly after protocol fallback.")
        return {
            "ok": True,
            "via": "appinstaller-url",
            "detail": (
                "Die .appinstaller-Datei wurde geoeffnet."
                if not prefer_appinstaller
                else "Das ms-appinstaller-Protokoll ist deaktiviert oder nicht verfuegbar. Die .appinstaller-Datei wurde direkt geoeffnet."
            ),
        }
    except Exception as exc:
        _LOGGER.warning("Direct open of AppInstaller URL failed: %s", exc)

    local_copy = _download_appinstaller_to_temp()
    if local_copy:
        try:
            os.startfile(str(local_copy))
            _LOGGER.info("Opened downloaded local AppInstaller file: %s", local_copy)
            return {
                "ok": True,
                "via": "appinstaller-file",
                "detail": "Die .appinstaller-Datei wurde lokal heruntergeladen und geoeffnet.",
            }
        except Exception as exc:
            _LOGGER.warning("Opening downloaded AppInstaller file failed: %s", exc)

    opened = webbrowser.open(APPINSTALLER_URL)
    if opened:
        _LOGGER.info("Opened AppInstaller URL via default browser.")
        return {
            "ok": True,
            "via": "browser",
            "detail": (
                "Die .appinstaller-URL wurde im Browser geoeffnet."
                if not prefer_appinstaller
                else "Das ms-appinstaller-Protokoll war nicht verfuegbar. Die .appinstaller-URL wurde im Browser geoeffnet."
            ),
        }

    return {"ok": False, "via": "none", "detail": "Update konnte nicht gestartet werden."}


def open_support_url() -> bool:
    if not SUPPORT_URL:
        return False
    try:
        os.startfile(SUPPORT_URL)
        return True
    except Exception:
        _LOGGER.warning("Support URL could not be opened via os.startfile.", exc_info=True)
    return bool(webbrowser.open(SUPPORT_URL))


def is_auto_update_settings_supported() -> bool:
    return _powershell_command_exists("Get-AppxPackageAutoUpdateSettings") and _powershell_command_exists(
        "Set-AppxPackageAutoUpdateSettings"
    )


def get_auto_update_settings(package_family_name: str, *, show_update_availability: bool = False) -> dict[str, Any] | None:
    package_family_name = str(package_family_name or "").strip()
    if not package_family_name or not is_auto_update_settings_supported():
        return None

    availability_flag = "-ShowUpdateAvailability" if show_update_availability else ""
    script = (
        f"$settings = Get-AppxPackageAutoUpdateSettings -PackageFamilyName '{_escape_ps(package_family_name)}' "
        f"{availability_flag} -ErrorAction Stop; "
        "if ($settings) { $settings | ConvertTo-Json -Depth 5 -Compress }"
    )
    payload = _run_powershell_json(script, timeout=12, raise_on_error=False)
    return payload if isinstance(payload, dict) else None


def set_auto_update_check_on_launch(package_family_name: str, enabled: bool) -> tuple[bool, str]:
    package_family_name = str(package_family_name or "").strip()
    if not package_family_name:
        return False, "Keine PackageFamilyName verfuegbar."
    if not APPINSTALLER_URL:
        return False, "Keine AppInstaller-URL konfiguriert."
    if not is_auto_update_settings_supported():
        return False, "Die AutoUpdate-Cmdlets sind auf diesem System nicht verfuegbar."

    script = (
        f"Set-AppxPackageAutoUpdateSettings -PackageFamilyName '{_escape_ps(package_family_name)}' "
        f"-AppInstallerUri '{_escape_ps(APPINSTALLER_URL)}' "
        f"-CheckOnLaunch ${str(bool(enabled)).lower()} -ErrorAction Stop | Out-Null"
    )
    try:
        _run_powershell(script, timeout=8)
    except RuntimeError as exc:
        _LOGGER.warning("Could not update CheckOnLaunch setting: %s", exc)
        return False, str(exc)
    return True, "Update-Einstellung gespeichert."


def format_auto_update_settings(settings: dict[str, Any] | None) -> str:
    if not settings:
        return "Keine AutoUpdate-Einstellungen verfuegbar."

    preferred_keys = (
        "PackageFamilyName",
        "AppInstallerUri",
        "CheckOnLaunch",
        "HoursBetweenUpdateChecks",
        "ShowPrompt",
        "UpdateBlocksActivation",
        "EnableAutomaticBackgroundTask",
        "ForceUpdateFromAnyVersion",
        "DisableAutoRepairs",
        "UpdateAvailability",
        "LastChecked",
    )
    lines = []
    for key in preferred_keys:
        if key in settings:
            lines.append(f"{key}: {settings.get(key)}")
    if not lines:
        lines = [f"{key}: {value}" for key, value in settings.items()]
    return "\n".join(lines)


def _runtime_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    argv0 = str(sys.argv[0] or "").strip()
    if argv0:
        try:
            return Path(argv0).resolve()
        except OSError:
            pass
    return Path(__file__).resolve()


def _detect_msix_package() -> dict[str, Any] | None:
    runtime_path = _runtime_path()
    package = _find_package_by_install_location(runtime_path)
    if package:
        return package

    if _path_looks_like_msix(runtime_path):
        package = _find_package_by_hints()
        if package:
            return package

    return None


def _find_package_by_install_location(runtime_path: Path) -> dict[str, Any] | None:
    runtime = str(runtime_path)
    script = (
        f"$runtimePath = '{_escape_ps(runtime)}'; "
        "$pkg = Get-AppxPackage | Where-Object { "
        "$_.InstallLocation -and "
        "$runtimePath.StartsWith($_.InstallLocation, [System.StringComparison]::OrdinalIgnoreCase) "
        "} | Select-Object -First 1 Name, PackageFamilyName, Version, InstallLocation, PackageFullName; "
        "if ($pkg) { $pkg | ConvertTo-Json -Depth 4 -Compress }"
    )
    payload = _run_powershell_json(script, timeout=12, raise_on_error=False)
    return payload if isinstance(payload, dict) else None


def _find_package_by_hints() -> dict[str, Any] | None:
    hints = _package_hints()
    if not hints:
        return None

    hint_list = ",".join(f"'{_escape_ps(item)}'" for item in hints)
    script = (
        f"$hints = @({hint_list}); "
        "$pkg = Get-AppxPackage | Where-Object { "
        "$combined = ($_.Name + ' ' + $_.PackageFamilyName); "
        "$matched = $false; "
        "foreach ($hint in $hints) { if ($combined -like ('*' + $hint + '*')) { $matched = $true; break } } "
        "$matched "
        "} | Select-Object -First 1 Name, PackageFamilyName, Version, InstallLocation, PackageFullName; "
        "if ($pkg) { $pkg | ConvertTo-Json -Depth 4 -Compress }"
    )
    payload = _run_powershell_json(script, timeout=12, raise_on_error=False)
    return payload if isinstance(payload, dict) else None


def _package_hints() -> list[str]:
    hints: list[str] = []
    for item in (APP_NAME, *PACKAGE_NAME_HINTS):
        text = str(item or "").strip()
        if not text:
            continue
        hints.append(text)
        for token in text.replace("_", " ").replace("-", " ").split():
            token = token.strip()
            if len(token) >= 3:
                hints.append(token)

    unique = []
    seen = set()
    for hint in hints:
        key = hint.casefold()
        if key in seen:
            continue
        seen.add(key)
        unique.append(hint)
    return unique


def _read_version_file() -> str | None:
    search_roots = []
    runtime_path = _runtime_path()
    search_roots.append(runtime_path.parent)
    search_roots.append(Path(__file__).resolve().parents[1])
    search_roots.append(Path.cwd())

    seen = set()
    for root in search_roots:
        try:
            candidate = (root / VERSION_FILE_NAME).resolve()
        except OSError:
            continue
        if candidate in seen:
            continue
        seen.add(candidate)
        try:
            if candidate.exists():
                value = candidate.read_text(encoding="utf-8").strip()
                if value:
                    return value
        except OSError:
            _LOGGER.warning("Version file could not be read: %s", candidate)
    return None


def _read_metadata_version() -> str | None:
    for name in (APP_NAME, *PACKAGE_NAME_HINTS):
        package_name = str(name or "").strip()
        if not package_name:
            continue
        try:
            return importlib.metadata.version(package_name)
        except importlib.metadata.PackageNotFoundError:
            continue
        except Exception:
            _LOGGER.warning("Metadata version lookup failed for %s", package_name, exc_info=True)
    return None


def _is_ms_appinstaller_protocol_enabled() -> bool:
    if os.name != "nt":
        return False

    try:
        import winreg
    except ImportError:
        return False

    key_path = r"SOFTWARE\Policies\Microsoft\Windows\AppInstaller"
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            value, _ = winreg.QueryValueEx(key, "EnableMSAppInstallerProtocol")
            return bool(int(value))
    except OSError:
        return False


def _download_appinstaller_to_temp() -> Path | None:
    if not APPINSTALLER_URL:
        return None

    try:
        request = Request(APPINSTALLER_URL)
        with urlopen(request, timeout=max(10.0, NETWORK_TIMEOUT_SECONDS)) as response:
            payload = response.read()
    except Exception:
        _LOGGER.warning("Downloading AppInstaller file failed.", exc_info=True)
        return None

    try:
        target_dir = Path(tempfile.gettempdir()) / "gawela-appinstaller"
        target_dir.mkdir(parents=True, exist_ok=True)
        target_path = target_dir / "GAWELA-Tourenplaner.appinstaller"
        target_path.write_bytes(payload)
        return target_path
    except OSError:
        _LOGGER.warning("Writing downloaded AppInstaller file failed.", exc_info=True)
        return None


def _powershell_command_exists(command_name: str) -> bool:
    script = f"if (Get-Command {command_name} -ErrorAction SilentlyContinue) {{ 'yes' }}"
    try:
        output = _run_powershell(script, timeout=4, raise_on_error=False)
    except Exception:
        return False
    return "yes" in output.lower()


def _run_powershell_json(script: str, *, timeout: int = 5, raise_on_error: bool = True) -> Any:
    output = _run_powershell(script, timeout=timeout, raise_on_error=raise_on_error)
    payload = str(output or "").strip()
    if not payload:
        return None
    try:
        return json.loads(payload)
    except json.JSONDecodeError:
        _LOGGER.warning("PowerShell JSON payload could not be parsed: %s", payload)
        return None


def _run_powershell(script: str, *, timeout: int = 5, raise_on_error: bool = True) -> str:
    startupinfo = None
    creationflags = 0
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    try:
        completed = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
    except subprocess.TimeoutExpired:
        message = f"PowerShell-Aufruf hat das Zeitlimit von {timeout}s ueberschritten."
        if raise_on_error:
            raise RuntimeError(message)
        _LOGGER.warning(message)
        return ""
    except OSError as exc:
        message = f"PowerShell konnte nicht gestartet werden: {exc}"
        if raise_on_error:
            raise RuntimeError(message)
        _LOGGER.warning(message)
        return ""
    stdout = str(completed.stdout or "").strip()
    stderr = str(completed.stderr or "").strip()
    if completed.returncode != 0 and raise_on_error:
        raise RuntimeError(stderr or stdout or "PowerShell-Aufruf fehlgeschlagen.")
    if completed.returncode != 0:
        _LOGGER.warning("PowerShell command failed: %s", stderr or stdout)
    return stdout


def _path_looks_like_msix(path: Path) -> bool:
    normalized = str(path).casefold()
    return "windowsapps" in normalized or "\\msixvc\\" in normalized


def _escape_ps(value: str) -> str:
    return str(value or "").replace("'", "''")
