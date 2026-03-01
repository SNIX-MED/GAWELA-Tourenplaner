from __future__ import annotations

# Release / Deployment Checklist
# - MSIX signieren und das Zertifikat auf Zielsystemen vertrauen.
# - `.msix` und `.appinstaller` ueber HTTPS hosten.
# - In der `.appinstaller` Datei UpdateSettings fuer OnLaunch/AutomaticBackgroundTask setzen.
# - Installation immer ueber die `.appinstaller` Datei testen, nicht nur ueber das nackte `.msix`.
# - Gruppenrichtlinien / App Installer Policies pruefen, falls `ms-appinstaller` deaktiviert ist.

APP_NAME = "GAWELA Tourenplaner"
APPINSTALLER_URL = "https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.appinstaller"
SUPPORT_URL = "https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases"
REQUIRE_MSIX_FOR_AUTOUPDATE = True
SHOW_UPDATE_PAGE_IN_MENU = True

# Optional package hints for PowerShell lookups. Leave empty if APP_NAME is sufficient.
PACKAGE_NAME_HINTS: tuple[str, ...] = ()
VERSION_FILE_NAME = "version.txt"
NETWORK_TIMEOUT_SECONDS = 3.0
