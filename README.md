# GAWELA Tourenplaner

Desktop-Anwendung zur Tourenplanung mit MSIX/AppInstaller-Distribution über GitHub Releases.

## Aktueller Release

- Version: `1.0.6`
- Release: `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/tag/v1.0.6`
- AppInstaller-Download: `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.appinstaller`
- MSIX-Download: `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.msix`

## Enthalten in v1.0.6

- Kalender als eigene Seite
- Kalender-Icon in der Sidebar
- Korrigierte Umlaute auf Start- und Update-Seite
- Robuste Laufzeit-Pfadauflösung für Sidebar-Icons in der installierten Version
- Update-Menü im Programm
- Neues App-Icon
- Fix für die kurz sichtbare CMD/PowerShell-Konsole beim Start
- Fallback für deaktiviertes `ms-appinstaller`: Die App öffnet dann direkt die `.appinstaller`-Datei bzw. eine lokal heruntergeladene Kopie

## Installation auf Windows

1. `GAWELA-Tourenplaner.cer` öffnen
2. Zertifikat installieren
3. `Aktueller Benutzer` wählen
4. `Alle Zertifikate in folgendem Speicher speichern` wählen
5. In `Vertrauenswürdige Stammzertifizierungsstellen` importieren
6. Danach `GAWELA-Tourenplaner.appinstaller` öffnen

## Auto-Updates

- Für automatische Updates muss die App über die `.appinstaller`-Datei installiert werden.
- Die Update-Quelle zeigt auf `releases/latest/download/...` und liefert aktuell `1.0.6.0`.
- Die veröffentlichte `.appinstaller`-Datei enthält `OnLaunch` und `AutomaticBackgroundTask`.
- Im Programm steht dafür das Update-Menü zur Verfügung.

## Release-Prozess

1. Versionsnummer in `version.txt`, `installer-dist/GAWELA-Tourenplaner.appinstaller`, `installer-dist/msix-package/Package.appxmanifest` und `installer-dist/build-msix.ps1` erhöhen.
2. EXE mit `pyinstaller GAWELA-Tourenplaner.spec --noconfirm` neu bauen.
3. MSIX mit `powershell -ExecutionPolicy Bypass -File .\\installer-dist\\build-msix.ps1` erstellen.
4. GitHub-Release-Assets mit `powershell -ExecutionPolicy Bypass -File .\\installer-dist\\prepare-github-release.ps1 -Version "<VERSION>" -PackageName "GAWELA.Tourenplaner" -Publisher "CN=GAWELA" -MsixPath ".\\installer-dist\\GAWELA-Tourenplaner.msix"` vorbereiten.
5. Tag und GitHub-Release anlegen.
6. Genau diese beiden Dateien hochladen:
   `installer-dist\\release-assets\\GAWELA-Tourenplaner.appinstaller`
   `installer-dist\\release-assets\\GAWELA-Tourenplaner.msix`

## Hinweise

- Das MSIX muss mit derselben Package Identity und demselben Publisher signiert bleiben.
- Für Windows-Updates muss die Paketversion bei jedem Release steigen.
- Die Installation sollte immer über die `.appinstaller`-Datei getestet werden, nicht nur über das nackte `.msix`.
