# GAWELA Tourenplaner

Desktop-Anwendung zur Tourenplanung mit MSIX/AppInstaller-Distribution ueber GitHub Releases.

## Aktueller Release

- Version: `1.0.4`
- Release: `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/tag/v1.0.4`
- AppInstaller-Download: `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.appinstaller`
- MSIX-Download: `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.msix`

## Enthalten in v1.0.4

- Update-Menue im Programm
- Neues App-Icon
- Fix fuer die kurz sichtbare CMD/PowerShell-Konsole beim Start
- Fallback fuer deaktiviertes `ms-appinstaller`: Die App oeffnet dann direkt die `.appinstaller`-Datei bzw. eine lokal heruntergeladene Kopie

## Installation auf Windows

1. `GAWELA-Tourenplaner.cer` oeffnen
2. Zertifikat installieren
3. `Aktueller Benutzer` waehlen
4. `Alle Zertifikate in folgendem Speicher speichern` waehlen
5. In `Vertrauenswuerdige Stammzertifizierungsstellen` importieren
6. Danach `GAWELA-Tourenplaner.appinstaller` oeffnen

## Auto-Updates

- Fuer automatische Updates muss die App ueber die `.appinstaller`-Datei installiert werden.
- Die Update-Quelle zeigt auf `releases/latest/download/...` und liefert aktuell `1.0.4.0`.
- Die veroeffentlichte `.appinstaller`-Datei enthaelt `OnLaunch` und `AutomaticBackgroundTask`.
- Im Programm steht dafuer das Update-Menue zur Verfuegung.

## Release-Prozess

1. Versionsnummer in `version.txt`, `installer-dist/GAWELA-Tourenplaner.appinstaller`, `installer-dist/msix-package/Package.appxmanifest` und `installer-dist/build-msix.ps1` erhoehen.
2. EXE mit `pyinstaller GAWELA-Tourenplaner.spec --noconfirm` neu bauen.
3. MSIX mit `powershell -ExecutionPolicy Bypass -File .\\installer-dist\\build-msix.ps1` erstellen.
4. GitHub-Release-Assets mit `powershell -ExecutionPolicy Bypass -File .\\installer-dist\\prepare-github-release.ps1 -Version "<VERSION>" -PackageName "GAWELA.Tourenplaner" -Publisher "CN=GAWELA" -MsixPath ".\\installer-dist\\GAWELA-Tourenplaner.msix"` vorbereiten.
5. Tag und GitHub-Release anlegen.
6. Genau diese beiden Dateien hochladen:
   `installer-dist\\release-assets\\GAWELA-Tourenplaner.appinstaller`
   `installer-dist\\release-assets\\GAWELA-Tourenplaner.msix`

## Hinweise

- Das MSIX muss mit derselben Package Identity und demselben Publisher signiert bleiben.
- Fuer Windows-Updates muss die Paketversion bei jedem Release steigen.
- Die Installation sollte immer ueber die `.appinstaller`-Datei getestet werden, nicht nur ueber das nackte `.msix`.
