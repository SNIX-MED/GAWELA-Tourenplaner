# GitHub Releases fuer MSIX + App Installer

Dieses Repository ist jetzt auf stabile GitHub-Release-Downloads vorbereitet.

## Feste Asset-Namen

Lade bei jedem Release genau diese Dateien hoch:

- `GAWELA-Tourenplaner.appinstaller`
- `GAWELA-Tourenplaner.msix`

Die App verwendet diese stabile URL:

- `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.appinstaller`

Die `.appinstaller`-Datei verweist wiederum auf:

- `https://github.com/SNIX-MED/GAWELA-Tourenplaner/releases/latest/download/GAWELA-Tourenplaner.msix`

## Ein Release vorbereiten

Beispiel:

```powershell
.\installer-dist\prepare-github-release.ps1 `
  -Version "1.0.1.0" `
  -PackageName "DEIN.PACKAGE.NAME" `
  -Publisher "CN=DEIN-PUBLISHER" `
  -MsixPath "C:\Users\Mike\OneDrive\GAWELA-Tourenplaner\GAWELA-Tourenplaner.msix"
```

Ergebnis:

- `installer-dist\release-assets\GAWELA-Tourenplaner.appinstaller`
- `installer-dist\release-assets\GAWELA-Tourenplaner.msix`

## Danach auf GitHub

1. Neues Tag/Release anlegen, z. B. `v1.0.1`
2. Beide Dateien aus `installer-dist\release-assets\` hochladen
3. Sicherstellen, dass die Asset-Namen exakt gleich bleiben
4. Auf einem installierten Client `Jetzt nach Updates suchen` testen

## Wichtig

- Die Version in der MSIX-Paketidentitaet und in der `.appinstaller`-Datei muss steigen.
- `Package Name` und `Publisher` muessen exakt zur MSIX-Identitaet passen.
- Das Paket muss signiert sein.
- Fuer `latest/download/...` sollte pro Release immer nur ein Asset mit exakt diesem Namen existieren.
