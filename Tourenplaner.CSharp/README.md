# Tourenplaner C#

C#-Neuentwicklung des bestehenden GAWELA-Tourenplaners mit Fokus auf vollständige Funktionsübernahme, gleicher Seitenstruktur und Windows-Deployment.

## Zielsetzung
- **100% Funktionsgleichheit** zur bestehenden Python-Version.
- **Gleiche Informationsarchitektur**: Start, Kalender, Karte, GPS, Auftragsliste, Nicht-Karten-Aufträge, Liefertouren, Mitarbeiter, Fahrzeuge, Einstellungen, Updates.
- **Modernisierte UI** auf Basis von WPF, ohne die wiedererkennbaren Bedienabläufe zu verlieren.
- **Windows 10/11 Deployment** per lauffähigem Build und vorbereitetem Installer-/Packaging-Prozess.

## Aktueller Stand
- Solution- und Projektgerüst für `.NET 8` vorbereitet.
- Projekte für `App`, `Application`, `Domain`, `Infrastructure` und `Tests` angelegt.
- Erste Kernlogik aus dem Python-Projekt nach C# übertragen:
  - Zeit-/Schedule-Berechnung
  - heuristische Routenoptimierung
  - Settings-Validierung
  - JSON-Dateispeicherung
- WPF-Shell mit identischer Hauptnavigation als erstes UI-Grundgerüst angelegt.
- Start-/Bereichsansichten lesen bereits Live-Zahlen aus den vorhandenen JSON-Dateien des Python-Projekts.
- Für Mitarbeiter, Fahrzeuge und Touren existieren bereits erste echte datengetriebene Tabellenansichten.
- Mitarbeiter-, Fahrzeug-/Anhänger- und Tourenbereiche unterstützen bereits grundlegende CRUD-Aktionen mit JSON-Speicherung und Snapshot-Reload.
- Das Karten-/Routenpanel ist als datengetriebene Vorstufe mit Tour-/Ressourcenbasis weiter angenähert.
- Build-/Publish-Skripte für Windows ergänzt.

## Projektstruktur
- `Tourenplaner.CSharp.sln`
- `src/Tourenplaner.CSharp.App` – WPF-Shell und spätere Seiten/Views
- `src/Tourenplaner.CSharp.Application` – Abstraktionen und Orchestrierung
- `src/Tourenplaner.CSharp.Domain` – Fachmodelle und Kernregeln
- `src/Tourenplaner.CSharp.Infrastructure` – JSON, Logging, Repositories
- `tests/Tourenplaner.CSharp.Tests` – Kernfunktionstests
- `build/` – Build-/Publish-/Packaging-Skripte

## Voraussetzungen
- .NET SDK 8.0
- Windows 10/11 für WPF-Build und Ausführung
- Für späteres MSIX-Packaging: Windows Packaging Tooling / Visual Studio mit MSIX-Unterstützung

## Lokaler Build
```powershell
pwsh .\build\build.ps1
```

```bash
./build/build.sh
```

## Tests
```bash
./build/test.sh
```

## Release-Build / Publish
```powershell
pwsh .\build\publish.ps1 -Configuration Release -Runtime win-x64
```

```bash
./build/publish.sh Release win-x64 ./artifacts/publish
```

## Installer / installierbare Version
Aktuell ist ein vorbereiteter MSIX-Packaging-Schritt dokumentiert. Für die endgültige installierbare Version wird im nächsten Schritt ein dediziertes Packaging-Projekt ergänzt.

```powershell
pwsh .\build\package-msix.ps1
```

## Nächste Umsetzungsschritte
1. Karten-/Tourenpanel mit echter Stopp-Reihenfolge, Zeitfenstern und Segmentzeiten konkret umsetzen.
2. SQL-Import, Geocoding, Routing und Backup/Restore migrieren.
3. WebView2-GPS-Seite umsetzen.
4. Kalender- und Update-Seiten fachlich ausbauen.
5. Windows-Installer (MSIX/AppInstaller oder alternatives Setup) fertigstellen.


## Verbleibende Risiken
- In dieser Linux-Umgebung konnte das `.NET SDK` wegen externer Netzwerkbeschränkungen nicht nachinstalliert werden.
- Der WPF-Build ist auf Windows ausgerichtet; mit `EnableWindowsTargeting=true` ist die Solution nun für Cross-Target-Restore/Build besser vorbereitet, muss aber auf einer Windows-Maschine mit installiertem SDK final validiert werden.
- Als nächster fachlicher Schritt fehlen weiterhin die echten Seitenimplementierungen, SQL-Import, Geocoding, Routing, Backup/Restore und WebView2-GPS.
