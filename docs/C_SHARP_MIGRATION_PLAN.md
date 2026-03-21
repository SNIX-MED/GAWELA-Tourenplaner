# C#-Migrationsplan – GAWELA Tourenplaner

## Zielarchitektur
Empfohlene Zielplattform: **.NET 8 + WPF + WebView2** auf Windows 10/11.

### Projektaufbau
- `Tourenplaner.CSharp.App` – WPF-UI
- `Tourenplaner.CSharp.Application` – Use-Cases, Commands, Orchestrierung
- `Tourenplaner.CSharp.Domain` – Entities, Value Objects, Regeln
- `Tourenplaner.CSharp.Infrastructure` – JSON, SQL, HTTP, Geocoding, Routing, Backups, Logging
- `Tourenplaner.CSharp.Tests` – Kernlogiktests
- `build/` – Build-, Publish-, Installer-Skripte

## Alt -> Neu Feature-Mapping

| Alt (Python) | Neu (C#) | Hinweise |
|---|---|---|
| CustomTkinter Hauptfenster + Sidebar | WPF Shell mit identischer Seitenreihenfolge | Wiedererkennbare Navigation beibehalten |
| Startseite mit Launcher-Kacheln | WPF Dashboard Page | Sidebar im Startbereich ausgeblendet |
| Kalenderseite + Tourübersicht | WPF Calendar Page | Doppelklick öffnet Touren-Ansicht |
| `tkintermapview` | WPF-Kartenkomponente (WebView2 + Leaflet/OpenStreetMap oder GMap.NET) | Bedienlogik und Markerstatus erhalten |
| GPS-Webseite via WebView2 | WPF WebView2 Page | Browser-Fallback beibehalten |
| Auftragsliste | WPF DataGrid Page | Such-/Importstatus und Kartenaktionen erhalten |
| Nicht-Karten-Aufträge | WPF DataGrid Page | Separate Informationsarchitektur beibehalten |
| Liefertourenliste | WPF DataGrid + Dialoge | Tabelle, Filter, Aktionen identisch |
| Mitarbeiterverwaltung | WPF CRUD Page + Dialoge | Max. 2 Mitarbeiter pro Tour |
| Fahrzeugverwaltung | WPF CRUD Page mit Segmenten | Fahrzeuge/Anhänger getrennt |
| Einstellungen + Backup/Restore | WPF Settings Page | Datengruppen-Wiederherstellung erhalten |
| Update-Seite für AppInstaller/MSIX | WPF Update Page | Windows-spezifische Updateintegration |
| JSON-Dateien | JSON-Repositories in Infrastructure | Bestehende Formate möglichst kompatibel |
| SQL-Import-Service | ADO.NET / Microsoft.Data.SqlClient | Bestehende Importlogik übernehmen |
| Nominatim-Geocoder + Cache | HTTP-Geocoder-Service + JSON-Cache | Retry/Rate-Limit übernehmen |
| OSRM-Routing | HTTP Routing-Service | Segmentzeiten + Streckenpfad übernehmen |
| BackupManager | Zip-basierter Backup-Service | Manifest + inkrementelle Sicherung |
| Logging via Python | Serilog / Microsoft.Extensions.Logging | Datei- und Fehlerlogging ergänzen |

## Umsetzungsphasen

### Phase 1 – Bestandsaufnahme und fachliche Spezifikation
- Vollständige Funktionsinventur.
- Datenformate dokumentieren.
- Sonderfälle, Validierungen, Abhängigkeiten katalogisieren.
- Abnahme-Matrix definieren.

### Phase 2 – Lösungsgerüst in C#
- Solution und Projekte anlegen.
- Grundlegende Architektur definieren.
- Konfigurations- und Logging-Basis einrichten.
- Gemeinsame Modelle und Interfaces erstellen.

### Phase 3 – Persistenz und Fachlogik
- JSON-Repositories.
- Settings-/Backup-/Restore-Services.
- Tour-, Mitarbeiter-, Fahrzeug- und Auftragsmodelle.
- Routing-, Geocoding- und SQL-Import-Abstraktionen.

### Phase 4 – UI-Migration Seite für Seite
1. Shell + Navigation
2. Start
3. Kalender
4. Karte + Routenpanel
5. Auftragsliste
6. Nicht-Karten-Aufträge
7. Liefertouren
8. Mitarbeiter
9. Fahrzeuge
10. Einstellungen
11. Updates
12. GPS

### Phase 5 – Qualitätssicherung
- Unit-Tests für Zeitplanung, Tourvalidierung, Konflikterkennung, Settingsvalidierung, Backup-Manifest.
- Integrationsnahe Tests für JSON-Repositories.
- Manuelle UI-Regression gegen diese Analyse.

### Phase 6 – Windows Build, Publish, Installer
- `dotnet publish` für Windows x64.
- Installer-Erzeugung (empfohlen: MSIX oder WiX/Velopack, abhängig von gewünschtem Updatekanal).
- Release-Skripte und Dokumentation.

## Priorisierte technische Entscheidungen
1. **UI-Framework:** WPF.
2. **Pattern:** MVVM.
3. **Logging:** `Microsoft.Extensions.Logging` + Serilog File Sink.
4. **Datenzugriff:** `System.Text.Json` und `Microsoft.Data.SqlClient`.
5. **Web/GPS:** `Microsoft.Web.WebView2`.
6. **Installer:** bevorzugt MSIX für Windows-nahe Update-Story; alternativ MSI/Setup falls Signaturprozess unklar bleibt.

## Risiken
- Kartenmigration: `tkintermapview` hat kein 1:1-WPF-Gegenstück; die UX muss über eine passende Kartenkomponente nachgebildet werden.
- Update-Mechanik: Python-App nutzt AppInstaller/MSIX-spezifische Hilfen; in C# muss die Windows-Installationsstrategie final festgelegt werden.
- SQL-Import: produktive DB-Struktur und Rechte müssen in C# gegengeprüft werden.
- Geocoding-/Routing-Dienste: externe Limits und Erreichbarkeit beeinflussen Verhalten.
- 100% Funktionsgleichheit erfordert Vergleich mit realen Daten und Anwenderabläufen.

## Nächste konkrete Schritte
1. C#-Projektgerüst erzeugen.
2. Domänenmodelle aus den JSON-Strukturen ableiten.
3. JSON-kompatible Repositories implementieren.
4. Shell, Navigation und Startseite in WPF nachbauen.
5. Danach die fachlich komplexe Karten-/Routenlogik migrieren.

## Lokale/organisatorische Blocker
- Das gewünschte neue GitHub-Repository **"Tourenplaner C#"** kann von hier nur vorbereitet, aber nicht online in GitHub angelegt werden.
- In dieser Umgebung ist aktuell kein `.NET SDK` installiert; deshalb kann ich die C#-Lösung zwar vorbereiten, aber nicht lokal kompilieren oder publishen, bis das SDK verfügbar ist.
