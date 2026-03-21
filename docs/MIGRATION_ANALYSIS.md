# GAWELA Tourenplaner – Bestandsanalyse für die C#-Migration

## Ziel der Analyse
Diese Analyse dokumentiert die aktuell im Python-/CustomTkinter-Projekt vorhandenen Funktionen, Seiten, Datenflüsse, Validierungen, Import-/Export-Pfade, Einstellungen und Sonderfälle als Grundlage für eine vollständige Neuentwicklung in C#.

## Bestehender Technologiestack
- Desktop-UI: `customtkinter` plus `tkinter`/`ttk`.
- Kartendarstellung: `tkintermapview` mit Google-Tiles.
- GPS-Webportal: native WebView2-Einbettung bzw. Browser-Fallback.
- Persistenz: JSON-Dateien im Projekt-/Benutzerdatenbereich.
- SQL-Import: direkter SQL-Server-Zugriff und datenbankspezifische Auftragsaufbereitung.
- Routing/Distanz: OSRM (`router.project-osrm.org`).
- Geocoding: Nominatim mit lokalem Cache und Fair-Use-Drosselung.
- Packaging/Windows-Installation: PyInstaller + MSIX/AppInstaller.

## Seitenstruktur und Navigation
Die Informationsarchitektur ist klar definiert und muss in der C#-Version in derselben Reihenfolge und mit wiedererkennbaren Hauptelementen erhalten bleiben.

1. Start
2. Kalender
3. Karte
4. GPS
5. Auftragsliste
6. Nicht-Karten-Aufträge
7. Liefertouren
8. Mitarbeiter
9. Fahrzeuge
10. Einstellungen
11. Updates (konfigurierbar sichtbar)

### Navigationsbesonderheiten
- Die Startseite blendet die linke Sidebar zunächst aus und zeigt stattdessen Launcher-Kacheln für alle Bereiche.
- Die Sidebar kann ein- und ausgeklappt werden.
- Es gibt Schnellzugriffe in der Sidebar mit konfigurierbaren Aktionen/Seiten.
- Beim Seitenwechsel werden bestimmte Seiten aktiv `refresh()`t.
- Ein Doppelklick im Kalender springt direkt in die Liefertouren mit vorausgewähltem Datum.

## Seiten und Hauptfunktionen

### 1) Startseite
- Direkter Zugriff auf alle Bereiche über große Buttons.
- Hinweistext zum Update-/Installationsstatus.
- App-Branding/Logo.

### 2) Kalender
- Monatsübersicht geplanter Touren.
- Farbliche Markierung je Datum nach Mitarbeitereinsätzen:
  - keine Einträge
  - genau 1 Einsatz
  - 2 oder mehr Einsätze
- Detailbereich für gewähltes Datum:
  - ausgewähltes Datum
  - Anzahl Touren
  - Anzahl Mitarbeitereinsätze
  - Tourtitel
- Horizontale Vorschau geplanter Touren der nächsten 10 Tage.
- Doppelklick/Activation auf ein Datum öffnet die Liefertouren-Seite mit aktivem Datum.

### 3) Karte
Zentraler Arbeitsbereich der Anwendung.

#### Obere Leiste
- Ortssuche.
- Kartenfilterdialog.
- Routenexport.

#### Kartenlegende
- Farbcode für Auftragsstatus.
- Formcode für Lieferarten.

#### Routenpanel
- Aktuelle Route / aktueller Tourkontext.
- Tour wählen.
- Tour bearbeiten.
- Neue Tour erstellen.
- Tour verlassen.
- Startzeit setzen/übernehmen.
- Route optimieren.
- Stoppliste mit Reihenfolge, Adresse, Zeitfenster, Aufenthalt, ETA/ETD, Gewicht.
- Drag & Drop bzw. Hoch/Runter zum Umordnen.
- Zeitfensterbearbeitung je Stopp.
- Stopp entfernen.
- Summen-/Zeitplanbereich.
- Segment-/Fahrzeitenliste.

#### Karten-/Detailbereich
- Markeranzeige für geocodierte Aufträge.
- Detailkarte für selektierten Auftrag.
- Auftrag zu Route hinzufügen.
- E-Mail an Auftrag senden.
- Auftrag bearbeiten.
- Auftragsstatus ändern.
- Zugehörige Tour anzeigen.

### 4) GPS
- Eingebettete native WebView2-Ansicht für `https://map.ktrac.ch/`.
- Neu laden.
- Im Browser öffnen.
- URL kopieren.
- Laufzeit-/Runtime-Hinweis zur gebündelten oder System-WebView2.
- Robuste Fallbacks bei fehlender WebView2-Unterstützung.

### 5) Auftragsliste
- Tabellarische Übersicht kartierbarer Aufträge.
- SQL-Importstatus.
- Suche/Filter/Interaktion mit Aufträgen.
- Hinweise für importierte, aber noch nicht geocodierte Aufträge.
- Aktionen Richtung Karte/Tour.

### 6) Nicht-Karten-Aufträge
- Separater Bereich für Aufträge, die aufgrund Lieferart nicht auf die Karte gehören.
- Gruppierung/Kategorisierung über `NichtKarteKategorie` bzw. Lieferart.
- Separate Liste abseits der Kartenlogik.

### 7) Liefertouren
- Liste gespeicherter Touren.
- Filter nach Datum bzw. Datumsbereich.
- Tour-Attribute in Tabelle:
  - ID
  - Datum
  - Name
  - Fahrzeug
  - Mitarbeiter
  - Stopps
  - Totalgewicht
- Aktionen:
  - Tour auf Karte anzeigen
  - Bearbeiten
  - Löschen
  - Neue Tour aus aktueller Route speichern
- Ressourcenkonfliktlogik für Fahrzeug/Anhänger an demselben Datum.

### 8) Mitarbeiter
- CRUD für Mitarbeiter.
- Attribute:
  - ID
  - Name
  - Kürzel
  - Telefon
  - Aktiv/Inaktiv
  - Erstellungszeitpunkt
- Auswahldialog für Touren mit maximal 2 Mitarbeitern.
- Aktive Mitarbeiter werden für Tourzuordnung verwendet.

### 9) Fahrzeuge
- Zwei Segmente/Tabs:
  - Zugfahrzeuge
  - Anhänger
- CRUD für beide Typen.
- Attribute u. a.:
  - Name
  - Kennzeichen
  - Nutzlast
  - Anhängelast (nur Zugfahrzeug)
  - Volumen
  - Ladeflächenmaße
  - Notizen
  - Aktiv/Inaktiv
  - Timestamps

### 10) Einstellungen
- Erscheinungsbild/System/Light/Dark.
- Schnellzugriffe konfigurieren.
- SQL-Datenordner.
- SQL-Server-Instanz.
- SQL-Datenbank.
- Backup-Funktionen:
  - Backups aktivieren
  - Auto-Backup aktivieren
  - Backup-Ordner
  - Standardmodus Voll/Inkrementell
  - Aufbewahrungstage
  - Auto-Backup-Intervall
  - Backup jetzt ausführen
  - Wiederherstellen mit Datengruppen
- Gruppenweise Wiederherstellung:
  - Aufträge & Adressen
  - Liefertouren
  - Mitarbeiter
  - Fahrzeuge
  - Einstellungen
  - Zusatzdaten
  - Weitere Daten

### 11) Updates
- Anzeige von Internet-/Installations-/Versionsstatus.
- Anzeige/Kopieren der AppInstaller-URL.
- Updates sofort suchen/installieren.
- Auto-Update-Einstellungen anzeigen.
- Hilfe öffnen.
- Option „Beim Start nach Updates suchen“.
- Reparatur-/Fallback-Logik für `ms-appinstaller`.

## Datenmodelle / Persistenz

### `pins.json`
Liste geocodierter bzw. kartierter Aufträge.
Wichtige Felder:
- `lat`, `lng`
- `status`
- `data` mit u. a.:
  - `ImportID`
  - `Auftragsnummer`
  - `Bestelldatum`
  - `Name`
  - `Strasse`
  - `PLZ`
  - `Ort`
  - `Land`
  - `Email`
  - `Telefon`
  - `Gewicht`
  - `Notizen`
  - `Lieferart`
  - `Status`

### `tours.json`
Liste gespeicherter Touren.
Wichtige Felder:
- `id`
- `date`
- `name`
- `employee_ids`
- `vehicle_id`
- `trailer_id`
- `start_time`
- `route_mode`
- `travel_time_cache`
- `stops[]` mit u. a.:
  - `id`
  - `name`
  - `address`
  - `lat`, `lon`, `lng`
  - `order`
  - `auftragsnummer`
  - `weight`
  - `time_window_start`, `time_window_end`
  - `service_minutes`
  - `planned_arrival`, `planned_departure`
  - `schedule_conflict`, `schedule_conflict_text`
  - `wait_minutes`

### `data/employees.json`
Liste von Mitarbeitern.

### `data/vehicles.json`
Objekt mit:
- `vehicles[]`
- `trailers[]`

### `settings.json`
Anwendungseinstellungen.
Wichtige Felder:
- `appearance_mode`
- `quick_access_items`
- `sql_data_dir`
- `sql_server_instance`
- `sql_database`
- `backups_enabled`
- `backup_dir`
- `backup_mode_default`
- `backup_retention_days`
- `auto_backup_enabled`
- `auto_backup_interval_days`
- `last_backup_iso`

### Weitere Dateien
- `config.json`: ältere/zusätzliche Konfigurationswerte.
- `geocode_cache.json`: Geocoding-Cache.
- diverse Backup-Artefakte.
- Windows-Release-/Installer-Dateien unter `installer-dist/`.

## Datenflüsse

### SQL-Import
1. SQL-Verbindung über konfigurierten Server/DB.
2. Abruf offener Aufträge.
3. Deduplikation über Identität aus `ImportID`/`Auftragsnummer`.
4. Trennung in:
   - kartierbare Aufträge
   - Nicht-Karten-Aufträge nach Lieferart
5. Übernahme vorhandener Koordinaten/Status, falls Auftrag schon lokal bekannt ist.
6. Nicht geocodierte Aufträge werden als Pending-Liste gehalten.
7. Hintergrund-Geocoding mit Statusfortschritt.
8. Erfolgreiche Geocodes werden zu Markern/Pins.
9. Nicht mehr offene Aufträge werden entfernt.
10. Es wird ein Fehlerreport für nicht geocodierbare Adressen geschrieben.

### Karten-/Routenfluss
1. Auftrag selektieren.
2. Zum aktuellen Routenpanel hinzufügen.
3. Reihenfolge manuell oder per Optimierung bestimmen.
4. Fahrzeiten pro Segment laden/cachen.
5. Zeitplan mit ETA/ETD aus Startzeit, Zeitfenstern, Servicezeiten berechnen.
6. Tour mit Ressourcen speichern.
7. Tour später wieder laden und auf Karte darstellen.

### Geocoding
1. Adresskandidaten aus Auftragsdaten generieren.
2. Nominatim-Anfrage mit Fair-Use-Verzögerung.
3. Cache-Hit bevorzugen.
4. Fehler/Timeouts mit Retry behandeln.
5. Ergebnisse in `geocode_cache.json` atomar speichern.

### Updates
1. Laufzeitkontext ermitteln (MSIX/portable etc.).
2. Erreichbarkeit der Updatequelle prüfen.
3. `ms-appinstaller` bevorzugen.
4. Fallback: `.appinstaller` direkt öffnen.
5. Fallback: lokale temporäre Datei / Browser.

### Backups
1. Snapshot relevanter Daten-/Konfigurationsdateien.
2. Voll- oder inkrementelles ZIP-basiertes `.bak`.
3. Manifest mit Dateihashes/Metadaten.
4. Optionale Wiederherstellung nach Datengruppen.
5. Bereinigung alter Backups nach Retention.

## Validierungen und Geschäftsregeln

### Touren
- Datum muss parsebar sein.
- Startzeit muss im Format `HH:MM` gültig sein.
- 1 oder 2 Mitarbeiter müssen ausgewählt werden.
- Fahrzeug ist Pflicht.
- Anhänger optional.
- Ressourcenkonflikte (gleiches Fahrzeug/Anhänger am selben Datum) werden erkannt und bestätigt.
- Vergangenheit wird als Warnung markiert.
- Tour mit zu wenigen Stopps wird nicht sinnvoll angezeigt.

### Mitarbeiter
- Maximal 2 Mitarbeiter pro Tour.
- Nur aktive Mitarbeiter für operative Auswahl.

### Fahrzeuge/Anhänger
- Numerische Felder wie Nutzlast/Volumen/Ladefläche.
- Aktiv/Inaktiv-Status.

### Einstellungen
- Appearance nur `System`, `Light`, `Dark`.
- Backup-Modus nur `full`, `incremental`.
- Retention und Auto-Backup-Intervall: ganzzahlig zwischen 1 und 365.
- Backup-Ordner muss schreibbar sein.
- Legacy-Migration von `xml_folder` zu `sql_data_dir`.
- Doppelte Schnellzugriffe werden bereinigt.

### JSON-Speicherung
- Atomisches Schreiben.
- Ungültige JSON-Dateien können gesichert/ersetzt werden.
- Strukturprüfungen pro Datei.

## Sonderfälle / Edge Cases
- Fehlerhafte JSON-Dateien werden als korrupt gesichert.
- Nicht-Windows oder fehlende WebView2-Unterstützung im GPS-Bereich.
- `ms-appinstaller` kann per Richtlinie deaktiviert sein.
- Netzwerk-/DNS-/HTTP-Probleme bei Updateprüfung.
- OSRM- oder Geocoding-Ausfälle.
- SQL-Datenbankname kann aus Datenordner inferiert werden.
- Doppelte SQL-Zeilen werden gezählt und ignoriert.
- Aufträge ohne Koordinaten bleiben als Pending erhalten.
- Marker ohne valide Identität werden bereinigt.
- Fehlende Stopps beim Laden einer Tour lösen Warnungen aus.

## Externe Abhängigkeiten / Integrationen für die C#-Version
- Microsoft SQL Server / SQL Server Express.
- OSRM HTTP API.
- Nominatim HTTP API oder kompatibler Geocoder.
- WebView2 Runtime.
- Windows-Installer-Technologie (MSIX, MSI oder Setup-Bootstrapper).

## Migrationsrelevante Schlussfolgerungen
- Die Anwendung ist funktional deutlich mehr als eine reine Karten-UI: SQL-Import, Geocoding, Tourplanung, Ressourcenzuordnung, Backup/Restore und Updates müssen vollständig übernommen werden.
- Die Python-Version nutzt dateibasierte Persistenz und viele serviceartige Hilfsfunktionen; diese eignen sich gut für eine Aufteilung in C#-Domäne, Application Layer, Infrastructure und WPF-UI.
- Die Seitenreihenfolge, der Start-Launcher, die Sidebar, das Routenpanel und die Tabellen-/Dialogabläufe sind zentrale Wiedererkennungsmerkmale und sollten im C#-UI erhalten bleiben.
- Für Windows 10/11 und WebView2 ist eine .NET-8-WPF-Anwendung mit WebView2-Integration fachlich naheliegend.

## Offene Punkte vor der vollständigen Migration
1. Online-Erstellung eines neuen GitHub-Repositories ist ohne GitHub-Zugangsdaten/Remote-Berechtigung aus dieser Umgebung nicht möglich.
2. Für eine vollständige Build-/Installer-Validierung fehlt in der Umgebung aktuell das .NET-SDK.
3. Für 100% Funktionsgleichheit sollten reale SQL-Beispieldaten und typische Backup-/Restore-Dateien testweise bereitgestellt werden.
