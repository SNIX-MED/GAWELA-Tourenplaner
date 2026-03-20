# GAWELA-Tourenplaner 2.0 (C# Port)

Dies ist eine C#-Neuumsetzung des bestehenden Python-Projekts mit Fokus auf die gleichen Kernfunktionen:

- JSON-basierte Speicherung (Mitarbeiter, Fahrzeuge, Touren)
- Normalisierung und Validierung der Daten
- Zeitfenster-Validierung und Zeitberechnungen
- Tour-Planungslogik (ETA/ETD, Wartezeiten, Konflikte)
- Routing via OSRM (Fahrzeit + Distanz)
- Desktop-Oberfläche mit ähnlicher Struktur (Start / Mitarbeiter / Fahrzeuge / Touren)

## Build (lokal)

```bash
dotnet build GAWELA-Tourenplaner-2.0.sln
```

## GitHub-Repository erstellen

Da dieses Ausführungsumfeld keinen Zugriff auf dein GitHub-Konto hat, kann ich die Remote-Erstellung nicht direkt durchführen.

Führe lokal aus:

```bash
cd GAWELA-Tourenplaner-2.0
git init
git add .
git commit -m "Initial C# port of GAWELA Tourenplaner"
git branch -M main
git remote add origin git@github.com:<DEIN-USERNAME>/GAWELA-Tourenplaner-2.0.git
git push -u origin main
```

Alternative mit GitHub CLI:

```bash
gh repo create GAWELA-Tourenplaner-2.0 --public --source=. --remote=origin --push
```
