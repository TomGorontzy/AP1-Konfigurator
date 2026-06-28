# DOKUMENTATION CHECKLISTE

## Vor dem Lauf

- [ ] `AP1-TN.xlsx` ist aktuell und im Ordner `data/1. Anpassen` vorhanden.
- [ ] Word und Excel sind geschlossen.
- [ ] Benötigte Vorlagen liegen unter `data/2. Bei Bedarf anpassen` bereit.
- [ ] Internetzugang für Nuera-Dateien ist verfügbar.
- [ ] Entscheidung zu Proxy `On`, `Off` oder `Skip` ist getroffen.

## Während des Laufs

- [ ] Skript über `src/AP1-Konfigurator.bat` oder `src/AP1-Konfigurator.ps1` gestartet.
- [ ] eventuelle Office-First-Run-Hinweise bestätigt.
- [ ] keine sichtbare Fehlermeldung im PowerShell-Fenster übersehen.

## Nach dem Lauf

- [ ] Nuera-Ordner liegt auf dem Desktop.
- [ ] Kandidatenordner liegt auf dem Desktop.
- [ ] Logdatei wurde unter `data/4. Logs` erstellt.
- [ ] Word-/Excel-Vorlagen wurden korrekt übernommen.
- [ ] Taskleisten-/Proxy-Einstellungen entsprechen dem gewünschten Zustand.

## Bei Problemen

- [ ] aktuelle Logdatei aus `data/4. Logs` prüfen.
- [ ] Office einmal manuell starten und erneut testen.
- [ ] bei Excel-Problemen optional CSV-Fallback verwenden.
- [ ] Pfade und Schreibrechte im Benutzerprofil kontrollieren.
