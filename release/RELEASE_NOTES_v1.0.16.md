# Release Notes v1.0.16

## Highlights

- Die EXE-GUI zeigt beim Start der AP1-Konfiguration jetzt einen sichtbaren Fortschrittsbalken.
- Der Laufstatus wird fortlaufend aktualisiert und bleibt während der Ausführung nachvollziehbar.
- Nach erfolgreichem Abschluss wird der Balken grün und zeigt klar den Status `Fertig`.

## Behoben

- Fehlende sichtbare Rückmeldung zum Fortschritt während der Konfiguration.
- Fehlende eindeutige Fertigmeldung nach Abschluss des AP1-Laufs.

## Technische Anpassungen

- Fortschrittsanzeige in `src/main.py` integriert (Statusbalken + Prozent/Fertig-/Fehlerstatus).
- Überwachung des Hintergrundprozesses ergänzt, inkl. Abschlusszustand und Reaktivierung des Startbuttons.
- Fortschrittsableitung über Log-Marker aus `data/4. Logs` umgesetzt.
- README, Anwender- und Technikdokumentation sowie Changelog auf den neuen GUI-Stand aktualisiert.

## Ergebnis

- Build über `src/build.ps1` erfolgreich.
- Release-Artefakte erstellt:
  - `dist/AP1-Konfigurator.exe`
  - `release/AP1-Konfigurator-v1.0.16/`
  - `release/AP1-Konfigurator-v1.0.16.zip`
