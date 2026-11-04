# Abgleich mit KBART Datei
Das Programm vergleicht KBART- mit Excel-Dateien anhand der ISBNs (`print_identifier` und `online_identifier`). Mit Beenden des Abgleichs wird eine neue KBART erstellt, welche die Titel aus der Excel-Datei enthält.

Es soll dazu dienen, aus einer großen Anzahl an Titeln in der KBART nur diejenigen auszuwählen, die in der Excel-Tabelle vorhanden sind und sie im KBART-Format bereitzustellen.

Dafür werden alle Spalten der Excel-Datei ausgewertet, die "ISBN" im Namen tragen.

Sollte kein ISBN-Wert einer Excel-Zeile in der KBART vorhanden sein, wird die Zeile mit ihrer Zeilennummer ausgegeben und kann manuell überprüft werden.

Stornierte Titel sollten vor dem Abgleich aus der Excel-Datei entfernt werden, da das Programm sie nicht erkennt und gefundene Titel sonst in die neue KBART übernimmt.

Anwendungsbeispiele sind z.B. die Auswahl von EBS-Titeln, Pick&Choose Titeln oder Einzelkäufe von Bibliotheken

## Oberfläche

Der gesamte Ablauf findet in einem Fenster statt: beide Dateien auswählen, ISBN-Spalte
bestimmen, Abgleich starten, Ergebnis prüfen, exportieren.

* Alle Spalten mit "ISBN" im Namen werden automatisch ausgewertet, auch mehrere ISBNs
  in einer Zelle (getrennt durch Semikolon). Eine einzelne Spalte bleibt wählbar,
  darunter erscheinen Beispielwerte zur Kontrolle.
* Vor dem Speichern zeigt das Programm, **was** gefunden wurde: Trefferzahl,
  Anzahl der Titel ohne Entsprechung, Trefferquote und beide Listen als Tabelle.
  Fehlende Titel werden mit ihrer Zeilennummer in der Excel-Datei ausgegeben.
* Der Reiter **Hinweise** nennt Zeilen ohne ISBN ("Prüfen: leere ISBN, Ocu oder bestellt?")
  sowie Ocu-Titel, erkannt über die Spalte `BBG`. Entfernt wird dabei nichts.
* Der Abgleich läuft in einem Hintergrund-Thread, das Fenster bleibt bedienbar.
* Der Filter für Zeitschriften (`publication_type = Serial`) ist als Option sichtbar
  und abschaltbar.
* Helles und dunkles Farbschema.

## Aufbau

* `kbart_gui.py`: Oberfläche und Programmstart.
* `kbart_logic.py`: Einlesen, Bereinigen, Abgleichen und Schreiben, ohne Oberflächenbezug.
* `tests/`: Automatische Tests der Auswertung.

## Installation

```
pip install -r requirements.txt
python kbart_gui.py
```

Das Programm basiert auf Python und ist auf verschiedenen Systemen wie Windows und Linux ausführbar.
Um es manuell in eine ausführbare Datei umzuwandeln, kann folgender Befehl genutzt werden:

```
pyinstaller --onefile --windowed --hidden-import openpyxl --collect-all customtkinter kbart_gui.py
```

Der Zusatz `--collect-all customtkinter` ist notwendig, weil PyInstaller die
Theme-Dateien des Pakets sonst nicht mit einpackt und die fertige Datei beim Start abbricht.

## Tests

```
pip install pytest
python -m pytest
```

* `tests/test_kbart_logic.py` prüft Bereinigung, Abgleich, Einlesen und Schreiben mit kleinen
  Beispieldaten und läuft überall.
* `tests/test_real_data.py` gleicht die echten Verlagslisten mit den erwarteten Ergebnissen ab.

Die Tests mit echten Verlagslisten laufen nur, wenn die Dateien im Ordner `testdaten/` liegen. Startbar mit: python.exe -m pytest  
Dieser Ordner wird nicht versioniert, ohne ihn werden diese Tests übersprungen.

### Testdaten

Die echten Listen liegen im UB-Wiki auf der Seite "Test Cases KARMA" und werden nicht
in dieses öffentliche Repository hochgeladen, weil die Kaufdatei interne Bestandsdaten enthält.
Für die Tests gehören sie mit unveränderten Namen nach `testdaten/`:

* `Cambridge_einzeltitel_Stand_2026_09_02.xlsx` (Kaufdatei mit 286 Zeilen)
* `Cambridge_einzeltitel_283.xlsx` (dieselbe Datei ohne die Zeilen 116, 147 und 264)
* `CUP_Cambridge_Core_All_Books_2026-09-02.tsv`
* `CUP_Cambridge_Core_-_Withdrawn_Books_2026-09-02.tsv`
* `CUP_Cambridge_Core__All_Books_including_withdrawn_titles_2026-09-02.tsv`
* `BiblioScout_Global_AllBooks.tsv`
* `kbart_gefiltert_nurlaufendegeprüft.tsv`
* `kbart_gefiltert_nurwithdrawngeprüft.tsv`
* `kbart_gefiltert_laufendundwithdrawn.tsv`
