# Nächster Termin (NextMeetingOverlay)

Kleines Windows-Overlay, das den nächsten Outlook-Termin anzeigt – als unauffällige Ampel-Pille am Bildschirmrand, mit Details bei Mausberührung.

*A tiny always-on-top Windows overlay showing your next Outlook meeting as a traffic-light pill (German UI).*

![Klein als Ampel-Pille, aufgeklappt mit Details und das Kontextmenü](docs/screenshot.png)

## Funktionen

- **Ampel-Pille:** grün = mehr als 2 Std frei · gelb = Termin in ≤ 2 Std · rot = in ≤ 15 Min oder läuft (blinkt in der letzten Minute) · „+1“ bei überschneidenden Terminen
- **Details** bei Mausberührung oder Klick: Uhrzeit, Countdown, Ort, parallele Termine, Folgetermin mit freier Zeit, Tagesübersicht
- **„Beitreten“-Knopf** für Teams-, Zoom-, Meet- und Webex-Links – verschwindet, sobald du über das Overlay beigetreten bist (erneut beitreten: Strg+Alt+Umschalt+J oder Rechtsklick → „Erneut beitreten“)
- **Nicht teilnehmen (✕):** blendet einen einzelnen Termin im Overlay aus – das nächste rückt nach. In Outlook bleibt der Termin unverändert, es geht keine Absage raus. Rückgängig über den Hinweis oder Rechtsklick → „Ausgeblendete Termine wieder anzeigen“
- **Filter:** als „frei“ markierte Blocker, Stichwörter (Mittag, Pause, …), ganztägige, abgesagte, private Termine – alles einzeln schaltbar
- **Andocken** oben links / Mitte / rechts, auf jedem Monitor
- **Bei Bildschirmfreigabe unsichtbar:** In Teams-/Zoom-Freigaben und auf Screenshots sehen andere das Overlay nicht – du schon
- **Bei Vollbild/Präsentation** automatisch ausgeblendet
- **Tastenkürzel** (Strg+Alt+Umschalt + O/J/A/K): ein/aus, beitreten, Aufnahme, Klick-durch-Modus
- **Optional: [JohnnyAI](https://github.com/HiSchuminator/JohnnyAI)-Steuerung** – Aufnahme mit dem Termin verknüpft starten/stoppen, Aufnahmestatus in der Pille, bei Terminbeginn anbieten oder automatisch starten. Ist JohnnyAI nicht installiert, blendet das Overlay Menü, Knopf und Tastenkürzel dafür automatisch aus.

## Kalender

Rechtsklick → **Kalender …**: beliebig viele Kalender eintragen, einzeln ein- und ausschalten. Die Termine aller eingeschalteten Kalender werden zusammen angezeigt; steht derselbe Termin in mehreren Kalendern, erscheint er nur einmal. Bei mehr als einem Kalender steht dessen Name beim Termin (z. B. „Arbeit“, „Privat“). Ist ein Kalender nicht erreichbar, laufen die anderen weiter – der Hinweis steht im Tooltip.

| Kalender | Einrichtung |
|---|---|
| Klassisches Outlook | „Klassisches Outlook hinzufügen“ – das Overlay liest den geöffneten Standardkalender direkt (COM) |
| Neues Outlook / Outlook im Web | „Kalender-Link hinzufügen“: Outlook → Einstellungen → Kalender → Freigegebene Kalender → „Kalender veröffentlichen“ → ICS-Link kopieren |
| Google Kalender | „Kalender-Link hinzufügen“: Einstellungen → Kalender → „Geheime Adresse im iCal-Format“ |
| iCloud | „Kalender-Link hinzufügen“: Kalender teilen → „Öffentlicher Kalender“ → Link kopieren |

Kalender-Links werden mit Windows DPAPI verschlüsselt gespeichert und funktionieren nur unter dem Windows-Konto, das sie eingetragen hat. Wer einen Link kennt, kann den Kalender lesen – nicht weitergeben.

## Installation

Unter [Releases](../../releases) eine der beiden Dateien herunterladen und starten – keine Installation nötig:

| Datei | Größe | Wann nehmen? |
|---|---|---|
| `NextMeetingOverlay.exe` | ca. 69 MB | Läuft auf jedem Windows-PC, .NET ist eingebaut. **Im Zweifel diese.** |
| `NextMeetingOverlay-klein.exe` | ca. 1 MB | Wenn die [.NET 8 Desktop Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) schon installiert ist. Fehlt sie, zeigt Windows beim Start einen Download-Link. |

Beide können dasselbe. Einstellungen landen im Ordner `NextMeetingOverlay-Daten` neben der Exe – ein Wechsel zwischen den Varianten ist ohne Einstellungsverlust möglich (Datei in denselben Ordner legen). Autostart: Rechtsklick → „Mit Windows starten“.

## Selbst bauen

Voraussetzung: .NET 8 SDK (Windows).

```bash
dotnet publish -c Release -p:Portable=true -o portable
```

Kleine portable Variante (< 1 MB, braucht die installierte .NET-8-Desktop-Laufzeit):

```bash
dotnet publish -c Release -p:Portable=true -p:SelfContained=false -p:EnableCompressionInSingleFile=false -o portable-klein
```

Ganz ohne `-p:Portable=true` entsteht eine Exe, die ihre Einstellungen unter `%APPDATA%\NextMeetingOverlay` statt neben der Exe ablegt.
`--profile <name>` startet eine getrennte Instanz mit eigenen Einstellungen (zum Testen).
`--demo-screenshot docs/screenshot.png` erzeugt das Bild oben neu – mit erfundenen Beispielterminen, ohne Kalenderzugriff.

## Lizenz

[MIT](LICENSE)
