# crbxparser

Konvertiert SIA 451 / CRBX Leistungsverzeichnisse (`.01S`, `.e1s`) in CSV.

## Verwendung

```bash
python parse_sia451.py <Eingabedatei> [Ausgabedatei.csv]
```

Wird keine Ausgabedatei angegeben, wird der Dateiname der Eingabedatei mit der Endung `.csv` verwendet.

**Beispiel:**
```bash
python parse_sia451.py SIA451.01S
# erzeugt SIA451.csv
```

## Ausgabe

Die CSV-Datei ist UTF-8 (mit BOM) kodiert und verwendet Semikolon als Trennzeichen – Excel öffnet sie direkt.

| Spalte | Beschreibung |
|---|---|
| Kapitel | Datensatz-Typ (z.B. `G361`, `G364`) |
| Abschnitt | Kapitel-Titel aus dem Leistungsverzeichnis |
| Pos-Nr | 6-stellige Positionsnummer |
| Kurztext | Kurzbeschreibung der Position |
| Langtext | Vollständiger Beschrieb (aus allen Langtextzeilen zusammengesetzt) |
| Einheit | Mengeneinheit (z.B. `m`, `m2`, `St`, `h`) |
| Menge | Ausgeschriebene Menge |
| Einheitspreis | Preis pro Einheit in CHF |
| Gesamtpreis | Menge × Einheitspreis in CHF |
| BKP | BKP-Code der Position |

## Format

Das SIA 451 Format (auch CRBX genannt) ist der Schweizer Standard für den elektronischen Datenaustausch von Leistungsverzeichnissen (Norm SN 509 451). Die Datei ist im DOS-Zeichensatz CP850 kodiert und verwendet ein festes Spaltenformat.
