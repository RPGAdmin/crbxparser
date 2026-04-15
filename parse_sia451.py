#!/usr/bin/env python3
"""
SIA 451 / CRBX Parser
Konvertiert SIA 451 Leistungsverzeichnisse in CSV.

Spalten-Layout (0-basiert):
  [0:4]   Datensatz-Typ (z.B. "G361", "G364")
  [5:7]   Kapitel-Nr bei Level-1-Datensätzen (z.B. "24")
  [7:13]  Positions-Basisnummer (6 Ziffern, z.B. "176223")
  [14:20] GP-Marker bei Preis-Datensätzen
  [41]    Satz-Art (1=Kapitel, 2=Kurztext, 3=Langtext, 5=Einheit, 6=Preis+Menge)
  [42:]   Text / Daten (je nach Satz-Art)

Level-6-GP-Layout:
  [43]    Mengen-Typ (A=Angebot, V=variabel, ...)
  [44:58] Menge (+sign + 13 Ziffern, 3 Dezimalstellen impliziert)
  [61:74] Einheitspreis (+sign + 12 Ziffern, 2 Dezimalstellen impliziert)
  [81:85] BKP-Code
"""

import csv
import sys
from pathlib import Path


def parse_qty(s: str) -> float | None:
    """Menge: +sign + 13 Ziffern, 3 Dezimalstellen impliziert."""
    s = s.strip()
    if not s or len(s) < 2:
        return None
    try:
        sign = -1 if s[0] == '-' else 1
        return sign * int(s[1:]) / 1000
    except ValueError:
        return None


def parse_price(s: str) -> float | None:
    """Preis: +sign + 12 Ziffern, 2 Dezimalstellen impliziert."""
    s = s.strip()
    if not s or len(s) < 2:
        return None
    try:
        sign = -1 if s[0] == '-' else 1
        return sign * int(s[1:]) / 100
    except ValueError:
        return None


def clean_text(s: str) -> str:
    return s.strip()


def parse_sia451(filepath: str) -> list[dict]:
    """Parst eine SIA 451 Datei und gibt eine Liste von Positionen zurück."""

    positions = []
    current: dict | None = None
    current_section = ""

    with open(filepath, "r", encoding="cp850") as f:
        for raw_line in f:
            line = raw_line.rstrip("\n\r")
            if len(line) < 42:
                continue

            rec_type = line[0:4].rstrip()

            # Nur G-Datensätze (Positionen/Kapitel) verarbeiten
            if not rec_type.startswith("G"):
                continue

            kapitel_code = rec_type  # z.B. "G361", "G364"

            # Satz-Art: Position 41 (0-basiert)
            level_char = line[41]
            if not level_char.isdigit():
                continue
            level = int(level_char)

            # GP-Marker: Zeichen 14-20
            gp_flag = "GP" in line[14:20].upper()

            # Text-Feld (ab Zeichen 42, gestrippt)
            text = clean_text(line[42:]) if len(line) > 42 else ""

            # ------------------------------------------------------------------
            # Level 1: Kapitel-Header (Abschnittstitel)
            # ------------------------------------------------------------------
            if level == 1:
                # Kapitel-Nummer steht in Chars 5-7
                section_nr = line[5:7].strip()
                section_text = " ".join(line[42:].split())  # merge continuation text
                current_section = f"{section_nr} {section_text}".strip()
                continue

            # ------------------------------------------------------------------
            # Positions-Basisnummer: Zeichen 7-13 (6 Ziffern)
            # ------------------------------------------------------------------
            pos_base = line[7:13].strip() if len(line) >= 13 else ""
            if not pos_base:
                continue  # kein gültiger Positions-Datensatz

            # ------------------------------------------------------------------
            # Level 2: Neue Position / Kurztext
            # ------------------------------------------------------------------
            if level == 2:
                if current is not None:
                    positions.append(current)
                current = {
                    "Kapitel":        kapitel_code,
                    "Abschnitt":      current_section,
                    "Pos-Nr":         pos_base,
                    "Kurztext":       text,
                    "Langtext":       "",
                    "Einheit":        "",
                    "Menge":          None,
                    "Einheitspreis":  None,
                    "Gesamtpreis":    None,
                    "BKP":            "",
                }
                continue

            if current is None:
                continue

            # ------------------------------------------------------------------
            # Level 3: Langtext-Zeile
            # ------------------------------------------------------------------
            if level == 3 and text:
                sep = " " if current["Langtext"] else ""
                current["Langtext"] += sep + text
                continue

            # ------------------------------------------------------------------
            # Level 5: Einheit
            # ------------------------------------------------------------------
            if level == 5:
                current["Einheit"] = text
                continue

            # ------------------------------------------------------------------
            # Level 6 + GP: Menge, Einheitspreis, BKP
            # ------------------------------------------------------------------
            if level == 6 and gp_flag and len(line) >= 85:
                current["Menge"] = parse_qty(line[44:58])
                current["Einheitspreis"] = parse_price(line[61:74])
                current["BKP"] = line[81:85].strip()

                menge = current["Menge"]
                ep = current["Einheitspreis"]
                if menge is not None and ep is not None:
                    current["Gesamtpreis"] = round(menge * ep, 2)

    # Letzte Position nicht vergessen
    if current is not None:
        positions.append(current)

    return positions


def write_csv(positions: list[dict], outfile: str) -> None:
    if not positions:
        print("Keine Positionen gefunden.", file=sys.stderr)
        return

    fieldnames = [
        "Kapitel", "Abschnitt", "Pos-Nr", "Kurztext", "Langtext",
        "Einheit", "Menge", "Einheitspreis", "Gesamtpreis", "BKP",
    ]

    with open(outfile, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=";")
        writer.writeheader()
        writer.writerows(positions)

    print(f"{len(positions)} Positionen -> {outfile}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Verwendung: python parse_sia451.py <Eingabedatei> [Ausgabedatei.csv]")
        sys.exit(1)

    infile = sys.argv[1]
    outfile = sys.argv[2] if len(sys.argv) >= 3 else Path(infile).stem + ".csv"

    positions = parse_sia451(infile)
    write_csv(positions, outfile)
