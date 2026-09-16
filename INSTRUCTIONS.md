# HeadWaters Lagerabrechnung — Anleitung

Dieses Repo erzeugt pro Lager (Kostenstelle) ein **PDF**: Übersichtsseiten, Kontoauszüge der betroffenen Erfolgskonten, Belege.  
Hauptskript: `generate_abrechnung.py` · Ausgabe: `output/Abrechnung_<Lagername>.pdf`

Python **nur mit [uv](https://docs.astral.sh/uv/)** — kein `pip`, kein manuelles `venv`.

---

## Kurzablauf (neues Lager)

1. **Export aus der Buchhaltung** in `data/` legen (siehe unten).
2. **`lagerinfo.json`** ergänzen oder interaktiv ausfüllen (Steckbrief).
3. Falls neues Lager-Kürzel: **Icon + Zuordnung** in `generate_abrechnung.py` (siehe «Neues Lager verdrahten»).
4. Starten:
   ```bash
   uv sync
   uv run generate_abrechnung.py
   ```
   Oder ohne Nachfragen (wenn Steckbrief in `lagerinfo.json` stimmt):
   ```bash
   uv run generate_abrechnung.py --ja --kostenstelle C-XYZ
   ```
5. PDF in `output/` prüfen — besonders **Fazit**, **Zahlungsseite**, **Seite pro Person** (Essen/Übernachtung).

---

## Was ihr braucht (Daten)

### Ordner `data/` (Standard)

| Pfad | Inhalt |
|------|--------|
| `data/xlsx/` | Excel-Export **mit Kostenstelle**: `Buchungsjournal.xlsx` + Unterordner mit Kontoauszügen (`Aufwand/`, `Ertrag/`, …), je Konto eine `.xlsx` |
| `data/pdf/` | Dieselben Kontoauszüge als **PDF** (gleiche Dateinamen-Logik, Kontonummer im Namen) |
| `data/keinekostenstelle/` | Optional: `Buchungsjournal.xlsx` **ohne** KS-Filter — für TWINT/SumUp/UBS-Eingänge, die noch keine Kostenstelle haben |
| `data/Belege/` | Belegdateien, benannt **`Beleg <Buchungs-ID> …`** (z. B. `Beleg 0423 Rechnung.pdf`) |

**Wichtig:** Im Journal-Export liegt pro Lauf meist **nur eine** Kostenstelle (alle Zeilen mit derselben KS). Vor dem nächsten Lager die `data/xlsx`- und `data/pdf`-Inhalte durch den neuen Export ersetzen.

**Fallback (alt):** Ordner wie `Kontoauszüge 2026`, `…2026-2`, `…2026-3` — wird automatisch erkannt, wenn `data/xlsx` fehlt.

### Was nicht ins Git gehört

Laut `.gitignore`: `output/`, Inhalte von `data/xlsx`, `data/pdf`, `data/keinekostenstelle`, `data/Belege` (nur `.gitkeep`). Exporte und PDFs lokal halten.

### Assets (im Repo)

| Datei | Zweck |
|-------|--------|
| `assets/schriftzug_untertitel.jpg` | Kopfzeile |
| `assets/<kürzel>_icon.png` | Lager-Icon links neben Titel (z. B. `btp_icon.png`, `fbc_icon.png`) |

Fehlt ein Icon für die Kostenstelle, bricht `--ja` mit Fehlermeldung ab.

---

## Lagersteckbrief (`lagerinfo.json`)

Array von Lagern; beim Start wird der passende Eintrag zur **Kostenstelle** geladen (oder interaktiv neu erfasst).

```json
{
  "name": "Family Bible Camp 2026",
  "kostenstelle": "C-FBC",
  "teilnehmer": 82,
  "kinder": 66,
  "kind_faktor": 0.5,
  "anreise": "2026-08-01",
  "abreise": "2026-08-07",
  "verpflegungstage": 6.0,
  "ort": "Grischalodge, Parpan",
  "kueche_vom_haus": false,
  "team_beitrag_max": 400
}
```

| Feld | Bedeutung |
|------|-----------|
| `teilnehmer` | Erwachsene/Jugendliche **ohne** Kinder |
| `kinder` | Anzahl Kinder |
| `kind_faktor` | Kinder zählen als X Personen (typisch `0.5`) |
| `verpflegungstage` | Für Essenskosten pro Person (kann ≠ Nächte sein) |
| `kueche_vom_haus` | `true`: Verpflegung in Miete → Annahme **CHF 12/Person/Tag** für Übernachtungskosten-Bereinigung |
| `team_beitrag_max` | Optional (z. B. FBC): Netto je Paycode auf Konto **3500** ≤ dieser Betrag = **Team**, darüber = **Familien**; `null` = keine Aufteilung auf der Zahlungsseite |

Interaktiver Modus speichert bestätigte Werte zurück in `lagerinfo.json`.

### CLI (alles ohne Nachfragen)

```bash
uv run generate_abrechnung.py --ja \
  --kostenstelle C-FBC \
  --lagername "Family Bible Camp 2026" \
  --ort "Grischalodge, Parpan" \
  --anreise 26.07.2026 \
  --abreise 01.08.2026 \
  --teilnehmer 82 \
  --kinder 66 \
  --kind-faktor 0.5 \
  --verpflegungstage 5.7 \
  --kueche-vom-haus
```

`team_beitrag_max` kommt aus `lagerinfo.json` (noch kein CLI-Flag).

---

## Neues Lager verdrahten (`generate_abrechnung.py`)

Wenn die Kostenstelle zum ersten Mal vorkommt:

1. **Icon** ablegen: `assets/xyz_icon.png` (PNG, ~quadratisch, ~150–200 px reicht).
2. **`LAGER_ICONS`** ergänzen, z. B. `"C-IGLU": ASSETS / "iglu_icon.png"`.
3. **`LAGER_HINWEISE`** ergänzen — Suchwörter im **Buchungstext** für Eingänge ohne KS (TWINT/UBS), z. B. `("iglu", "iglu camp")`.
4. **`lagerinfo.json`**: neuen Eintrag mit KS, Daten, Köpfen, optional `team_beitrag_max`.
5. Optional **`paycode_aus_buchung`**: Regex `(?:FBC|BTP|…|TB)\s+([A-Z0-9]{4})` — neues Kürzel eintragen, falls Rückzahlungstexte den Code anders schreiben.

Kostenstellen-Kürzel in der Buchhaltung: `C-BTP`, `C-YBC`, `C-FBC`, `C-OBC`, …

---

## Was das PDF enthält

1. **Deckblatt / Aufwand** — gruppiert nach Kontenplan  
2. **Ertrag + Fazit** — Saldo, Lagerspenden, ggf. Missionsabend, Überschuss, Vorjahr, übriges Geld  
3. **Pro Person** — Essen, Übernachtung (mit Küche-vom-Haus-Logik), Hinweise  
4. **Zahlungen & Kosten** — TWINT, SumUp, eBill, UBS; optional Team/Familien und Rückzahlungsübersicht  
5. **Kontoauszüge** — nur Konten **mit Buchungen**; **Bankkonten (1021, 1090, …) werden weggelassen** (redundant zu Gegenkonten)  
6. **Belege** — Aufwand dieser KS, sortiert nach Betrag, **4200 Lebensmittel am Schluss**

Leere PDF-Seiten der Kontoauszüge werden übersprungen; Anhänge werden auf A4 skaliert.

---

## Buchhaltungslogik (Fazit)

| Begriff | Logik |
|---------|--------|
| **Saldo dieses Jahr** | Ertrag ohne Vorjahr (`3610`) und ohne Spenden (`3400`) minus Aufwand **ohne** Missionsabend-Durchlauf (`4900`, wenn durch Spenden gedeckt) |
| **Lagerspenden** | `3400` netto minus Missionsabend-Betrag |
| **Missionsabend** | Grösster `4900`-Betrag, wenn ≤ Spenden (`3400`) — reine Durchlaufposten |
| **Überschuss** | Saldo + Lagerspenden (Missionsabend netto 0) |
| **Übriges Geld** | Überschuss + Vorjahr |

Teilnehmerbeiträge fliessen in Ertrag (`3500`); Zahlungskanal wird separat aus Journal + `keinekostenstelle` geschätzt.

---

## Zahlungen (TWINT / SumUp / eBill)

- **Paycodes** (4-stellig, Beleg-Nr. / Ende TB-Text) aus KS-Buchungen → Zuordnung Eingänge in `keinekostenstelle`.
- **Rechnungs-IDs** aus Kontoauszügen `3500`/`1100` derselben KS.
- **eBill:** Betreff enthält `ebill` (ohne Bindestrich) → CHF **0.40** Gebühr pro Zahlung (`EBILL_GEBUEHR`).
- **TWINT:** über Konto `1091` / Gebühren `6300` mit «twint» im Text.
- **SumUp:** Konto `1092`, Gebühren `6300` ↔ `1092`.

Ohne `data/keinekostenstelle` fehlen oft UBS/eBill-Anteile — Export lohnenswert.

---

## Teilnehmerbeiträge & Rückzahlungen (3500)

- Pro **Paycode**: Eingänge (Haben 3500) minus Ausgänge (Soll 3500) = **Netto** für Team/Familien.
- **Rückzahlungen** auf der Zahlungsseite:
  - **Reduktion:** Text deutet auf Ermässigung (auch zerstückelt im Banktext, z. B. «ERM S»).
  - **Rückerstattung:** sonst (z. B. «Teilrückerstattung», volle Rückzahlung).
- Hinweistext: Rückzahlungen sind in den Beiträgen oben **bereits abgezogen** (Netto je Paycode).

Beispiel FBC: 600 rein, 200 raus → Netto 400 → **Team** (bei `team_beitrag_max: 400`).

---

## Belege

- Dateiname muss mit **`Beleg <ID>`** beginnen; `<ID>` = Buchungs-ID aus dem Journal (führende Nullen egal).
- Es werden nur Belege zu **Aufwandsbuchungen** dieser KS angehängt.
- Duplikat-Dateien (Unicode-Varianten): PDF bevorzugt.

Fehlende Belege: kein Abbruch — Buchung erscheint trotzdem in Kontoauszügen.

---

## Typische Probleme

| Symptom | Prüfen |
|---------|--------|
| «Kei Buchige mit Kostenstelle …» | Journal enthält die KS? Richtiger `data/xlsx`-Export? |
| Zahlungen unvollständig / nur SumUp | `data/keinekostenstelle/Buchungsjournal.xlsx` vorhanden? Paycodes in TB-Rechnungen? |
| TWINT falsch zugeordnet | `LAGER_HINWEISE` für diese KS; Lagername im RaiseNow-Text? |
| Saldo «falsch» vs. Excel | Missionsabend/4900 und Spenden/3400 — siehe Fazit-Tabelle |
| Icon fehlt | `assets/…` + `LAGER_ICONS` |
| Steckbrief falsch geschätzt | `lagerinfo.json` anpassen, `--ja` erneut |
| Riesiges PDF | Normal bei vielen Belegen; Bank-PDFs sind bereits ausgefiltert |

---

## Checkliste für KI-Agenten (Wiederholung «next camp»)

1. `data/xlsx/Buchungsjournal.xlsx` lesen → **`kostenstellen_im_journal`** (welche KS ist im Export?).
2. Prüfen ob **`lagerinfo.json`** Eintrag für diese KS existiert; sonst aus Buchungstexten **schätzen** (Ort, Datum, Köpfe) und User ggf. korrigieren lassen.
3. **`assets/<lager>_icon.png`** vorhanden? Sonst **`LAGER_ICONS`** + **`LAGER_HINWEISE`** in `generate_abrechnung.py` ergänzen.
4. Speziallager (FBC): **`team_beitrag_max`** in `lagerinfo.json` setzen.
5. **`uv run generate_abrechnung.py --ja --kostenstelle <KS>`** ausführen.
6. Konsolenoutput prüfen: Aufwand/Ertrag/Saldo, Zahlungszeilen, Team/Familien, Rückzahlungen.
7. PDF-Pfad unter `output/` nennen.
8. **Nicht** committen, es sei denn, der User verlangt es explizit (`output/`, `data/*` meist ignoriert).

Relevante Code-Stellen:

- Pfade / Icons / Hinweise: Anfang `generate_abrechnung.py`
- Fazit: `auswerten()`, `zeichne_seite_ertrag()`
- Zahlungen: `analysiere_zahlungen()`, `analysiere_beitraege()`, `zeichne_seite_zahlungen()`
- PDF-Anhang: `pdfs_zum_anhaengen()` (filtert `BANK_KONTEN`)
- Belege: `index_belege()`, `belege_zum_anhaengen()`
- Küche/Übernachtung: `ESSEN_ANNAHME_PRO_TAG`, `kueche_vom_haus`

---

## Abhängigkeiten

- Python ≥ 3.12  
- `openpyxl`, `pymupdf` (via `uv sync`)

Kein separater Webserver, kein HTML-Output mehr — **finale Abgabe ist das PDF**.
