# ELECTOR Card Generator

Generate printable **PNG** and **PDF** cards (and optional sticker sheets) from `ELECTOR.xlsx` using party/type templates, skill icons, and portrait avatars.

Main script: `generate_cards.py`

---

## Directory structure

```
.
├── generate_cards.py          # Main generation script
├── ELECTOR.xlsx               # Card data (source of truth)
├── requirements.txt           # Python dependencies
├── Sticker-temp.png           # Letter-size sticker sheet template (3×3)
├── Result-sample.png          # Reference sample (optional)
├── ally-Template-1.png        # Design reference templates (optional)
├── can-Template-2.png
├── eff-Template-3.png
│
├── fonts/                     # Arial fonts used for card text (include when sharing)
│   ├── Arial.ttf
│   └── Arial Bold.ttf
│
├── ELECTOR/
│   ├── cardtemps/             # Party × type card backgrounds
│   │   ├── demally.png / demcan.png / demeff.png
│   │   ├── repally.png / repcan.png / repeff.png
│   │   ├── greenally.png / greencan.png / greeneff.png
│   │   ├── libally.png / libcan.png / libeff.png
│   │   ├── indally.png / indcan.png / indeff.png
│   │   ├── unially.png / unican.png / unieff.png
│   │   └── cardback1.png / cardback2.png
│   ├── iconslabels/           # Party emblems + skill badges
│   │   ├── DEMOCRAT.png, REPUBLICAN.png, GREEN.png, …
│   │   ├── SYNERGIZE.png, IMMUNITY.png, SHIELDED.png, …
│   │   └── rarity.png
│   └── avatar/                # Per-card portraits
│       ├── default-img.png    # Fallback when no numbered avatar exists
│       └── 001.png            # Optional: CARDNUM-matched (e.g. 001.png)
│
└── output/                    # Generated files (created/updated by the script)
    ├── png/                   # Individual cards as PNG
    ├── pdfs/                  # Individual cards as PDF (letter page, card centered)
    └── stickers/              # Sticker sheets (PNG + PDF, 9 cards each)
```

---

## Prerequisites

- Python 3.10+ recommended
- Card text uses **Arial** from `fonts/` (bundled with the project). Keep that folder when giving the project to a client so it runs the same on any machine.

### Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Dependencies (`requirements.txt`):

- `openpyxl` — read Excel
- `Pillow` — compose card images
- `reportlab` — write PDFs

---

## How it works

1. Reads every data row from `ELECTOR.xlsx` (active sheet, header row skipped).
2. Picks a template from `ELECTOR/cardtemps/` based on **PARTY** + **TYPE**.
3. Places a portrait from `ELECTOR/avatar/` (or a path you pass with `--avatar`).
4. Draws stats, name, subtype, effect text, party emblem, skills, and footer.
5. Writes each card to `output/png/` and `output/pdfs/`.
6. Unless `--no-stickers` is set, packs cards into 3×3 sticker sheets under `output/stickers/`.

### Card size

| Output | Spec |
|--------|------|
| Individual card | Poker **2.5″ × 3.5″** @ **400 DPI** → **1000 × 1400** px |
| Sticker sheet | US Letter @ **300 DPI**, **9 cards** (60×85 mm slots) |

### Template mapping

| Excel `PARTY` | Prefix | + `TYPE` | Suffix | Example file |
|---------------|--------|----------|--------|--------------|
| DEMOCRAT | `dem` | ALLY | `ally` | `demally.png` |
| REPUBLICAN | `rep` | CANDIDATE | `can` | `repcan.png` |
| GREEN | `green` | EFFECT | `eff` | `greeneff.png` |
| LIBERTARIAN | `lib` | … | … | `libally.png` |
| INDEPENDENT | `ind` | … | … | `indcan.png` |
| UNIVERSAL | `uni` | … | … | `unieff.png` |

File pattern: `{prefix}{suffix}.png` inside `ELECTOR/cardtemps/`.

### Excel columns → card

Columns are matched by **header name** (order-tolerant). New workbook uses `SKILLS1`/`SKILLS2`; older `SKILLS` still works.

| Column | Used for |
|--------|----------|
| ID | Footer serial `# EL-{id}` |
| NAME | Card title |
| PARTY | Template, emblem, alliance label |
| TYPE | Layout (`ALLY` / `CANDIDATE` / `EFFECT`) |
| SUBTYPE | Subtype pill (ally / effect) |
| MP | Candidate left stat |
| REGEN | Candidate right stat |
| COST | Ally / effect power (or cost) |
| VC | Ally votes |
| POWER COST | Circled ability cost in effect text |
| EFFECT | Ability text |
| SKILLS1 / SKILLS (legacy) | First skill badge(s) → `iconslabels/{SKILL}.png` |
| SKILLS2 | Second skill badge (optional) |
| SET | Set badge (default `BASE`) |
| CARDNUM | Print run `001/N` and avatar `{CARDNUM}.png` |
| RARITY | Rarity stars in footer |
| YEAR | Copyright year (default 2026) |
| SETTOTAL | Print-run denominator (fallback: row count) |
| PARTYABB | Party abbreviation under emblem (fallback: built-in map) |
| IMAGE | Optional avatar filename stem (after CARDNUM) |

### Avatars

`--avatar` can be either:

- A **folder** (default: `ELECTOR/avatar/`) — looks for `{CARDNUM}.png` (zero-padded to 3 digits), else `default-img.png`
- A **single image file** — used for every selected card

---

## Usage

Activate the venv first:

```bash
source .venv/bin/activate
```

### Generate everything

```bash
python generate_cards.py
```

### Prototype (Billy Snow only)

```bash
python generate_cards.py --prototype
```

### Specific card IDs

```bash
python generate_cards.py --ids 0000001 0000151
```

### First N cards

```bash
python generate_cards.py --limit 10
```

### Custom Excel / output / avatars

```bash
python generate_cards.py --xlsx ELECTOR.xlsx --out output --avatar ELECTOR/avatar
```

Use one shared portrait for all cards:

```bash
python generate_cards.py --avatar path/to/portrait.png --limit 5
```

### Cards only (skip sticker sheets)

```bash
python generate_cards.py --no-stickers
```

### Debug layout overlays

Draws labeled boxes for each layout region on the PNG (useful when tweaking positions):

```bash
python generate_cards.py --prototype --debug-layout --no-stickers
```

### CLI reference

```
python generate_cards.py [-h]
  --xlsx PATH          Excel workbook (default: ELECTOR.xlsx)
  --avatar PATH        Avatar folder or single image (default: ELECTOR/avatar)
  --out PATH           Output root (default: output)
  --ids [IDS ...]      Only these card IDs
  --limit N            Only first N matching cards
  --prototype          Billy Snow only
  --debug-layout       Overlay layout boxes on PNGs
  --no-stickers        Skip sticker sheet PNG/PDF generation
```

---

## Outputs

After a successful run:

```
output/
├── png/
│   └── 001_BILLY_SNOW.png          # {CARDNUM}_{NAME}.png
├── pdfs/
│   └── 001_BILLY_SNOW.pdf          # Same stem; card centered on letter page
└── stickers/
    ├── sticker_sheet_001.png
    ├── sticker_sheet_001.pdf
    ├── sticker_sheet_002.png
    └── …
```

Sticker sheets are filled in generation order, **9 cards per sheet**.

---

## Adding or updating cards

1. Edit rows in `ELECTOR.xlsx` (keep the header row and column order).
2. Add matching templates under `ELECTOR/cardtemps/` if you introduce a new party/type combo.
3. Add skill PNGs under `ELECTOR/iconslabels/` named like the skill token (e.g. `SYNERGIZE.png`).
4. Optionally add `ELECTOR/avatar/{CARDNUM}.png` for that card’s portrait.
5. Re-run `python generate_cards.py` (or `--ids` / `--limit` for a subset).

---

## Notes

- Templates and icons must exist for every party/type/skill you reference, or that card will fail / skip the missing asset.
- Generated files under `output/` are overwritten when you regenerate the same card or sheet names.
