# FSAE BOM Builder

🇺🇸 English | [🇧🇷 Português](READMEpt.md)

Automated BOM (Bill of Materials) generator for an FSAE (Formula SAE) car,
extracting part data directly from **SolidWorks** via its COM API — custom
properties, native material, and mass — and assembling it into a standardized
BOM spreadsheet.

Built to remove manual, error-prone BOM transcription from CAD: instead of
copying part names, materials, and masses by hand into a spreadsheet, the
pipeline reads the SolidWorks assembly tree directly and produces a
structured, standardized output.

## ⚠️ Status: untested on a real SolidWorks machine

This code was written against the official SolidWorks API, but **has not
been tested on a real machine** (it was written on Linux, without
SolidWorks). Expect adjustments on first run — especially method
names/signatures, which can vary between SolidWorks versions (see the
troubleshooting notes below).

## What still requires manual work (by design)

- Filling in `FSAE_System`, `FSAE_Process_N`, etc. on each part in SolidWorks
  (a standardization GUI to streamline this is planned but not yet built).
- Confirming the fuzzy material matcher's suggestions — the script never
  auto-applies a match without flagging it for review.
- Fasteners and tooling are still filled in manually on the generated tab.

## Tech stack

| Layer | Technology |
|---|---|
| CAD integration | SolidWorks COM API via `pywin32` |
| Data processing | Python |
| Output | Excel (`.xlsx`) BOM spreadsheet |
| Environment | Windows (SolidWorks required), Python 3 |

## Project structure

```
files/
├── sw_extract/
│   ├── connector.py          # COM connection to SolidWorks
│   ├── extractor.py          # reads custom properties, native material, mass, and traverses the assembly tree
│   ├── material_matcher.py   # fuzzy corrector (SolidWorks material -> exact catalog name)
│   ├── translator.py         # converts extracted data into the format bom_builder.py expects
│   └── test_connection.py    # isolated connection test (run this first)
├── bom_builder.py            # builds the standardized BOM spreadsheet
├── demo_add_parts.py         # example/demo of adding parts to a BOM
└── extract_and_build.py      # full pipeline (run last)
```

## Getting started

**Prerequisites:** Windows with SolidWorks installed.

```bash
pip install pywin32
```

### Recommended test order (do not skip steps)

**1. Connection test on a simple part**

Open a part with a **known** mass (already weighed or hand-calculated), not
an assembly yet. From this folder (`files\`, the folder that *contains*
`sw_extract\` — don't `cd` into it, it's a relative import that only works
when run as a module from outside):

```bash
python -m sw_extract.test_connection
```

This should print the document title, custom properties, native material,
and mass. If something fails here, it's much easier to debug on a simple
part than on a full assembly — fix it before moving on.

**Most important thing to verify at this stage:** the `mass_kg` value
printed must match the part's real mass. `get_mass_properties()` assumes
SolidWorks always returns mass in kg, but that depends on the unit system of
the template used (an MMGS template, for example, would return grams) —
this has not been confirmed on a real machine yet. If this value comes back
wrong and nobody notices here, every auto-extracted material cost from then
on is silently wrong (potentially off by 1000x, kg vs. g).

Expected errors at this stage and what they usually mean:
- `AttributeError` on `Get5` → the SolidWorks version uses a different
  method signature. Try `Get4` or `Get3` (see comment in `extractor.py`).
- Native material comes back `None` → the part likely has no material
  assigned yet in SolidWorks (Material Editor).
- `GetActiveObject` fails → SolidWorks isn't open, or is open with no
  document loaded.

**2. Test with `FSAE_*` properties filled in manually**

Manually add (via SolidWorks: File > Properties > Custom) the properties
`FSAE_System`, `FSAE_PN_Base`, `FSAE_Suffix`, `FSAE_Details`,
`FSAE_Process_1`, `FSAE_Process_1_Use`, `FSAE_Process_1_Qty` on a test part,
then run `python -m sw_extract.test_connection` again (from `files\`) —
confirm the values show up correctly in the output.

**3. Traversal test on a small assembly**

Take a small, known sub-assembly (e.g. just the brake system) and test
`traverse_assembly()` in isolation before moving to the whole car — it's
easier to check that the quantity count and parts list match your manual
expectations.

**4. Full pipeline**

Only after the steps above work, run:

```bash
python extract_and_build.py --template path\to\template.xlsx --out path\to\output.xlsx
```

## Author

Built as part of an FSAE team's tooling, by a Computer Engineering student.
