# Boma → Pomproc catalog ETL

Transforms a **Boma product-catalog export** into the **Pomproc `Products`
import sheet** (target schema per `PomProc_Data_Engine_Spec.md` §4.1 and the
`Products` tab of `PomProc_database_templates.xlsx`).

## Install

```bash
python -m venv .venv && source .venv/bin/activate      # optional
pip install -r requirements.txt
```

## Run

```bash
python boma_to_pomproc.py \
  --input  boma_export.xls \
  --output products_pomproc.xlsx \
  --rejects rejects.xlsx \
  --categories-file PomProc_database_templates.xlsx
```

Key options (`--help` for all):

| Option | Purpose | Default |
|---|---|---|
| `--categories-file` | Workbook whose `Categories` sheet lists valid category names | built-in canon of 11 |
| `--mapping FILE.json` | Override the Rubrique→Category map, e.g. `{"Equipement":"Outillage"}` | built-in map |
| `--default-category X` | Fallback for unmapped rubriques (omit ⇒ such rows are rejected) | *(reject)* |
| `--lenient-categories` | Don't reject rows whose category isn't in the known set | strict |
| `--default-entity` | Value for the `Entity` column (`FFA`/`IIS`/`DC`) | `IIS` |
| `--supplier` | Value for the `Fournisseurs` column | `Boma` |

## What the script does

1. **Auto-detects the source format.** Boma "`.xls`" files are actually HTML
   tables; the script reads them with `lxml`, preserving raw cell text. Genuine
   `.xlsx`/`.xls` are also supported (`--input` sniffs the file signature).
2. **Parses European decimals correctly.** `"7,71" → 7.71`. A naïve
   `pandas.read_html` reads the comma as a thousands separator and turns it into
   `771` — this script avoids that entirely.
3. **Normalizes headers** via a synonym table (`Prix`/`Price`, `N° article`/`Ref`, …).
4. **Validates each row independently** — a bad row is logged and skipped, never
   blocking the rest. Missing Name/Reference/Rubrique, negative price, etc. go
   to the rejects report.
5. **Maps categories** (Boma `Rubrique` → Pomproc `CategoryName`) and checks the
   result against Pomproc's real category list.
6. **Generates a stable, unique SKU** (`SLUG(name)-REFERENCE`). Pomproc matches
   imports on SKU, and Boma provides none. The rule is deterministic, so
   **re-running the same file is idempotent**.
7. **Drops exact duplicate rows** (the sample has 3: refs 710522, 270422, 140800).
8. **Writes a formatted workbook**: `SKU`/`Reference` forced to **text** (protects
   codes like `MAC-002` and stops Excel mangling numeric refs), `Price` numeric,
   header frozen, sensible column widths, exact column order.

## Column mapping

| Boma column | Pomproc column | Rule |
|---|---|---|
| `N° article` | `Reference` | trimmed, kept as **text** |
| `N° article` + `Description` | `SKU` | generated: `SLUG(name)-REF`, unique |
| `Description` | `Name` | whitespace-collapsed |
| `Rubrique` | `CategoryName` | mapped (see below) |
| `Unité` | `Unit` | trimmed; blank → `Pièce` |
| `Quantité par emballage` | `PackageQty` | int ≥ 1; blank → 1 |
| `Prix` | `Price` | EU-decimal → float; blank → 0 |
| — | `Entity` | default `IIS` |
| — | `Active` | `yes` |
| — | `ImageURL` / `Associations` / `Variants` | blank (not in Boma) |
| — | `Fournisseurs` | `Boma` |

**Default Rubrique → Category map** (override with `--mapping`):

| Boma Rubrique | Pomproc Category |
|---|---|
| Chariots et mops | Matériel *(confirmed by template sample row 710522)* |
| Equipement | Matériel |
| Machines et accessoires | Machines |
| Nettoyage de vitres | Vitrerie |
| Poubelles et sacs poubelles | Poubelles / Sacs |
| Produits de nettoyage | Consommables |
| Hygiène et équipement sanitaire | Consommables |
| Protection et sécurité | EPI |

> Rows 2, 5, 6, 7, 8 in that table are reasonable business defaults, not
> statements from the spec. If your classification differs (e.g. *Hygiène* →
> `EPC / Protection`), pass a `--mapping` JSON to override.

## Deliberately dropped columns

Boma's eco fields — `Certificats`, `EcoCert`, `EU ecolabel`, `Nordic ecolabel`,
`Cradle to cradle`, `Plastic Second Life`, `FSC`, `Approprié alimentaire`,
`Recycled`, `Eco Score` — have **no field in the Pomproc `Product` schema** and
are dropped. If you want them retained, they'd need a schema extension (e.g. a
JSON `attributes` column on `Product`); the script logs their presence so the
data loss is visible.

## Sample run result

```
source rows read      : 163
products written      : 160
exact duplicates drop : 3
rows rejected         : 0
categories produced   : Consommables 73 · Matériel 43 · EPI 16 ·
                        Vitrerie 13 · Poubelles / Sacs 12 · Machines 3
```
