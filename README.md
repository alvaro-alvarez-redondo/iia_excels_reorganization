# IIA Excel Reorganization Workflow

This repository contains a complete Python workflow for reorganizing historical Excel workbooks into a standardized workbook structure.

## What the workflow does

The transformer currently supports the rules you specified:

- keeps every source sheet but writes it with the **same name in lowercase**
- converts hierarchical labels into explicit metadata columns:
  - `hemisphere`
  - `continent`
  - `country`
  - `unit`
  - `footnotes`
- preserves the year or period headers exactly as they appear in the source workbook
- extracts footnotes from country labels by taking every `(...)` segment, removing parentheses, and joining notes with `; `
- preserves source cell colors on copied data rows
- keeps repeated country/entity rows exactly as they appear in the source workbook
- assigns `unit` automatically from the document category, sheet variable, and source product using the rules you provided
- harmonizes reviewed document names into the canonical `r_iia_<yearbook>_<year>_<page_start>_<page_end>_<english_product>` format
- derives missing yearbook metadata from the folder path, for example `raw inputs/trade/extracted_pages_1938_39/...` becomes `trade` and `1938`
- strips source suffixes such as `sup`, `prod`, `rend`, `imp`, `exp`, and `num` before translating the product portion of the document name
- supports both the standard FAO unit rules and the special `inputs` unit rules
- includes automated tests and a GitHub Actions CI workflow

## Project structure

```text
.
├── .github/workflows/ci.yml
├── config/example.units.yml
├── pytest.ini
├── setup.py
├── src/iia_excel_reorg/
│   ├── __init__.py
│   ├── cli.py
│   ├── config.py
│   ├── naming.py
│   ├── transformer.py
│   ├── unit_rules.py
│   └── xlsx_io.py
└── tests/test_transformer.py
```

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[dev]
```

## Configuration

Create a YAML file that supplies the metadata needed by the unit and naming rules.

Example:

```yaml
unit_mode: standard

document_categories:
  reviewed_239_239azucar_caña_brutaprod: 1
  reviewed_466_475arrozimp_exp: 2

product_aliases:
  tea: te

product_translations:
  azucar cana bruta: raw cane sugar
  arroz: rice

unit_overrides:
  imports: tonnes
```

### Config fields

- `unit_mode`: `standard` or `inputs`
- `document_categories`: maps each original or canonical document stem to its category number, which is required by the unit assignment logic
- `product_aliases`: optional mapping from extracted source products to the canonical product names used in the unit rules
- `product_translations`: optional mapping from extracted source products to the English product slug used in the harmonized output filename
- `unit_overrides`: optional explicit unit override by sheet name or by `document_stem:sheet_name`
- `include_sheets`: optional list of sheet names to process

You can use the example file at `config/example.units.yml` as a starting point.

> Note: the parser intentionally supports the simple YAML structure used for this project configuration.

## Folder structure

### Expected input structure

The tool recognises a two-level convention used for IIA yearbook scans.  
The critical segment is a directory named `extracted_pages_YYYY_YY` (where `YYYY` is the full four-digit year and `YY` is the two-digit end year).  
The directory that sits **directly above** `extracted_pages_YYYY_YY` becomes the **yearbook** name.

```text
<input_root>/
└── <yearbook>/                        ← any descriptive name, e.g. "trade"
    └── extracted_pages_1938_39/       ← marks the year boundary; YYYY = 1938
        ├── reviewed_466_475arrozimp_exp.xlsx        ← file directly inside
        └── crops/                                   ← optional sub-topic folder
            └── reviewed_239_239azucar_caña_brutaprod.xlsx
```

Files that do **not** sit inside any `extracted_pages_YYYY_YY` directory are still processed but land directly in the output root without any sub-directory nesting.

### Generated output structure

The output mirrors the input hierarchy under two levels of sanitized folder names:

```text
<output_root>/
└── iia_extracted_pages_YYYY/          ← top-level year bucket, e.g. iia_extracted_pages_1938
    ├── r_iia_<yearbook>_<year>_<start>_<end>_<product>.xlsx   ← file directly inside extracted_pages
    └── iia_<subtopic>_YYYY/           ← only created when a sub-topic folder exists
        └── r_iia_<yearbook>_<year>_<start>_<end>_<product>.xlsx
```

#### Concrete example

Input tree:

```text
raw_inputs/
└── trade/
    └── extracted_pages_1938_39/
        ├── reviewed_466_475arrozimp_exp.xlsx
        └── crops/
            └── reviewed_239_239azucar_caña_brutaprod.xlsx
```

Generated output tree:

```text
output/
└── iia_extracted_pages_1938/
    ├── r_iia_trade_1938_466_475_rice.xlsx
    └── iia_crops_1938/
        └── r_iia_trade_1938_239_239_raw_cane_sugar.xlsx
```

#### Folder and file name rules

- Spaces in any folder or file name segment are replaced with `_`.
- Consecutive underscores are collapsed to a single `_`.
- Leading and trailing underscores are stripped.
- The yearbook name is taken from the folder that sits immediately above `extracted_pages_YYYY_YY` and normalised with the rules above.
- The product segment is stripped of trailing suffixes (`sup`, `prod`, `rend`, `imp`, `exp`, `num`) and then translated to its English equivalent using `product_translations` from the config.

## Usage

Transform a single workbook:

```bash
iia-excel-reorg path/to/source.xlsx output/ --config config/example.units.yml
```

Transform every workbook in a directory:

```bash
iia-excel-reorg path/to/input_dir output/ --config config/example.units.yml
```

Each generated workbook is written using the harmonized document name inside the automatically created output subdirectory.

## Transformation rules implemented

### Sheet names

Every processed worksheet is copied into the output workbook with the original sheet name converted to lowercase.

### Headers

The output header row is always:

```text
hemisphere | continent | country | unit | footnotes | <year/period columns...>
```

Year and period labels are preserved exactly as they appear in row 1 of the source sheet.

### Hierarchy extraction

The parser treats rows containing variants of `HÉMISPHÈRE` / `HEMISPHERE` as hemisphere labels.

The parser treats continent rows such as:

- `EUROPE`
- `AMÉRIQUE`
- `ASIE`
- `AFRIQUE`
- `OCÉANIE`

as continent labels.

These structural rows are not written as data rows. Instead, they populate the `hemisphere` and `continent` columns for subsequent country rows.

### Footnote extraction

Country labels such as:

```text
Belgique-Luxembourg (reexports) (special case)
```

become:

- `country = Belgique-Luxembourg`
- `footnotes = reexports; special case`

If notes contain references to units, those references stay in `footnotes`; they do not alter the `unit` assignment.

### Harmonized document names

Reviewed source names are converted with these rules:

- `reviewed_` becomes `r_`
- missing agency defaults to `iia`
- yearbook metadata comes from the folder path, for example `raw inputs/trade/extracted_pages_1938_39` becomes `trade` and `1938`
- the product segment is stripped of trailing suffixes like `sup`, `prod`, `rend`, `imp`, `exp`, and `num`
- the remaining product is translated to English for the output filename

Examples:

```text
reviewed_239_239azucar_caña_brutaprod -> r_iia_trade_1938_239_239_raw_cane_sugar
reviewed_466_475arrozimp_exp -> r_iia_trade_1938_466_475_rice
```

### Unit assignment

The workflow now derives `unit` from:

- the sheet variable (`area`, `production`, `imports`, `exports`, etc.)
- the extracted source product name
- the document category supplied in config
- the configured rule mode (`standard` or `inputs`)

This replaces the provisional sheet-to-unit mapping with the rule-based logic you specified.

### Color preservation

The workflow preserves source fills for:

- row label cells copied into metadata columns
- numeric year/period value cells copied into the output sheet

## Development

Run tests locally:

```bash
pytest
```

Run the CLI module directly:

```bash
python -m iia_excel_reorg.cli path/to/source.xlsx output/ --config config/example.units.yml
```

## GitHub Actions

The repository includes a CI workflow in `.github/workflows/ci.yml` that runs `pytest` on every push and pull request.
