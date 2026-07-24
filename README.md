# excel-ps-batch-export

[🇬🇧 EN](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README.md) | [🇨🇳 中文](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README_ZH_CN.md)

Python script for reading a PSD template and applying contents in a spreadsheet to export plenty of images. It's an alternative to Photoshop Image > Variables > Define.

You author PSD layer names and spreadsheet data, then review the images. An agent on your machine handles install and export runs.

📺 Example - Create spreadsheet

https://github.com/user-attachments/assets/a21f8b2d-310f-4f28-a873-6bd166c07955

📺 Example - Manual export

https://github.com/user-attachments/assets/c52c6e05-1bc9-4a2b-ae4c-b283a25067f6

📺 Example - Auto export

https://github.com/user-attachments/assets/bfd2d23f-84ec-4ea9-8874-523a298049be

This is [what you do using Photoshop](https://victor42.eth.limo/post-en/3650/):

1. Edit content in a spreadsheet.
2. Save spreadsheet as a CSV file.
3. Define variables for layers in Photoshop.
4. Import the CSV file.
5. Export data sets as files. Unfortunately, you get .psd files.
6. Make a automate-batch to save PSD as JPG or PNG.
7. Use the automate-batch to get final images.

With this project you edit the spreadsheet (and PSD once), then let the agent run the renderer. No Photoshop Variables / batch dance.

## Authoring templates

Data lives under `demo/` by default, or a custom data directory configured as `EPS_DATA_DIR`.

1. Put PSD templates in `workspace/`.
2. Rename changeable layers/groups with `@Variable_name#Operation_Parameter`, e.g. `@badge#v`, `@description#t_p`, `@bg#i`:
    - `@` marks a changeable layer; `Variable_name` must match a spreadsheet column
    - `#v` — visibility from TRUE/FALSE
    - `#t` — replace text; modifiers: `_c` / `_r` align, `_a[angle]` rotate (keep the layer straight in PSD), `_p` paragraph wrap, `_pm` / `_pb` vertical align in paragraph (with `_p`). Combinations like `#t_c_a15`, `#t_p_pm` work. PSD paragraph-align UI is ignored; names win.
    - `#i` — fill pixels from an image path in the sheet; scale `_cover` (default) / `_contain`; 9-grid align `_lt` `_ct` `_rt` `_lm` `_cm` (default) `_rm` `_lb` `_cb` `_rb`
    - **Do not** free-transform (`Cmd/Ctrl+T`) changeable text layers—set size via font size only
    - For rotation, only use `#t_a…` in the name; keep the layer horizontal in the PSD
3. Have the agent run `xlsx_generator` so column headers appear; edit the first sheet (or drive it with formulas from another sheet). Keep the `File_name` column (blank → default `image_1`, …).
4. Put fonts under `workspace/assets/fonts/` and other assets under `workspace/assets/`. Image paths in the sheet are relative to `workspace/` (e.g. `assets/1_img/image.jpg`).
5. Optional `workspace/fonts.json` maps PSD prefix → font filename.

Looks heavy? Doing the same with Photoshop Variables is worse. Once templates exist, daily work is “paste rows → ask the agent to export”.

## Day-to-day use

- Paste or edit rows in the spreadsheet, then ask the agent to export (or to watch the file with the monitor script).
- **Clipboard path:** copy a table (Excel/web) → ask the agent to run the clipboard importer → pick the target workbook if prompted → images generate.
- **Multi-PSD, one sheet:** files sharing a prefix before the first `#` (e.g. `campaign#summer.psd` + `campaign#winter.psd`) share `campaign.xlsx`. Each row yields one image per PSD; filenames include the suffix when `File_name` is empty.
- **Transform rules:** if `workspace/PREFIX.json` exists, edit `PREFIX_raw.csv`; rules write the renderable `.xlsx`. Types: `direct`, `conditional`, `template`, `derived`, `derived_raw`. See `transform_guide.md`.

## Thanks

Special thanks to [psd-tools](https://github.com/psd-tools/psd-tools) for APIs that let Photoshop stay the design tool while Excel/Python own the data.

---

Created by [Victor42](https://victor42.work/) & [Agent Vik](https://github.com/agent-vik)

---

# For Agent

Run batch PSD→image export with Python (no Adobe Photoshop app). After install, operate generators/renderers/importers when asked.

## Prerequisites

- Python 3 + pip
- `pip install -r requirements.txt`
- Readable PSD/Excel (or CSV) under the data directory’s `workspace/`

## Install / first-time setup

1. Clone the repo; install `requirements.txt`.
2. Default data root is `demo/`. Confirm `demo/workspace/` exists.
3. Optional custom root: copy `.env.example` → `.env`, set `EPS_DATA_DIR` to an absolute path.

Layout when using a custom data root:

```text
EPS_DATA_DIR/
├── workspace/   # PSD, xlsx/csv, fonts.json, assets/
├── export/
└── log.csv
```

If `EPS_DATA_DIR` is set but missing on disk, tools will **not** fall back to `demo/`.

## Usage

Work from the repo root. Data paths come from `src/config.py` (`demo/` or `EPS_DATA_DIR`).

Entrypoints:

- `python src/xlsx_generator.py` — build/refresh XLSX columns from PSD `@Variable` layer names
- `python src/psd_renderer.py PREFIX FORMAT [OUTPUT_DIR]` — batch render; `FORMAT` is `jpg` or `png`; default output is `DATA_DIR/export/` with a timestamp subfolder
- `python src/clipboard_importer.py` — clipboard table → Excel → render
- `python src/file_monitor.py` — watch workspace spreadsheets and auto-export on change
- `python src/transform.py` (also auto via renderer) — apply `workspace/PREFIX.json` rules from `PREFIX_raw.csv`

Typical flows:

1. Smoke-test: `python src/psd_renderer.py 1 jpg` → check `export/`.
2. After sheet edits: re-run `psd_renderer.py` for that prefix.
3. Clipboard table ready: run `clipboard_importer.py` and follow prompts.
4. Hands-free watch: start `file_monitor.py` and leave it running.
5. Stop when the requested export succeeds. Leave PSD layer naming and business spreadsheet content alone unless asked to edit those files.

## Hand off to the human

- Authoring `@variable#ops` layer names and spreadsheet rows
- Choosing production `EPS_DATA_DIR`
- Visual QA of exported images

## Red lines

- Do not require Photoshop GUI—this stack is `psd-tools` + Pillow
- Do not commit `.env` or client data directories
- Contracts: `notes.md` / `transform_guide.md`
