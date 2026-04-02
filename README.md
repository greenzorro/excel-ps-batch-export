# excel-ps-batch-export

[🇬🇧 EN](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README.md) | [🇨🇳 中文](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README_ZH_CN.md)

Python script for reading a PSD template and applying contents in a spreadsheet to export plenty of images. It's an alternative to Photoshop Image > Variables > Define.

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

Here's what you do using my Python script.

1. Edit content in a spreadsheet.
2. Run src/psd_renderer.py

That's it. Images will be just there for you. All you need is a Python environment with a few packages.

## Download

Two ways of downloading:

1. `git clone https://github.com/greenzorro/excel-ps-batch-export.git`
2. click the button Code, then Download ZIP

## Setup

For the first time, you'll need some basic setup:

1. Place your PSD template files in the `workspace/` directory.
2. Edit PSD template file. Rename all changeable layers or groups with this pattern:
    - Format: `@Variable_name#Operation_Parameter`
    - A layer name may be like: `@badge#v`, `@description#t_p` or `@bg#i`
    - `@` tells the script that this layer is changeable, go and get content from the spreadsheet
    - `Variable_name` should appear in the spreadsheet as column head
    - `#Operation_Parameter` tells the script what to do with the layer
    - `#v` to set visibility according to TRUE/FALSE in the spreadsheet
    - `#t` to replace a text layer content with spreadsheet data, parameters including:
        - Text align left top by default
        - `_c` for horizontally center alignment
        - `_r` for horizontally right alignment
        - `_a[angle]` to rotate text by specified angle (e.g., `_a15` rotates 15° clockwise, `_a-30` rotates 30° counterclockwise)
        - `_p` for paragraph with text wraping, fill the paragraph text layer in PSD with at least one line
        - `_pm` for vertically middle alignment
        - `_pb` for vertically bottom alignment
        - All these parameters work together, like `#t_c_a15`, `#t_r_p`
        - Alignment set in PSD will not affect the result, the program only checks layer names
    - `#i` to fill a pixel layer with the image whose file path is written in the spreadsheet, parameters including:
        - Scale mode: `_cover` (crop mode, default) fills the layer by cropping excess, `_contain` (letterbox mode) fits the entire image with transparent padding
        - Alignment (9-grid): `_lt` (left-top), `_ct` (center-top), `_rt` (right-top), `_lm` (left-middle), `_cm` (center, default), `_rm` (right-middle), `_lb` (left-bottom), `_cb` (center-bottom), `_rb` (right-bottom)
        - Examples:
            - `@产品图#i` - default: cover + center
            - `@产品图#i_cover` - cover + center
            - `@产品图#i_contain` - contain + center
            - `@产品图#i_cover_rb` - cover + right-bottom
            - `@产品图#i_contain_lt` - contain + left-top
    - Notes:
        - **Do not use cmd/ctrl+T to scale changeable text layers**. Adjust their sizes only via font size attribute, otherwise the script will get wrong text sizes from the PSD file. If you already did, make new text layers to replace them.
        - For text that needs rotation, only add the rotation parameter in the layer name. **Keep the text layer horizontal/straight in PSD**, do not rotate it manually, otherwise the script may not be able to correctly read the layer's position information, causing the output to be misaligned.
3. Run `src/xlsx_generator.py`. Your XLSX files will appear, with columns ready.
4. Edit XLSX file. Python reads the first sheet, put your data there. Or you may follow the example, put your data in another sheet and use Excel formulas in the first one to read and calculate everything. It's especially useful when you want to toggle layer visibility. DO NOT delete the first `File_name` column, leave it blank to use the default file name format(image_1, image_2, etc).
5. Put everything else the templates need in `workspace/assets` folder, including fonts in `workspace/assets/fonts/`, background images, etc. Make sure the path to image assets match the data in the spreadsheet. Image paths must be relative to the workspace directory (e.g., `assets/1_img/image.jpg`).
6. Configure fonts in `workspace/fonts.json` (optional). If you have multiple PSD templates with different font requirements, create a `fonts.json` file in the workspace directory to specify which font file each template should use:
   ```json
   {
     "_comment": "字体配置文件 - 为每个PSD模板指定对应的字体文件",
     "1": "AlibabaPuHuiTi-2-85-Bold.ttf",
     "2": "SourceHanSansCN-Medium.otf",
     "product": "CustomFont.ttf"
   }
   ```
   The key is the PSD file prefix (part before the first `#`), and the value is the font filename in `workspace/assets/fonts/`. If not configured, the default font `workspace/assets/fonts/AlibabaPuHuiTi-2-85-Bold.ttf` will be used.

Looks complicated huh? Trust me, it's way more complicated doing the same thing using Photoshop. And once you've done setting up, this would be your life saver.

## Export

When it comes to exporting. Things become a piece of cake:

1. Paste content in the spreadsheet.
2. Run src/psd_renderer.py

I even made another script (src/file_monitor.py) to monitor the spreadsheet and export images automatically once the spreadsheets are modified.

## Clipboard Importer

For even faster workflow, you can use the clipboard_importer.py script:

1. Copy table data to clipboard (from Excel, web tables, etc.)
2. Run `python src/clipboard_importer.py`
3. Select target Excel file (if multiple)
4. Data is automatically written to Excel and images are generated

Now you don't even need to open Photoshop or Excel.

## Multi-file Processing

This tool supports processing multiple PSD templates with one Excel file. Here's how it works:

- **Grouping by Prefix**: All PSD files in the same directory are grouped by their prefix. The prefix is defined as the part of the filename before the first hash (`#`). For example:
  - `product_intro#templateA.psd` and `product_intro#templateB.psd` share the same prefix `product_intro`
- **Shared Excel**: For each group, a single Excel file (named `[prefix].xlsx`) is created. This Excel file contains variables from all PSDs in the group.
- **Batch Export**: When running `src/psd_renderer.py [prefix] ...`, the script will process all PSDs in the group. Each row in the Excel will generate one image for every PSD in the group. The output image filenames include the PSD's suffix (e.g., `image_1_templateA.jpg`).

Example:
  - PSD files: `campaign#summer.psd`, `campaign#winter.psd`
  - Excel file: `campaign.xlsx`
  - Command: `python src/psd_renderer.py campaign jpg`
  - Output: For each row in `campaign.xlsx`, two images are generated: `image_1_summer.jpg`, `image_1_winter.jpg`, etc. (assuming File_name column is empty, otherwise uses the File_name value)

## Advanced Features

### Data Transformation Rules

For templates that require complex data processing, you can use JSON-based transformation rules. When a `.json` file exists in the workspace directory alongside your template:

**How it works:**
1. Edit `_raw.csv` with your raw data
2. The system automatically applies transformation rules defined in `.json`
3. Processed data is written to `.xlsx` for rendering

**Supported transformations:**
- `direct` - Copy field values directly
- `conditional` - Copy only when a parent field is not empty
- `template` - Combine multiple fields into one (e.g., filename generation)
- `derived` - Boolean values based on other fields
- `derived_raw` - Boolean values based on raw field existence

**Example:** Templates 1, 2, and 3 include transformation rules. See `transform_guide.md` for detailed documentation.

## Prerequisite

### Install Dependencies

Install all required dependencies using the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

## Usage Guide

### Basic Usage

```bash
# Basic command format
python src/psd_renderer.py [Excel_file_prefix] [output_format] [output_directory (optional)]

# Examples
python src/psd_renderer.py 1 jpg                              # Default output to export/ directory
python src/psd_renderer.py 1 jpg output/custom               # Custom output directory
python src/psd_renderer.py 1 jpg /absolute/path/to/output    # Absolute path output directory
```

**Note**: Font files are configured via `workspace/fonts.json`. If not configured, the default font `workspace/assets/fonts/AlibabaPuHuiTi-2-85-Bold.ttf` will be used.

**Output Directory Options**:
- When no output directory is specified, images are saved to the default `export/` directory
- When a relative path is provided (like `output/custom`), it's relative to the project root directory
- When an absolute path is provided (like `/Users/username/Desktop/rendered`), images are saved to that location
- Each export creates a timestamped subdirectory to prevent file conflicts (e.g., `20260402_162657_1/`)

## Thanks

Special thanks to [psd-tools](https://github.com/psd-tools/psd-tools) for providing such powerful APIs to interact with PSD files. Therefore I could utilize the power of Photoshop at image editing and Excel/Python at data processing.

---

Created by [Victor42](https://victor42.work/) & [Agent Vik](https://github.com/agent-vik)
