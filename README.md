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
2. Run psd_renderer.py

That's it. Images will be just there for you. All you need is a Python environment with a few packages.

## Download

Two ways of downloading:

1. `git clone https://github.com/greenzorro/excel-ps-batch-export.git`
2. click the button Code, then Download ZIP

## Setup

For the first time, you'll need some basic setup:

1. Place your PSD template files here.
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
    - `#i` to fill a pixel layer with the image whose file path is written in the spreadsheet
    - Notes:
        - Do not use cmd/ctrl+T to scale changeable text layers. Adjust their sizes only via font size attribute, otherwise the script will get wrong text sizes from the PSD file. If you already did, make new text layers to replace them.
        - For text that needs rotation, only add the rotation parameter in the layer name. Keep the text layer horizontal/straight in PSD, do not rotate it manually, otherwise the script may not be able to correctly read the layer's position information, causing the output to be misaligned.
3. Run `xlsx_generator.py`. Your XLSX files will appear, with columns ready.
4. Edit XLSX file. Python reads the first sheet, put your data there. Or you may follow the example, put your data in another sheet and use Excel formulas in the first one to read and calculate everything. It's especially useful when you want to toggle layer visibility. DO NOT delete the first `File_name` column, leave it blank to use the default file name format(image_1, image_2, etc).
5. Put everything else the templates need in `assets` folder, including fonts, background images, etc. Make sure the path to image assets match the data in the spreadsheet.

Looks complicated huh? Trust me, it's way more complicated doing the same thing using Photoshop. And once you've done setting up, this would be your life saver.

## Export

When it comes to exporting. Things become a piece of cake:

1. Paste content in the spreadsheet.
2. Run psd_renderer.py

I even made another script (file_monitor.py) to monitor the spreadsheet and export images automatically once the spreadsheets are modified.

## Clipboard Importer

For even faster workflow, you can use the clipboard_importer.py script:

1. Copy table data to clipboard (from Excel, web tables, etc.)
2. Run `python clipboard_importer.py`
3. Select target Excel file (if multiple)
4. Data is automatically written to Excel and images are generated

Now you don't even need to open Photoshop or Excel.

## Multi-file Processing

This tool supports processing multiple PSD templates with one Excel file. Here's how it works:

- **Grouping by Prefix**: All PSD files in the same directory are grouped by their prefix. The prefix is defined as the part of the filename before the first hash (`#`). For example:
  - `product_intro#templateA.psd` and `product_intro#templateB.psd` share the same prefix `product_intro`
- **Shared Excel**: For each group, a single Excel file (named `[prefix].xlsx`) is created. This Excel file contains variables from all PSDs in the group.
- **Batch Export**: When running `psd_renderer.py [prefix] ...`, the script will process all PSDs in the group. Each row in the Excel will generate one image for every PSD in the group. The output image filenames include the PSD's suffix (e.g., `image_1_templateA.jpg`).

Example:
  - PSD files: `campaign#summer.psd`, `campaign#winter.psd`
  - Excel file: `campaign.xlsx`
  - Command: `python psd_renderer.py campaign AlibabaPuHuiTi-2-85-Bold.ttf jpg`
  - Output: For each row in `campaign.xlsx`, two images are generated: `image_1_summer.jpg`, `image_1_winter.jpg`, etc. (assuming File_name column is empty, otherwise uses the File_name value)

## Prerequisite

### Install Dependencies

It's recommended to use the `requirements.txt` file to install all dependencies:

```bash
pip install pillow pandas openpyxl psd-tools tqdm
```

### Testing

The project includes a comprehensive test suite to ensure code quality and functionality:

```bash
# Run all tests (recommended)
python tests/run_tests.py all

# Run specific test file
python -m pytest tests/test_simple.py -v

# Generate coverage report
python tests/run_tests.py coverage
```

**Test Coverage**: 173 tests covering core functionality, business logic, error handling, boundary conditions, performance scenarios, and text rotation with strict validation standards.

## Usage Guide

### Basic Usage

```bash
# Basic command format
python psd_renderer.py [Excel_file_prefix] [font_file] [output_format]

# Example
python psd_renderer.py 1 AlibabaPuHuiTi-2-85-Bold.ttf jpg
```

## Thanks

Special thanks to [psd-tools](https://github.com/psd-tools/psd-tools) for providing such powerful APIs to interact with PSD files. Therefore I could utilize the power of Photoshop at image editing and Excel/Python at data processing.

---

Created by [Victor_42](https://victor42.work/)
