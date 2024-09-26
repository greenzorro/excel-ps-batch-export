# excel-ps-batch-export

[🇬🇧 EN](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README.md) | [🇨🇳 中文](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README_ZH_CN.md)

Python script for reading a PSD template and applying contents in a spreadsheet to export plenty of images. It's an alternative to Photoshop Image > Variables > Define.

📺 Manual export: [example video](https://www.bilibili.com/video/BV1nqxGe4ETf)  
📺 Auto export: [example video](https://www.bilibili.com/video/BV1qmx5etEY8)

This is what you do using Photoshop:

1. Edit content in a spreadsheet.
2. Save spreadsheet as a CSV file.
3. Define variables for layers in Photoshop.
4. Import the CSV file.
5. Export data sets as files. Unfortunately, you get .psd files.
6. Make a automate-batch to save PSD as JPG or PNG.
7. Use the automate-batch to get final images.

Here's what you do using my Python script.

1. Edit content in a spreadsheet.
2. Run batch_export.py

That's it. Images will be just there for you. All you need is a Python environment with a few packages.

## Download

Two ways of downloading:

1. `git pull` this repo
2. click the button Code, then Download ZIP

## Setup

For the first time usage, you'll need some basic setup:

1. Place your PSD template here, rename it to `{num}_template.psd`. Default num is 1. Increase num for extra working template-data pairs. Switch pairs in the script file.
2. Create a `{num}_data.xlsx` file here. Default num is 1. This will be your data template.
3. Edit PSD template file. Rename all changeable layers or groups with this pattern:
    - Format: `@Variable_name#Operation_type`
    - A layer name may be like this: `@badge#v`
    - `@` tells the script that this layer is changeable, go and get content from the spreadsheet
    - `Variable_name` should appear in the spreadsheet as column head.
    - `#Operation_type` tells the script what to do with the layer
    - `#v` to set visibility according to TRUE/FALSE in the spreadsheet
    - `#t` to replace a text layer content with spreadsheet data
    - `#t-c` or `#t-r` for text align center or right while replacing
    - `#i` to fill a pixel layer with the image whose file path is written in the spreadsheet
    - One thing to note: Do not use cmd/ctrl+T to scale changeable text layers. Adjust their sizes only via font size attribute, otherwise the script will get wrong text sizes from the PSD file. If you already did, make new text layers to replace them.
4. Edit XLSX file. Python reads the first sheet, put your data there. Or you may follow the example, put your data in another sheet and use Excel formulas in the first sheet to read and calculate everything. It's especially useful when you want to toggle layer visibility.
5. Put everything else the template needs in `assets` folder, including fonts, background images, etc.

Looks complicated huh? Trust me, it's way more complicated doing the same thing using Photoshop. And once you've done setting up, this would be your life saver.

## Export

When it comes to exporting. Things become a piece of cake:

1. Paste content in the spreadsheet.
2. Run batch_export.py

I even made another script to moniter the spreadsheet and export images automatically once the spreadsheets are modified.

## Prerequisite

```
pip install pillow pandas openpyxl psd-tools
```

## Thanks

Special thanks to [psd-tools](https://github.com/psd-tools/psd-tools) for providing such powerful APIs to interact with PSD files. Therefore I could ultilize the power of Photoshop at image editing and Excel/Python at data processing.
