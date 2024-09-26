# excel-ps-batch-export

[🇬🇧 EN](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README.md) | [🇨🇳 中文](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README_ZH_CN.md)

这是一个Python脚本，用于读取PSD模板并将电子表格中的内容应用于模板，导出大量图像，替代Photoshop的变量定义功能。

📺 查看[演示视频](https://github.com/greenzorro/excel-ps-batch-export/blob/main/example/example.mp4).

在Photoshop你得这么干：

1. 在电子表格中编辑内容。
2. 将电子表格保存为CSV文件。
3. 在Photoshop中定义图层变量。
4. 导入CSV文件。
5. 导出数据组为文件，导出的都是.psd文件。
6. 创建批处理将PSD保存为JPG或PNG。
7. 使用批处理输出最终图像。

用我的Python脚本，你只需要：

1. 在电子表格中编辑内容。
2. 运行batch_export.py。

就这么简单，图就都出来了。你只需要有Python环境，装几个Python包。

## 下载

有两种下载方式：

1. `git pull`此仓库
2. 点击Code按钮，然后Download ZIP

## 设置

首次使用需要一些基本设置：

1. 将PSD模板放在此处，重命名为`{num}_template.psd`，默认num为1，如果有多个模板-数据同时工作，num可以往上加。在Python文件里切换。
2. 在此处创建`{num}_data.xlsx`文件，默认num为1，作为数据模板。
3. 编辑PSD模板文件，按以下模式重命名所有变量图层或组：
    - 格式：`@Variable_name#Operation_type`
    - 图层名称示例：`@badge#v`
    - `@`表示该图层是变量，脚本将从电子表格中获取内容。
    - `Variable_name`要对应电子表格中的表头。
    - `#Operation_type`表示脚本对图层做什么操作。
    - `#v`根据电子表格中的TRUE/FALSE设置可见性。
    - `#t`用电子表格数据替换文本图层内容。
    - `#t-c`或`#t-r`在替换时会使文本居中或右对齐。
    - `#i`用电子表格中的路径对应的图片填充图片图层。
4. 编辑XLSX文件，Python脚本默认读取第一张工作表，把你的数据放在这。也可以将数据放在另一张工作表中，并在第一张工作表中使用Excel公式读取和计算，特别适合切换图层可见性。
5. 将模板所需的其他文件放入`assets`文件夹，包括字体、背景图像等。

看起来非常复杂？相信我，用Photoshop做同样的事情要复杂得多。一旦设置完成，这将是你的救星。

## 导出

导出时，事情无比简单：

1. 在电子表格中粘贴内容。
2. 运行batch_export.py。

如果你熟悉Python，甚至可以改改代码，让脚本监控电子表格，并在电子表格修改后自动导出图像。

## 使用前提

```
pip install pillow pandas openpyxl psd-tools
```

## 感谢

特别感谢 [psd-tools](https://github.com/psd-tools/psd-tools) 提供强大的API，使我能够结合Photoshop的图像编辑能力，同时发挥Excel/Python在数据处理方面的优势。
