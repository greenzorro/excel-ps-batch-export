# excel-ps-batch-export

[🇬🇧 EN](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README.md) | [🇨🇳 中文](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README_ZH_CN.md)

这是一个Python脚本，用于读取PSD模板并将电子表格中的内容应用于模板，导出大量图像，替代Photoshop的变量定义功能。

📺 示例：从PSD模板创建Excel文件

https://github.com/user-attachments/assets/a21f8b2d-310f-4f28-a873-6bd166c07955

📺 示例：手动批量出图

https://github.com/user-attachments/assets/c52c6e05-1bc9-4a2b-ae4c-b283a25067f6

📺 示例：监控Excel文件变化自动出图

https://github.com/user-attachments/assets/bfd2d23f-84ec-4ea9-8874-523a298049be

在Photoshop[你得这么干](https://victor42.eth.limo/post/3650/)：

1. 在电子表格中编辑内容。
2. 将电子表格保存为CSV文件。
3. 在Photoshop中定义图层变量。
4. 导入CSV文件。
5. 导出数据组为文件，导出的都是.psd文件。
6. 创建批处理将PSD保存为JPG或PNG。
7. 使用批处理输出最终图像。

用我的Python脚本，你只需要：

1. 在电子表格中编辑内容。
2. 运行psd_renderer.py。

就这么简单，图就都出来了。你只需要有Python环境，装几个Python包。

## 下载

有两种下载方式：

1. `git clone https://github.com/greenzorro/excel-ps-batch-export.git`
2. 点击Code按钮，然后Download ZIP

## 设置

首次使用需要一些基本设置：

1. 将PSD模板文件放在此处。
2. 编辑PSD模板文件，按以下模式重命名所有可变图层或组：
    - 格式：`@Variable_name#Operation_Parameter`
    - 图层名称示例：`@badge#v`、`@description#t_p`、`@bg#i`
    - `@`表示该图层是可变的，脚本将从电子表格中获取内容。
    - `Variable_name`应出现在电子表格的列标题中。
    - `#Operation_Parameter`告诉脚本对图层做什么操作。
    - `#v`根据电子表格中的TRUE/FALSE设置可见性。
    - `#t` 用电子表格数据替换文本图层内容，参数包括：
        - 默认左上对齐
        - `_c` 水平居中对齐
        - `_r` 水平右对齐
        - `_p` 段落文本换行，在PSD中至少填满一行文本
        - `_pm` 垂直居中对齐
        - `_pb` 垂直底部对齐
        - 这些参数可以组合使用，如 `#t_c_p`、`#t_r_pb`
        - PSD里设置的对齐方向对结果没有影响，程序只认图层名
    - `#i`用电子表格中的路径对应的图片填充图片图层。
    - 注意：请勿使用cmd/ctrl+T缩放可变文本图层。只能通过字体大小属性调整其尺寸，否则脚本将从PSD文件中获取错误的字号。如果已经这么做了，请创建新的文本图层替换它们。
3. 运行`xlsx_generator.py`。你的XLSX文件就创建完成了，所有的列都已准备好。
4. 编辑XLSX文件。Python脚本默认读取第一张工作表，把你的数据放在这。也可以将数据放在另一张工作表中，并在第一张工作表中使用Excel公式读取和计算，特别适合切换图层可见性。请勿删除第一列`File_name`，留空会使用默认文件名格式（如image_1, image_2等）。
5. 将模板所需的其他文件放入`assets`文件夹，包括字体、背景图像等。确保图片资源的路径与电子表格中的数据匹配。

看起来非常复杂？相信我，用Photoshop做同样的事情要复杂得多。一旦设置完成，这将是你的救星。

## 导出

导出时，事情无比简单：

1. 在电子表格中粘贴内容。
2. 运行psd_renderer.py。

我甚至写了另一个脚本（file_monitor.py）监控电子表格，并在电子表格修改后自动导出图像。

## 多文件处理

本工具支持用一个Excel文件处理多个PSD模板。工作原理如下：

- **按前缀分组**：同一目录中的所有PSD文件按前缀分组。前缀定义为文件名中第一个井号（`#`）之前的部分。例如：
  - `产品介绍#模板A.psd` 和 `产品介绍#模板B.psd` 共享相同的前缀 `产品介绍`
- **共享Excel**：每组创建一个Excel文件（命名为`[前缀].xlsx`），包含组内所有PSD的变量。
- **批量导出**：运行 `psd_renderer.py [前缀] ...` 时，脚本将处理组内所有PSD。Excel中的每一行将为组内每个PSD生成一张图片。输出图片文件名包含PSD的后缀（如 `行1_模板A.jpg`）。

示例：
  - PSD文件：`活动#夏季版.psd`, `活动#冬季版.psd`
  - Excel文件：`活动.xlsx`
  - 命令：`python psd_renderer.py 活动 AlibabaPuHuiTi-2-85-Bold.ttf jpg`
  - 输出：对于 `活动.xlsx` 中的每一行，生成两张图片：`行1_夏季版.jpg`, `行1_冬季版.jpg` 等。

## 使用前提

### 安装依赖

推荐使用 `requirements.txt` 文件安装所有依赖：

```bash
pip install pillow pandas openpyxl psd-tools tqdm
```

## 使用说明

### 基本用法

```bash
# 基本命令格式
python batch_export.py [Excel文件前缀] [字体文件] [输出格式]

# 示例
python batch_export.py 1 AlibabaPuHuiTi-2-85-Bold.ttf jpg
```

## 感谢

特别感谢 [psd-tools](https://github.com/psd-tools/psd-tools) 提供强大的API，使我能够结合Photoshop的图像编辑能力，同时发挥Excel/Python在数据处理方面的优势。

---

由 [Victor_42](https://victor42.work/) 创造
