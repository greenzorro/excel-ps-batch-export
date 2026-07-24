# excel-ps-batch-export

[🇬🇧 EN](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README.md) | [🇨🇳 中文](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README_ZH_CN.md)

用 Python 读取 PSD 模板，把电子表格内容套进去批量出图，替代 Photoshop「图像 > 变量 > 定义」那套流程。

你负责 PSD 图层命名和表格数据，并验收成品图；本机的安装依赖与出图交给 Agent。

📺 示例：从 PSD 模板创建 Excel

https://github.com/user-attachments/assets/a21f8b2d-310f-4f28-a873-6bd166c07955

📺 示例：手动批量出图

https://github.com/user-attachments/assets/c52c6e05-1bc9-4a2b-ae4c-b283a25067f6

📺 示例：监控 Excel 自动出图

https://github.com/user-attachments/assets/bfd2d23f-84ec-4ea9-8874-523a298049be

在 Photoshop [你得这么干](https://victor42.eth.limo/post/3650/)：

1. 在电子表格里改内容。
2. 存成 CSV。
3. 在 Photoshop 里给图层定义变量。
4. 导入 CSV。
5. 导出数据组——结果却是一堆 `.psd`。
6. 再做批处理把 PSD 存成 JPG/PNG。
7. 跑批处理才得到最终图。

用本项目：改好表格（模板只需设一次），让 Agent 跑渲染即可。不用再走 Variables / 批处理那套。

## 模板怎么做

数据默认在 `demo/`，或已配置的自定义数据目录（`EPS_DATA_DIR`）。

1. 把 PSD 模板放进 `workspace/`。
2. 可变图层/组按 `@Variable_name#Operation_Parameter` 命名，例如 `@badge#v`、`@description#t_p`、`@bg#i`：
    - `@` 表示可变；`Variable_name` 必须对应表头列名
    - `#v` — 按 TRUE/FALSE 控制可见性
    - `#t` — 替换文字；修饰符：`_c` / `_r` 水平对齐，`_a[角度]` 旋转（PSD 里图层保持水平），`_p` 段落换行，`_pm` / `_pb` 段内垂直对齐（需配合 `_p`）。可组合，如 `#t_c_a15`、`#t_p_pm`。PSD 里的段落对齐 UI 无效，只认图层名
    - `#i` — 用表格里的图片路径填像素层；缩放 `_cover`（默认）/ `_contain`；九宫格对齐 `_lt` `_ct` `_rt` `_lm` `_cm`（默认）`_rm` `_lb` `_cb` `_rb`
    - 可变文字**不要**用自由变换（`Cmd/Ctrl+T`）拉尺寸，只用字号
    - 需要旋转时，只在图层名写 `#t_a…`，PSD 里保持水平
3. 让 Agent 跑 `xlsx_generator` 生成列头后，在第一张表填数（或用公式引用其它表）。保留 `File_name` 列（留空则默认 `image_1`…）。
4. 字体放 `workspace/assets/fonts/`，其它素材放 `workspace/assets/`。表里的图片路径相对 `workspace/`（如 `assets/1_img/image.jpg`）。
5. 可选 `workspace/fonts.json`：PSD 前缀 → 字体文件名。

看起来复杂？用 Photoshop Variables 更折磨。模板设好后，日常就是「贴行 → 让 Agent 出图」。

## 日常怎么用

- 在表格里粘贴或改行，然后让 Agent 出图（或让它用监控脚本盯着文件）。
- **剪贴板路径：** 从 Excel/网页复制表格 → 让 Agent 跑剪贴板导入 → 若有多个簿再选目标 → 出图。
- **多 PSD 共用一表：** 文件名第一个 `#` 前相同则共用一张表（如 `campaign#summer.psd` + `campaign#winter.psd` → `campaign.xlsx`）。每行给组内每个 PSD 各出一张图；`File_name` 为空时文件名会带上后缀。
- **变换规则：** 若存在 `workspace/前缀.json`，编辑 `前缀_raw.csv`，规则会写出可渲染的 `.xlsx`。类型：`direct`、`conditional`、`template`、`derived`、`derived_raw`。详见 `transform_guide.md`。

## 感谢

感谢 [psd-tools](https://github.com/psd-tools/psd-tools)：设计仍用 Photoshop，数据与出图交给 Excel/Python。

---

Created by [Victor42](https://victor42.work/) & [Agent Vik](https://github.com/agent-vik)
