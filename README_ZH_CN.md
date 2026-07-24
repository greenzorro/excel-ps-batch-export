# excel-ps-batch-export

[🇬🇧 EN](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README.md) | [🇨🇳 中文](https://github.com/greenzorro/excel-ps-batch-export/blob/main/README_ZH_CN.md)

用 Python 读取 PSD 模板，把电子表格内容套进去批量出图，替代 Photoshop「变量 → 定义」那套流程。

请让 Agent 按英文 README 的 `# For Agent` 在本机装好依赖并负责出图。你主要负责 **PSD 图层命名 + 表格数据**，再验收成品图。

📺 示例：从 PSD 模板创建 Excel

https://github.com/user-attachments/assets/a21f8b2d-310f-4f28-a873-6bd166c07955

📺 示例：手动批量出图

https://github.com/user-attachments/assets/c52c6e05-1bc9-4a2b-ae4c-b283a25067f6

📺 示例：监控 Excel 自动出图

https://github.com/user-attachments/assets/bfd2d23f-84ec-4ea9-8874-523a298049be

在 Photoshop [你得这么干](https://victor42.eth.limo/post/3650/)：编辑表格 → 存 CSV → 定义变量 → 导入 → 导出一堆 PSD → 再批处理成 JPG/PNG。

用本项目：改好表格（模板只需设一次），让 Agent 跑渲染即可。

## 模板怎么做（给你）

数据默认在 `demo/`，或 Agent 配好的自定义数据目录（`EPS_DATA_DIR`）。

1. PSD 放进 `workspace/`。
2. 可变图层/组按 `@变量名#操作_参数` 命名，如 `@badge#v`、`@description#t_p`、`@bg#i`：
    - `@` 表示可变；`变量名` 对应表头列
    - `#v` 可见性；`#t` 换文字（`_c`/`_r` 对齐，`_a角度` 旋转且 PSD 里保持水平，`_p` 段落，`_pm`/`_pb` 段内垂直对齐）
    - `#i` 填图（`_cover`/`_contain`，九宫格 `_lt`…`_rb`）
    - 可变文字**不要**用自由变换拉尺寸，只用字号；旋转只写在图层名里
3. 让 Agent 跑生成器出好列后，在第一张表填数（或用公式引用别的表）。保留 `File_name` 列，空则默认 `image_1`…
4. 字体放 `workspace/assets/fonts/`，其它素材放 `workspace/assets/`；表里的图片路径相对 `workspace/`。
5. 可选 `workspace/fonts.json`：PSD 前缀 → 字体文件名。

## 日常怎么用（给你）

- 改完表格行，让 Agent 出图（或让它开监控脚本）。
- **剪贴板**：复制表格 → 让 Agent 跑剪贴板导入 → 如有多个簿再选目标 → 出图。
- **多 PSD 共用一表**：文件名第一个 `#` 前相同则共用 `[前缀].xlsx`，一行出多张图。
- **变换规则**：有 `workspace/<前缀>.json` 时改 `<前缀>_raw.csv`，详见 `transform_guide.md`。

## 感谢

感谢 [psd-tools](https://github.com/psd-tools/psd-tools)。

---

Created by [Victor42](https://victor42.work/) & [Agent Vik](https://github.com/agent-vik)
