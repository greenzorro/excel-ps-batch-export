# Excel-PS 批量导出工具 - 字体配置改进计划

## 1. 背景与问题

### 1.1 当前设计
- **字体来源**：通过命令行第二个参数全局指定
  ```bash
  python psd_renderer.py [模板名] [字体文件] [格式]
  #                          ↑ 这里
  ```

- **代码位置**：`psd_renderer.py:873`
  ```python
  font_file = sys.argv[2]  # 从命令行参数获取字体文件
  text_font = font_file     # 赋值给全局变量
  ```

### 1.2 存在的问题

| 问题 | 说明 |
|------|------|
| **字体和模板分离** | 字体在命令行全局指定，与 PSD 模板无关 |
| **无法区分模板** | 多个 PSD 模板（如 `1#海报.psd`、`1#方图.psd`）必须共用同一个字体 |
| **无法区分图层** | 同一个 PSD 里不同文字图层（标题/正文/注释）也无法使用不同字体 |
| **不符合实际需求** | 用户反馈："字体应该是和每个 PSD 绑定的" |

### 1.3 用户需求
> "我不需要每个图层有不同的字体，我只需要每个 PSD 有自己配置的字体就行了。"

**结论**：每个 PSD 模板前缀配置一个字体即可。

---

## 2. 解决方案

### 2.1 选定方案：JSON 配置文件

**配置文件**：`fonts.json`
```json
{
  "_comment": "字体配置文件 - 为每个PSD模板指定对应的字体文件",
  "_usage": "键名为PSD文件前缀（不含扩展名和#后缀），值为字体文件路径",
  "_path_rules": "路径规则：1) 相对路径相对于 assets/fonts/ 目录，如 'AlibabaPuHuiTi-2-85-Bold.ttf'  2) 绝对路径也可以使用，如 'C:/Windows/Fonts/simhei.ttf'",
  "1": "AlibabaPuHuiTi-2-85-Bold.ttf",
  "2": "SourceHanSans-Bold.ttf",
  "3": "NotoSansSC-Regular.ttf"
}
```

### 2.2 方案选择对比

| 方案 | 优点 | 缺点 | 是否采用 |
|------|------|------|----------|
| JSON 配置文件 | 集中管理、修改方便、无需额外依赖 | 需要维护额外文件 | ✅ 采用 |
| PSD 文件名嵌入 | 无需额外文件 | 文件名冗长、修改需重命名 | ❌ |
| Excel 配置 sheet | 与数据在一起 | Excel 结构变复杂、易被误删 | ❌ |
| 环境变量 | 适合 CI/CD | 跨项目困难、用户不友好 | ❌ |

---

## 3. 实施计划

### 3.1 任务一：实现 JSON 字体配置功能

#### 3.1.1 已完成
- [x] 创建 `fonts.json` 配置文件模板

#### 3.1.2 待完成

**A. 修改导入和全局变量**
```python
# psd_renderer.py:10-20 附近
import json
from typing import Optional

# 新增全局变量
fonts_config: Dict[str, str] = {}
DEFAULT_FONT = 'assets/fonts/AlibabaPuHuiTi-2-85-Bold.ttf'
```

**B. 添加配置加载函数**
```python
def load_fonts_config():
    """加载字体配置文件

    :return dict: 字体配置字典 {psd_prefix: font_path}
    """
    global fonts_config
    config_path = 'fonts.json'

    if not os.path.exists(config_path):
        safe_print_message(f"警告：字体配置文件不存在: {config_path}")
        return {}

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        # 过滤掉注释字段（以 _ 开头的键）
        fonts_config = {k: v for k, v in config.items() if not k.startswith('_')}
        safe_print_message(f"已加载字体配置: {len(fonts_config)} 个PSD模板")
        return fonts_config
    except json.JSONDecodeError as e:
        safe_print_message(f"错误：字体配置文件格式错误: {e}")
        return {}
    except Exception as e:
        safe_print_message(f"错误：加载字体配置失败: {e}")
        return {}
```

**C. 添加 PSD 前缀提取函数**
```python
def get_psd_prefix(psd_file_name: str) -> str:
    """从 PSD 文件名提取前缀

    :param str psd_file_name: PSD 文件名，如 "1#海报.psd" 或 "2.psd"
    :return str: PSD 前缀，如 "1" 或 "2"

    Examples:
        >>> get_psd_prefix("1#海报.psd")
        "1"
        >>> get_psd_prefix("2.psd")
        "2"
        >>> get_psd_prefix("产品#横版#v2.psd")
        "产品"
    """
    # 去掉扩展名
    name_without_ext = os.path.splitext(psd_file_name)[0]

    # 提取第一个 # 之前的部分作为前缀
    if '#' in name_without_ext:
        prefix = name_without_ext.split('#', 1)[0]
    else:
        prefix = name_without_ext

    return prefix
```

**D. 添加字体获取函数**
```python
def get_font_for_psd(psd_file_name: str) -> str:
    """根据 PSD 文件名获取对应的字体文件路径

    :param str psd_file_name: PSD 文件名
    :return str: 字体文件路径

    优先级：
    1. fonts.json 中配置的字体
    2. 默认字体 DEFAULT_FONT
    """
    global fonts_config

    # 提取 PSD 前缀
    psd_prefix = get_psd_prefix(psd_file_name)

    # 从配置中查找字体
    if psd_prefix in fonts_config:
        font_path = fonts_config[psd_prefix]
        safe_print_message(f"  [{psd_prefix}] 使用字体: {font_path}")
        return font_path

    # 未找到配置，使用默认字体
    safe_print_message(f"  [{psd_prefix}] 未配置字体，使用默认字体: {DEFAULT_FONT}")
    return DEFAULT_FONT
```

**E. 修改 `export_single_image` 函数**
```python
# psd_renderer.py:469 附近
# 添加字体参数
def export_single_image(row, index, psd_object, psd_file_name, font=None):
    """处理单行数据并导出图像（单进程串行版本）

    :param pd.Series row: 包含单行数据的Series
    :param int index: 当前行索引
    :param PSDImage psd_object: 预加载的PSD对象
    :param str psd_file_name: PSD文件名（用于输出文件名）
    :param str font: 字体文件路径（可选）
    """
    # ... 原有代码 ...

    # 修改 update_text_layer 调用，传入字体
    elif operation_type.startswith('t'):
        update_text_layer(layer, str(row[field_name]), pil_image, font)
```

**F. 修改 `psd_renderer_images` 函数**
```python
# psd_renderer.py:753 附近
def psd_renderer_images():
    """批量输出图片
    """
    # 在开始时加载字体配置
    load_fonts_config()

    # ... 原有代码 ...

    # 修改循环，为每个 PSD 获取对应字体
    for psd_file in matching_psds:
        if psd_objects[psd_file] is not None:
            # 获取当前 PSD 的字体
            psd_font = get_font_for_psd(psd_file)

            for index, row in df.iterrows():
                try:
                    # 传递字体参数
                    export_single_image(row, index, psd_objects[psd_file], psd_file, psd_font)
```

### 3.2 任务二：删除命令行字体参数

#### 3.2.1 待完成

**A. 修改命令行参数说明**
```python
# psd_renderer.py:867 附近
if __name__ == "__main__":
    # 修改参数检查
    if len(sys.argv) < 3:
        print("用法: python psd_renderer.py [模板名] [输出格式]")
        print("示例: python psd_renderer.py 1 jpg")
        print("\n说明：字体配置请使用 fonts.json 文件")
        sys.exit(1)

    file_name = sys.argv[1]  # Excel/PSD 文件前缀
    image_format = sys.argv[2]  # 输出图片格式

    # 删除 font_file = sys.argv[2]
```

**B. 删除全局变量**
```python
# psd_renderer.py:38-48 附近
# 删除以下变量
# font_file = None
# text_font = None
```

**C. 删除主函数中的字体赋值**
```python
# psd_renderer.py:872-883 附近
# 删除
# font_file = sys.argv[2]
# text_font = font_file
```

**D. 更新帮助信息和错误提示**
```python
# 确保所有提到字体参数的地方都更新
```


### 3.3 任务三：修复兼容性问题

#### 3.3.1 待完成

**A. 修复 clipboard_importer.py 兼容性**

问题：clipboard_importer.py 使用旧的命令行参数格式调用 psd_renderer.py

修改清单：
- 删除 PREFERRED_FONT 配置
- 删除 FONTS_DIR 配置
- 简化 get_rendering_config() 只返回格式
- 修改 run_psd_renderer() 删除字体参数传递

**B. 修复 file_monitor.py 兼容性**

问题：file_monitor.py 硬编码了字体参数

修改清单：
- 删除 font_file 变量
- subprocess 调用中删除字体参数

**C. 检查并更新测试文件**

检查 tests/ 目录下的测试文件，确保函数调用正确。

---

## 4. 代码修改清单

| 文件 | 修改行数 | 说明 |
|------|----------|------|
| `fonts.json` | 新建 | 字体配置文件（已完成） |
| `psd_renderer.py` | 导入部分 | 添加 json, Optional |
| `psd_renderer.py` | 全局变量 | 添加 fonts_config, DEFAULT_FONT |
| `psd_renderer.py` | 函数区域 | 添加 load_fonts_config() |
| `psd_renderer.py` | 函数区域 | 添加 get_psd_prefix() |
| `psd_renderer.py` | 函数区域 | 添加 get_font_for_psd()（增强路径处理） |
| `psd_renderer.py` | export_single_image() | 添加 font 参数及默认值处理 |
| `clipboard_importer.py` | 配置部分 | 删除 PREFERRED_FONT, FONTS_DIR |
| `clipboard_importer.py` | 函数部分 | 简化 get_rendering_config() |
| `clipboard_importer.py` | run_psd_renderer() | 删除字体参数传递 |
| `file_monitor.py` | 配置部分 | 删除 font_file 变量 |
| `file_monitor.py` | subprocess 调用 | 删除字体参数 |
| `tests/*.py` | 待检查 | 更新函数调用（如需要） |
| `psd_renderer.py` | psd_renderer_images() | 调用 load_fonts_config() |
| `psd_renderer.py` | psd_renderer_images() | 调用 get_font_for_psd() |
| `psd_renderer.py` | __main__ | 删除字体参数处理 |
| `notes.md` | 待更新 | 更新文档说明 |

---

## 5. 向后兼容性

| 场景 | 行为 |
|------|------|
| fonts.json 不存在 | 使用默认字体 `AlibabaPuHuiTi-2-85-Bold.ttf`，显示警告 |
| PSD 前缀未配置 | 使用默认字体，显示警告 |
| 旧命令行调用 | 参数减少一个，需要更新调用方式 |
| 相对路径配置 | 自动相对于 assets/fonts/ 目录解析 |
| 绝对路径配置 | 直接使用 |

---

## 6. 测试计划

### 6.1 功能测试
- [ ] fonts.json 不存在时，使用默认字体
- [ ] fonts.json 存在但 PSD 未配置时，使用默认字体
- [ ] fonts.json 正确配置时，使用指定字体
- [ ] 多个 PSD 使用不同字体

### 6.2 兼容性测试
- [ ] Windows 路径
- [ ] Linux/macOS 路径
- [ ] clipboard_importer.py 自动调用
- [ ] file_monitor.py 自动调用

### 6.3 回归测试
- [ ] 运行完整测试套件

---

## 7. 更新文档

修改 `notes.md` 中的相关内容：
- 删除命令行字体参数说明
- 添加 fonts.json 配置说明
- 更新使用示例

---

## 8. 实施状态

| 任务 | 状态 |
|------|------|
| 创建 fonts.json | 🔄 进行中 |
| 实现 JSON 字体配置功能 | ⏳ 待开始 |
| 删除命令行字体参数 | ⏳ 待开始 |
| 修复 clipboard_importer.py 兼容性 | ⏳ 待开始 |
| 修复 file_monitor.py 兼容性 | ⏳ 待开始 |
| 更新 notes.md 文档 | ⏳ 待开始 |
| 测试 | ⏳ 待开始 |

---

## 9. 注意事项

1. **字体路径处理**：
   - 相对路径自动相对于 assets/fonts/ 目录解析
   - 绝对路径直接使用
   - 在 get_font_for_psd() 中统一处理

2. **错误提示**：
   - 当字体配置文件不存在时，显示警告并使用默认字体
   - 当 PSD 前缀未配置时，显示警告并使用默认字体
   - 当字体文件不存在时，需要给出清晰的错误提示

3. **编码问题**：
   - JSON 文件使用 UTF-8 编码
   - 确保中文路径正常

4. **向后兼容**：
   - 保持 update_text_layer 函数的默认参数
   - export_single_image() 的 font 参数有默认值处理

5. **依赖模块更新**：
   - clipboard_importer.py 和 file_monitor.py 通过 subprocess 调用 psd_renderer.py
   - 必须同步删除字体参数，否则调用会失败
