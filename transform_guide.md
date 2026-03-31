# 数据变换规则系统

## 概述

对于需要复杂数据处理的模板，系统提供了基于规则的数据变换功能。

## 工作流程

当 workspace 目录中存在 `.json` 规则文件时，数据流程如下：

1. 用户/agent 编辑 `_raw.csv` 原始数据
2. `transform.py` 读取 `.json` 规则和 `_raw.csv`
3. 按规则计算成品数据，写入 `.xlsx`
4. `psd_renderer.py` 读取 `.xlsx` 进行渲染

## 5 种变换类型

### direct — 直接映射

从原始字段直接复制，可选去空格。

```json
"分类": { "type": "direct", "source": "分类" },
"标题": { "type": "direct", "source": "标题", "remove_spaces": true }
```

### conditional — 条件映射

仅当主字段非空时才复制，否则整行为空。用于依赖"上一级"存在的字段。

```json
"标题2": {
  "type": "conditional",
  "source": "标题2",
  "depends_on": "标题1",
  "remove_spaces": true
}
```

### template — 字符串拼接

将多个字段拼接为新的字符串，可选去空格、跳过空字段。

**内置变量：**
- `{_row}` - 当前行号（从1开始自动递增，无需在原始数据中提供）

```json
"File_name": {
  "type": "template",
  "template": "{_row}-{分类}-{标题1}-{标题2}",
  "remove_spaces": ["分类"],
  "skip_if_empty": ["标题2"]
}
```

**注意：** 原始数据中不需要包含"顺序"列，使用`{_row}`即可自动生成行号。

### derived — 布尔推导（基于成品字段）

根据另一个成品字段是否为空推导布尔值。

```json
"单行": { "type": "derived", "field": "标题2", "when_empty": true },
"两行": { "type": "derived", "field": "标题2", "when_empty": false },
"时间框": { "type": "derived", "field": "直播时间", "when_empty": false }
```

### derived_raw — 布尔推导（基于原始字段）

直接检查原始字段是否有值，不经过 conditional 过滤。

```json
"站内标": { "type": "derived_raw", "source": "站内标" },
"站外标": { "type": "derived_raw", "source": "站外标" }
```

## JSON 配置文件示例

```json
{
  "_comment": "模板数据变换规则",
  "primary_field": "标题1",
  "columns": {
    "File_name": {
      "type": "template",
      "template": "{_row}-{分类}-{标题1}-{标题2}",
      "remove_spaces": ["分类"],
      "skip_if_empty": ["标题2"]
    },
    "分类": { "type": "direct", "source": "分类" },
    "标题1": {
      "type": "direct",
      "source": "标题1",
      "remove_spaces": true
    },
    "标题2": {
      "type": "conditional",
      "source": "标题2",
      "depends_on": "标题1",
      "remove_spaces": true
    },
    "直播时间": {
      "type": "conditional",
      "source": "直播时间",
      "depends_on": "标题1"
    },
    "小标签内容": {
      "type": "conditional",
      "source": "小标签内容",
      "depends_on": "标题1",
      "remove_spaces": true
    },
    "单行": {
      "type": "derived",
      "field": "标题2",
      "when_empty": true
    },
    "两行": {
      "type": "derived",
      "field": "标题2",
      "when_empty": false
    },
    "时间框": {
      "type": "derived",
      "field": "直播时间",
      "when_empty": false
    },
    "小标签": {
      "type": "derived",
      "field": "小标签内容",
      "when_empty": false
    },
    "站内标": {
      "type": "derived_raw",
      "source": "站内标"
    },
    "站外标": {
      "type": "derived_raw",
      "source": "站外标"
    },
    "背景图": {
      "type": "template",
      "template": "assets/1_img/{分类}.jpg",
      "remove_spaces": ["分类"]
    }
  }
}
```

## 模板对照表

| 特征 | 模板 1 | 模板 2 | 模板 3 | 模板 4 | 模板 5 |
|------|--------|--------|--------|--------|--------|
| direct | ✅ | ✅ | ✅ | ❌ | ❌ |
| conditional | ✅ | ✅ | ✅ | ❌ | ❌ |
| template | ✅ | ✅ | ✅ | ❌ | ❌ |
| derived | ✅ | ✅ | ✅ | ❌ | ❌ |
| derived_raw | ✅ | ✅ | ✅ | ❌ | ❌ |
| 需要规则文件 | 是 | 是 | 是 | 否 | 否 |

模板 4 和 5 不需要规则文件，用户直接编辑 xlsx 即可。
