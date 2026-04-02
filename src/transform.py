"""
File: transform.py
Project: excel-ps-batch-export
Author: Victor Cheng
Email: hi@victor42.work
Description: 数据变换引擎 — 读取原始CSV + JSON规则 → 计算成品数据 → 写入xlsx

将Excel公式中的业务逻辑迁移到Python，脱离GUI依赖。
支持5种变换类型：direct, conditional, template, derived, derived_raw
"""

import os
import sys
import json
import re
import pandas as pd
from typing import Any


def load_rules(template: str) -> dict:
    """加载JSON规则文件

    :param str template: 模板前缀（如 "1", "2", "3"）
    :return dict: 规则字典
    :raises FileNotFoundError: 规则文件不存在
    """
    json_path = os.path.join("../workspace", f"{template}.json")
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"规则文件不存在: {json_path}")
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_raw_data(template: str) -> pd.DataFrame:
    """读取原始CSV数据

    :param str template: 模板前缀
    :return pd.DataFrame: 原始数据
    :raises FileNotFoundError: CSV文件不存在
    """
    csv_path = os.path.join("../workspace", f"{template}_raw.csv")
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"原始数据文件不存在: {csv_path}")
    return pd.read_csv(csv_path, dtype=str, keep_default_na=False)



def remove_spaces(value: Any) -> str:
    """移除字符串中的所有空格

    :param value: 输入值
    :return str: 去空格后的字符串
    """
    if value is None:
        return ""
    return str(value).replace(" ", "")


def is_empty(value: Any) -> bool:
    """判断值是否为空（None、空字符串、纯空格、"NULL"）

    :param value: 输入值
    :return bool: 是否为空
    """
    if value is None:
        return True
    s = str(value).strip()
    return s == "" or s == "NULL"


def apply_direct(rule: dict, raw_row: pd.Series) -> str:
    """模式1: 直接映射 — 从raw字段直接复制，可选去空格

    :param dict rule: 列规则
    :param pd.Series raw_row: 原始数据行
    :return str: 映射后的值
    """
    source = rule.get("source", "")
    value = raw_row.get(source, "")
    if rule.get("remove_spaces", False):
        value = remove_spaces(value)
    return "" if is_empty(value) else str(value)


def apply_conditional(rule: dict, raw_row: pd.Series, product_row: dict) -> str:
    """模式2: 条件映射 — 仅在depends_on字段非空时才复制

    :param dict rule: 列规则
    :param pd.Series raw_row: 原始数据行
    :param dict product_row: 已计算的成品数据（用于检查depends_on）
    :return str: 映射后的值，或空字符串
    """
    depends_on = rule.get("depends_on", "")
    guard_value = product_row.get(depends_on, "")
    if is_empty(guard_value):
        return ""

    source = rule.get("source", "")
    raw_value = raw_row.get(source, "")
    if is_empty(raw_value) or str(raw_value).strip() == " ":
        return ""
    if rule.get("remove_spaces", False):
        raw_value = remove_spaces(raw_value)
    return str(raw_value)


def apply_template(rule: dict, raw_row: pd.Series, row_index: int) -> str:
    """模式3: 字符串拼接 — 将多个字段拼接，可选去空格、跳过空字段

    :param dict rule: 列规则
    :param pd.Series raw_row: 原始数据行
    :param int row_index: 行号（1-based，用于 {_row}）
    :return str: 拼接后的字符串
    """
    template = rule.get("template", "")
    remove_spaces_fields = rule.get("remove_spaces", [])
    skip_if_empty = rule.get("skip_if_empty", [])

    # 收集所有需要替换的字段名
    fields = re.findall(r"\{(\w+)\}", template)
    # 也匹配中文字段名
    fields_cn = re.findall(r"\{([^}]+)\}", template)
    all_fields = list(set(fields + fields_cn))

    values = {}
    for field in all_fields:
        if field == "_row":
            values[field] = str(row_index)
        else:
            values[field] = str(raw_row.get(field, ""))

    # 移除指定字段的空格
    for field in remove_spaces_fields:
        if field in values:
            values[field] = remove_spaces(values[field])

    # 跳过空字段（从模板中移除对应的 -{field} 部分）
    result = template
    for field in skip_if_empty:
        if is_empty(values.get(field, "")):
            # 移除 -{field} 及其前面的连字符
            result = result.replace(f"-{{{field}}}", "")
            # 如果字段在末尾没有连字符前缀
            result = result.replace(f"{{{field}}}", "")

    # 执行模板替换
    for field, value in values.items():
        if is_empty(value):
            result = result.replace(f"{{{field}}}", "")
        else:
            result = result.replace(f"{{{field}}}", value)

    # 清理多余的连续连字符
    result = re.sub(r"-{2,}", "-", result)
    # 清理首尾的连字符
    result = result.strip("-")

    return result


def apply_derived(rule: dict, product_row: dict) -> str:
    """模式4: 布尔推导（基于成品字段）— 检查成品字段是否为空

    :param dict rule: 列规则
    :param dict product_row: 已计算的成品数据
    :return str: "TRUE" 或 "FALSE"
    """
    field = rule.get("field", "")
    value = product_row.get(field, "")
    when_empty = rule.get("when_empty", True)

    if when_empty:
        return "True" if is_empty(value) else "False"
    else:
        return "True" if not is_empty(value) else "False"


def apply_derived_raw(rule: dict, raw_row: pd.Series) -> str:
    """模式5: 布尔推导（基于原始字段）— 直接检查raw字段是否有值

    :param dict rule: 列规则
    :param pd.Series raw_row: 原始数据行
    :return str: "TRUE" 或 "FALSE"
    """
    source = rule.get("source", "")
    value = raw_row.get(source, "")
    return "True" if not is_empty(value) else "False"


def transform_row(raw_row: pd.Series, rules: dict, row_index: int) -> dict:
    """对单行原始数据执行全部变换规则

    :param pd.Series raw_row: 原始数据行
    :param dict rules: JSON规则
    :param int row_index: 行号（1-based）
    :return dict: 成品数据行
    """
    primary_field = rules.get("primary_field", "")
    columns = rules.get("columns", {})
    product_row = {}

    # 行级守卫：主字段为空则跳过整行
    if primary_field and is_empty(raw_row.get(primary_field, "")):
        return {}

    # 按顺序处理每列（顺序很重要，derived依赖前面的成品字段）
    for col_name, rule in columns.items():
        rule_type = rule.get("type", "")

        if rule_type == "direct":
            product_row[col_name] = apply_direct(rule, raw_row)

        elif rule_type == "conditional":
            product_row[col_name] = apply_conditional(rule, raw_row, product_row)

        elif rule_type == "template":
            product_row[col_name] = apply_template(rule, raw_row, row_index)

        elif rule_type == "derived":
            product_row[col_name] = apply_derived(rule, product_row)

        elif rule_type == "derived_raw":
            product_row[col_name] = apply_derived_raw(rule, raw_row)

        else:
            product_row[col_name] = ""

    return product_row


def transform(template: str) -> int:
    """执行完整的数据变换流程

    :param str template: 模板前缀
    :return int: 生成的数据行数
    """
    rules = load_rules(template)
    raw_df = load_raw_data(template)
    columns = rules.get("columns", {})
    output_columns = list(columns.keys())

    product_rows = []
    for idx, raw_row in raw_df.iterrows():
        product_row = transform_row(raw_row, rules, idx + 1)
        if product_row:  # 跳过空行（主字段为空）
            product_rows.append(product_row)

    product_df = pd.DataFrame(product_rows, columns=output_columns)

    xlsx_path = os.path.join("../workspace", f"{template}.xlsx")

    # 始终使用JSON配置中定义的列顺序，不读取现有xlsx
    product_df.to_excel(xlsx_path, index=False, sheet_name="Sheet1")
    print(f"已写入 {len(product_rows)} 行数据到 {xlsx_path}")
    return len(product_rows)


if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    if len(sys.argv) < 2:
        print("用法: python transform.py <模板前缀>")
        print("示例: python transform.py 1")
        sys.exit(1)

    template_name = sys.argv[1]
    count = transform(template_name)
    print(f"完成！共处理 {count} 行数据")
