# -*- coding: utf-8 -*-
"""generate_resume.py 用的自定义异常。

与 SKILL.md §13 对齐:
  - ColorParseError    : 配色解析失败 (parse_color)
  - DataValidationError: 简历必填数据缺失 (validate_resume)
  - TemplateNotFoundError: 模板不存在 (预留)
"""


class ColorParseError(ValueError):
    """配色解析失败（无效预设名或十六进制）。"""


class DataValidationError(ValueError):
    """简历数据结构校验失败（缺失必填字段）。"""


class TemplateNotFoundError(ValueError):
    """指定的模板不存在。"""
