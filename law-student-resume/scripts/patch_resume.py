#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
patch_resume.py - 交互式微调简历 JSON

支持路径语法对 JSON 做定向修改，修改后重新生成即可，无需重填整个流程。

用法示例：
  python scripts/patch_resume.py \
    --base resume.json \
    --set "personal.email=new_email@example.com" \
    --set "internship[0].bullets[1]=将第1段实习的第2条bullet改为新文本" \
    --out resume_patched.json

  python scripts/patch_resume.py \
    --base resume.json \
    --add "internship[0].bullets=\"新补充的第4条bullet\"" \
    --out resume_patched.json

  python scripts/patch_resume.py \
    --base resume.json \
    --del "internship[1]" \
    --out resume_patched.json

  python scripts/patch_resume.py \
    --base resume.json \
    --patch resume_fragment.json \
    --out resume_patched.json
"""

import argparse
import json
import re
import sys


def _parse_path(path):
    """将点分路径解析为 token 列表。支持数组索引，如 internship[0].bullets[1]。"""
    tokens = []
    # 先按 [] 拆分
    parts = re.split(r"\.|\[(\d+)\]", path)
    for p in parts:
        if p is None or p == "":
            continue
        if p.isdigit():
            tokens.append(int(p))
        else:
            tokens.append(p)
    return tokens


def _get_parent(data, tokens):
    """沿着路径 tokens 走到父节点，返回 (父容器, 最后 token)。"""
    node = data
    for tok in tokens[:-1]:
        if isinstance(node, dict):
            node = node[tok]
        elif isinstance(node, list):
            node = node[tok]
        else:
            raise TypeError(f"路径 {tokens} 无法继续深入，遇到非容器类型")
    return node, tokens[-1]


def _set_value(data, path, raw_value):
    """解析 raw_value 为 JSON 值后设置到路径。"""
    try:
        value = json.loads(raw_value)
    except Exception:
        value = raw_value
    tokens = _parse_path(path)
    node, last = _get_parent(data, tokens)
    if isinstance(node, dict):
        node[last] = value
    elif isinstance(node, list):
        if not isinstance(last, int):
            raise ValueError(f"数组索引必须是整数：{last}")
        node[last] = value
    else:
        raise TypeError(f"无法设置到非容器：{path}")


def _add_value(data, path, raw_value):
    """在数组路径末尾追加元素。path 指向待追加的列表本身，如 internship[0].bullets。"""
    try:
        value = json.loads(raw_value)
    except Exception:
        value = raw_value
    tokens = _parse_path(path)
    if not tokens:
        raise ValueError(f"空路径: {path}")
    # 走到倒数第二层，last 为列表的 key
    node, last = _get_parent(data, tokens)
    if isinstance(node, dict):
        arr = node[last]
    elif isinstance(node, list):
        arr = node[last]
    else:
        raise TypeError(f"无法索引到非容器：{path}")
    if not isinstance(arr, list):
        raise TypeError(f"--add 目标必须是列表：{path}")
    arr.append(value)


def _del_value(data, path):
    """删除路径对应的节点。"""
    tokens = _parse_path(path)
    node, last = _get_parent(data, tokens)
    if isinstance(node, dict):
        del node[last]
    elif isinstance(node, list):
        if not isinstance(last, int):
            raise ValueError(f"数组索引必须是整数：{last}")
        del node[last]
    else:
        raise TypeError(f"无法从非容器删除：{path}")


def _apply_json_patch(data, patch_path):
    """将 patch JSON 文件合并到 base（同层 key 覆盖，深层不做递归合并）。"""
    with open(patch_path, "r", encoding="utf-8") as f:
        patch = json.load(f)
    if not isinstance(patch, dict):
        raise ValueError("--patch 文件根必须是对象")
    for k, v in patch.items():
        data[k] = v


def main():
    ap = argparse.ArgumentParser(description="交互式微调简历 JSON")
    ap.add_argument("--base", required=True, help="原始简历 JSON 路径")
    ap.add_argument("--set", action="append", default=[],
                    help="设置路径值，如 --set personal.email=xxx")
    ap.add_argument("--add", action="append", default=[],
                    help="在列表追加，如 --add internship[0].bullets=\"新bullet\"")
    ap.add_argument("--del", dest="delete", action="append", default=[],
                    help="删除节点，如 --del internship[1]")
    ap.add_argument("--patch", default=None, help="合并一个 JSON 片段文件")
    ap.add_argument("--out", required=True, help="输出 JSON 路径")
    args = ap.parse_args()

    with open(args.base, "r", encoding="utf-8") as f:
        data = json.load(f)

    for s in args.set:
        if "=" not in s:
            print(f"[error] --set 参数缺少 '='：{s}", file=sys.stderr)
            sys.exit(1)
        k, v = s.split("=", 1)
        _set_value(data, k, v)

    for a in args.add:
        if "=" not in a:
            print(f"[error] --add 参数缺少 '='：{a}", file=sys.stderr)
            sys.exit(1)
        k, v = a.split("=", 1)
        _add_value(data, k, v)

    for d in args.delete:
        _del_value(data, d)

    if args.patch:
        _apply_json_patch(data, args.patch)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅ 已保存微调后的 JSON：{args.out}")
    print("  下一步：用 generate_resume.py --data 此文件重新生成即可")


if __name__ == "__main__":
    main()
