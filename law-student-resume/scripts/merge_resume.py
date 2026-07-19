#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
断点续传：合并「已保存的简历 JSON 片段」与「新补充的 JSON」，输出合并结果。

用法：
  python merge_resume.py --base saved.json --patch new.json --out merged.json

合并规则：
  - 标量字段（name / phone / gpa ...）：patch 覆盖 base。
  - 数组字段（education / internship / research / ...）：保留 base，
    追加 patch 中 base 没有的项（按 JSON 内容去重），实现“续填不死覆盖”。
  - meta 字段同样按上述规则合并。
"""
import argparse
import json


def _key(obj):
    return json.dumps(obj, ensure_ascii=False, sort_keys=True)


def deep_merge(base, patch):
    out = dict(base)
    for k, v in patch.items():
        if isinstance(v, list) and isinstance(out.get(k), list):
            seen = {_key(x) for x in out[k]}
            for item in v:
                kk = _key(item)
                if kk not in seen:
                    out[k].append(item)
                    seen.add(kk)
        else:
            out[k] = v
    return out


def main():
    ap = argparse.ArgumentParser(description="合并两份简历 JSON（断点续传）")
    ap.add_argument("--base", required=True, help="已保存的简历 JSON")
    ap.add_argument("--patch", required=True, help="新补充的简历 JSON 片段")
    ap.add_argument("--out", default="merged.json", help="合并后输出路径")
    args = ap.parse_args()

    with open(args.base, "r", encoding="utf-8") as f:
        base = json.load(f)
    with open(args.patch, "r", encoding="utf-8") as f:
        patch = json.load(f)

    merged = deep_merge(base, patch)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)
    print(f"已合并 -> {args.out}")


if __name__ == "__main__":
    main()
