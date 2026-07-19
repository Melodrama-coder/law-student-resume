#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
渲染 5 种模板的缩略图 PNG 到 assets/templates/，供 SKILL 预览与 preview.html 使用。

依赖：python-docx / pymupdf(pillow 不需要) / win32com 或 LibreOffice（DOCX->PDF）。
用法：
  python scripts/render_thumbnails.py
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import generate_resume as G  # noqa: E402

TPL_DIR = os.path.join(ROOT, "assets", "templates")
os.makedirs(TPL_DIR, exist_ok=True)
DATA = os.path.join(ROOT, "assets", "sample_resume.json")

with open(DATA, "r", encoding="utf-8") as f:
    data = G.json.load(f)

# 用稍小字号确保缩略图为单页（仅用于预览，不代表最终输出尺寸）
SIZE = 10
for tpl in ["classic", "minimal", "centered", "modern", "compact"]:
    T = G.TEMPLATES[tpl]
    doc = G.build_document(data, "微软雅黑", SIZE, T, None)
    tmp_docx = os.path.join(TPL_DIR, f"_{tpl}.docx")
    pdf_path = os.path.join(TPL_DIR, f"_{tpl}.pdf")
    doc.save(tmp_docx)
    ok, _ = G.convert_to_pdf(tmp_docx, pdf_path)
    if not ok:
        print(f"  [warn] {tpl} PDF 转换失败，跳过")
        continue
    import fitz  # PyMuPDF
    d = fitz.open(pdf_path)
    pix = d[0].get_pixmap(dpi=110)
    out_png = os.path.join(TPL_DIR, f"{tpl}.png")
    pix.save(out_png)
    d.close()
    print(f"  saved {out_png}")
print("✅ 缩略图渲染完成 ->", TPL_DIR)
