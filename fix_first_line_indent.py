#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""给 docx 中所有正文段落（Normal 样式）添加首行缩进（2字符 ≈ Pt(24)）。"""

import shutil
from pathlib import Path
from docx import Document
from docx.shared import Pt

DOCX_PATH = Path(r"C:\Coding\WebDetectionSystem-thesis\WebDetectionSystem\基于Python的木质板材缺陷检测系统的设计与实现.docx")
BACKUP    = DOCX_PATH.with_name(DOCX_PATH.stem + "_缩进前备份.docx")

# 需要加缩进的样式名称集合（Normal 及其中文变体）
TARGET_STYLES = {"Normal", "正文", "Body Text"}

def fix_indent():
    shutil.copy2(DOCX_PATH, BACKUP)
    print(f"备份: {BACKUP}")

    doc = Document(DOCX_PATH)
    updated = 0

    for para in doc.paragraphs:
        style_name = para.style.name
        # 匹配 Normal 及以 Normal 开头的样式（如 Normal (Web)）
        is_normal = (
            style_name in TARGET_STYLES
            or style_name.startswith("Normal")
            or style_name.startswith("正文")
        )
        if not is_normal:
            continue
        if not para.text.strip():
            continue
        # 设置首行缩进
        fmt = para.paragraph_format
        if fmt.first_line_indent != Pt(24):
            fmt.first_line_indent = Pt(24)
            updated += 1

    doc.save(DOCX_PATH)
    print(f"首行缩进已添加，共更新 {updated} 个段落。")


if __name__ == "__main__":
    fix_indent()
