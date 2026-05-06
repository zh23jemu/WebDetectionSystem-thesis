#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""列出 docx 内所有图片文件，帮助确认 image 编号。"""
import zipfile
from pathlib import Path

DOCX = Path(r"C:\Coding\WebDetectionSystem-thesis\WebDetectionSystem\基于Python的木质板材缺陷检测系统的设计与实现.docx")

with zipfile.ZipFile(DOCX) as z:
    media = [f for f in z.namelist() if f.startswith("word/media/")]
    for m in sorted(media):
        print(m)
