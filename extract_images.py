#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 docx 提取所有图片到 tmp 目录，方便视觉确认编号。"""
import zipfile, shutil
from pathlib import Path

DOCX  = Path(r"C:\Coding\WebDetectionSystem-thesis\WebDetectionSystem\基于Python的木质板材缺陷检测系统的设计与实现.docx")
OUT   = Path(r"C:\Coding\WebDetectionSystem-thesis\tmp_imgs")
OUT.mkdir(exist_ok=True)

with zipfile.ZipFile(DOCX) as z:
    for name in z.namelist():
        if name.startswith("word/media/"):
            fname = Path(name).name
            data  = z.read(name)
            (OUT / fname).write_bytes(data)
            print(f"提取: {fname}")
