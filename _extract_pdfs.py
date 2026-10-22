# -*- coding: utf-8 -*-
from pathlib import Path
import fitz

files = [
    "ЛР1.2.pdf",
    "Регресс.АД.pdf",
    "Регр.Фикт.Перем..pdf",
    "Л.Диагн.Регр.Модели.pdf",
    "Л.Пуассон.Регр.Мод.pdf",
    "Коэфф.Эласт.pdf",
    "Частн.Коррел..pdf",
]

out = Path("_pdf_text")
out.mkdir(exist_ok=True)

for name in files:
    doc = fitz.open(name)
    parts = []
    for i, page in enumerate(doc):
        parts.append(f"\n\n===== PAGE {i+1}/{doc.page_count} =====\n")
        parts.append(page.get_text("text"))
    dest = out / (Path(name).stem.replace(".", "_") + ".txt")
    dest.write_text("".join(parts), encoding="utf-8")
    print(f"{name}: {doc.page_count} pages -> {dest.name} ({dest.stat().st_size} bytes)")
