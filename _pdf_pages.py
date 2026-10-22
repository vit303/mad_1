# -*- coding: utf-8 -*-
from pathlib import Path
import fitz

out = Path("_pdf_pages")
out.mkdir(exist_ok=True)

jobs = {
    "Регресс.АД.pdf": list(range(18)),
    "Л.Диагн.Регр.Модели.pdf": list(range(5)),
    "Коэфф.Эласт.pdf": [0, 1, 2, 10, 11, 12, 16, 17],
    "Регр.Фикт.Перем..pdf": [2, 4, 5],
    "Л.Пуассон.Регр.Мод.pdf": [1, 4, 5, 6, 7, 10],
    "Частн.Коррел..pdf": [0, 1, 2, 3],
}

for name, pages in jobs.items():
    doc = fitz.open(name)
    stem = Path(name).stem.replace(".", "_")
    for i in pages:
        if i >= doc.page_count:
            continue
        pix = doc[i].get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
        dest = out / f"{stem}_p{i+1:02d}.png"
        pix.save(str(dest))
        print(dest.name, pix.width, pix.height)
