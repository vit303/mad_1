# -*- coding: utf-8 -*-
"""Generate a complete Lab 1.2 theory Word document with every formula."""
from __future__ import annotations

import io
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

OUT = Path(r"d:\python_projects\mad\lab1\Теория_ЛР1.2_Регрессионный_анализ.docx")
FONT = "Times New Roman"


def set_run_font(run, size=14, bold=False, italic=False, name=FONT):
    run.font.name = name
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:cs"), name)


def set_paragraph_format(p, first_line=True, after=6, before=0, align="justify"):
    pf = p.paragraph_format
    pf.space_after = Pt(after)
    pf.space_before = Pt(before)
    pf.line_spacing = 1.5
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    if first_line:
        pf.first_line_indent = Cm(1.25)
    else:
        pf.first_line_indent = Cm(0)
    if align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "left":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_text(doc, text, *, size=14, bold=False, italic=False, first_line=True, after=6, align="justify"):
    p = doc.add_paragraph()
    set_paragraph_format(p, first_line=first_line, after=after, align=align)
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_heading_ru(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        set_run_font(run, size=16 if level == 1 else 14, bold=True)
        run.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.first_line_indent = Cm(0)
    return p


def sanitize_latex(latex: str) -> str:
    latex = latex.replace(r"\mathbb{E}", r"\mathrm{E}")
    latex = latex.replace(r"\mathbb{N}", r"\mathrm{N}")
    latex = latex.replace(r"\overset{H_0}{\approx}", r"\approx_{H_0}")
    latex = latex.replace(r"\exists\,", r"\mathrm{exists}\,")
    latex = latex.replace(r"\varphi", r"\phi")
    return latex


def render_formula(latex: str, fontsize: int = 18) -> bytes:
    plt.rcParams["mathtext.fontset"] = "cm"
    latex = sanitize_latex(latex)
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.patch.set_facecolor("white")
    fig.text(0.0, 0.0, f"${latex}$", fontsize=fontsize, ha="left", va="bottom", color="black")
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", pad_inches=0.12, facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def add_formula(doc, latex: str, fontsize: int = 18):
    p = doc.add_paragraph()
    set_paragraph_format(p, first_line=False, after=8, before=4, align="center")
    run = p.add_run()
    try:
        png = render_formula(latex, fontsize=fontsize)
    except Exception as exc:
        raise RuntimeError(f"Formula failed: {latex}\n{exc}") from exc
    pic = run.add_picture(io.BytesIO(png))
    max_w = Cm(16.5)
    if pic.width > max_w:
        ratio = max_w / pic.width
        pic.width = max_w
        pic.height = int(pic.height * ratio)
    return p


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.clear()
        set_paragraph_format(p, first_line=False, after=3, align="justify")
        run = p.add_run(item)
        set_run_font(run, size=14)


def shade_header(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        from docx.oxml import OxmlElement

        shd = OxmlElement("w:shd")
        tcPr.append(shd)
    shd.set(qn("w:fill"), "D9E2F3")
    shd.set(qn("w:val"), "clear")


def add_table(doc, headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        set_run_font(run, size=11, bold=True)
        shade_header(cell)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.rows[i + 1].cells[j]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            set_run_font(run, size=11)
    doc.add_paragraph()
    return table


def setup_styles(doc):
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(14)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def title_page(doc):
    for _ in range(3):
        add_text(doc, "", first_line=False, after=0)
    add_text(
        doc,
        "ЛАБОРАТОРНАЯ РАБОТА 1.2",
        size=16,
        bold=True,
        first_line=False,
        align="center",
        after=6,
    )
    add_text(
        doc,
        "РЕГРЕССИОННЫЙ АНАЛИЗ",
        size=22,
        bold=True,
        first_line=False,
        align="center",
        after=18,
    )
    add_text(
        doc,
        "Полный теоретический конспект:",
        size=14,
        italic=True,
        first_line=False,
        align="center",
        after=4,
    )
    add_text(
        doc,
        "все виды моделей, каждый коэффициент, все формулы",
        size=14,
        italic=True,
        first_line=False,
        align="center",
        after=24,
    )
    add_text(
        doc,
        "По материалам лекций: «Регрессионный анализ», «Фиктивные переменные», "
        "«Диагностика регрессионной модели», «Коэффициенты эластичности, β и Δ», "
        "«Частная корреляция», «Пуассоновская регрессия», задание ЛР 1.2, "
        "а также дополнительные модели (гамма-GLM и квантильная регрессия).",
        first_line=False,
        align="center",
        after=12,
    )
    doc.add_page_break()


def build():
    doc = Document()
    setup_styles(doc)
    title_page(doc)

    add_heading_ru(doc, "1. Цель и постановка задачи", 1)
    add_text(
        doc,
        "Цель лабораторной работы 1.2 — изучение регрессионного анализа: выбор вида модели "
        "по типу зависимой переменной, оценка коэффициентов, запись уравнения, проверка "
        "качества и значимости, анализ остатков, сравнение моделей.",
    )
    add_text(
        doc,
        "Регрессионный анализ описывает зависимость количественного отклика Y от набора "
        "объясняющих переменных (регрессоров, факторов) X₁, …, Xₖ. Задача — оценить "
        "условное среднее (или иной параметр распределения) Y при заданных X и проверить, "
        "насколько эта связь статистически и содержательно значима.",
    )
    add_text(doc, "Общая запись:", first_line=True)
    add_formula(doc, r"y = h_{\beta}(x)+\varepsilon")
    add_text(
        doc,
        "где y — наблюдаемое значение зависимой переменной (эндогенной), h_β(x) — функция "
        "регрессии (систематическая часть), ε — случайная ошибка (остаточный член). "
        "Для классической линейной модели",
    )
    add_formula(doc, r"h_{\beta}(x)=\beta_0+\beta_1 x_1+\ldots+\beta_k x_k=\beta^{\!T}x")
    add_text(doc, "В матричном виде для n наблюдений:", first_line=True)
    add_formula(doc, r"y=X\beta+\varepsilon")
    add_text(
        doc,
        "Здесь y — столбец n×1, X — матрица регрессоров n×(k+1) (первый столбец — единицы, "
        "если есть свободный член), β — вектор неизвестных коэффициентов (k+1)×1, "
        "ε — вектор ошибок n×1.",
    )
    add_text(
        doc,
        "В лабораторной на данных insurance.csv зависимая переменная charges (расходы на "
        "страховку) непрерывна и положительна. Поэтому базовая модель — линейная регрессия "
        "с фиктивными переменными. Пуассоновская модель строится отдельно по пункту 11 "
        "задания, хотя charges не является счётчиком.",
    )

    add_heading_ru(doc, "2. Виды регрессионных моделей", 1)
    add_text(doc, "По числу факторов:")
    add_bullets(
        doc,
        [
            "парная (простая) регрессия: один фактор, y = β₀ + β₁x + ε;",
            "множественная регрессия: k ≥ 2 факторов.",
        ],
    )
    add_text(doc, "По форме связи:")
    add_bullets(
        doc,
        [
            "линейная по параметрам и по переменным;",
            "линейная по параметрам, но нелинейная по факторам (полином, логарифмы, произведения) — оценивается тем же МНК после замены переменных;",
            "нелинейная по параметрам (показательная, логистическая) — обычно ММП или нелинейный МНК.",
        ],
    )
    add_text(doc, "По закону распределения отклика (обобщённые линейные модели, GLM):")
    add_bullets(
        doc,
        [
            "нормальная (гауссова) линейная регрессия — МНК совпадает с ММП;",
            "пуассоновская — счётный неотрицательный отклик, связь ln λ = Xβ;",
            "гамма — положительный непрерывный скошенный отклик, Var(Y) = φμ²;",
            "биномиальная / логистическая — бинарный отклик, сигмоида.",
        ],
    )
    add_text(
        doc,
        "По целевому параметру распределения: обычная регрессия оценивает условное среднее E(Y|X); "
        "квантильная — условный квантиль Q_τ(Y|X). По способу оценивания: МНК, метод наименьших "
        "модулей (МНМ), метод моментов (ММ), метод максимального правдоподобия (ММП), "
        "штрафные методы (ридж, лассо).",
    )
    add_text(
        doc,
        "Логистическая регрессия из лекционного слайда (для полноты видов, в ЛР 1.2 на charges не применяется):",
    )
    add_formula(doc, r"h_{\beta}(x)=g(\beta^{T}x),\qquad g(z)=\dfrac{1}{1+e^{-z}}")
    add_formula(doc, r"h_{\beta}(x)=\dfrac{1}{1+e^{-(\beta_0+\beta_1 x_1+\ldots+\beta_k x_k)}}")

    add_heading_ru(doc, "3. Допущения классической линейной регрессии", 1)
    add_text(
        doc,
        "На лекции «Регрессионный анализ» сформулированы допущения, лежащие в основе линейной модели "
        "(условия Гаусса–Маркова плюс нормальность для точных t- и F-критериев):",
    )
    add_bullets(
        doc,
        [
            "Истинная форма связи между эндогенной и объясняющими переменными линейна (по параметрам).",
            "M(ε) = 0. Ошибки имеют нулевое математическое ожидание (и для точных критериев — нормальны: ε ~ N(0, σ²)).",
            "D(ε) = σ² = const. Гомоскедастичность: дисперсия ошибок постоянна.",
            "Ошибки некоррелированы друг с другом: Cov(ε_i, ε_j) = 0 при i ≠ j.",
            "Ошибки не коррелированы со всеми объясняющими переменными (регрессоры экзогенны / неслучайны в классической постановке).",
            "Независимые переменные не являются линейно зависимыми: отсутствует совершенная мультиколлинеарность, rank(X) = k+1.",
        ],
    )
    add_text(
        doc,
        "Если выполнены условия Гаусса–Маркова (без обязательной нормальности), МНК-оценки β̂ "
        "несмещены, состоятельны и эффективны в классе линейных несмещённых оценок (BLUE). "
        "Нормальность ошибок нужна, чтобы t- и F-статистики точно имели распределения Стьюдента и Фишера "
        "в конечных выборках. При больших n работает ЦПТ, и критерии остаются приближённо верными.",
    )

    add_heading_ru(doc, "4. Методы оценивания коэффициентов", 1)
    add_text(
        doc,
        "Для эмпирического уравнения ŷ = b₀ + b₁x₁ + … лекция перечисляет методы нахождения коэффициентов "
        "b₀ (свободный член) и b_j (при факторах):",
    )
    add_bullets(
        doc,
        [
            "метод наименьших квадратов (МНК) — минимизация суммы квадратов остатков;",
            "метод наименьших модулей (МНМ) — минимизация суммы абсолютных остатков; устойчив к выбросам; совпадает с квантильной регрессией при τ = 0,5;",
            "метод моментов (ММ) — приравнивание выборочных моментов теоретическим;",
            "метод максимального правдоподобия (ММП) — максимизация ln L; для нормальных ошибок совпадает с МНК.",
        ],
    )

    add_heading_ru(doc, "5. Метод наименьших квадратов", 1)
    add_heading_ru(doc, "5.1. Парная линейная регрессия", 2)
    add_text(
        doc,
        "Предположим, что функция Y(X) линейна, но конкретные коэффициенты неизвестны. "
        "Искомая функция строится так, чтобы сумма квадратов отклонений ординат экспериментальных "
        "точек y_i от ординат ŷ_i = a + b x_i была минимальной:",
    )
    add_formula(doc, r"S=\sum_{i=1}^{n}(\hat y_i-y_i)^2=\sum_{i=1}^{n}(a+b x_i-y_i)^2\to\min")
    add_text(
        doc,
        "Необходимые условия минимума — равенство нулю частных производных (нормальные уравнения):",
    )
    add_formula(doc, r"\dfrac{\partial S}{\partial a}=-2\sum_{i=1}^{n}(y_i-a-b x_i)=0")
    add_formula(doc, r"\dfrac{\partial S}{\partial b}=-2\sum_{i=1}^{n}x_i(y_i-a-b x_i)=0")
    add_text(doc, "Отсюда система:")
    add_formula(doc, r"na+b\sum x_i=\sum y_i")
    add_formula(doc, r"a\sum x_i+b\sum x_i^2=\sum x_i y_i")
    add_text(doc, "Явное решение:")
    add_formula(
        doc,
        r"b=\dfrac{n\sum x_i y_i-\sum x_i\sum y_i}{n\sum x_i^2-\left(\sum x_i\right)^2}"
        r"=\dfrac{\mathrm{cov}(x,y)}{\mathrm{var}(x)}",
    )
    add_formula(doc, r"a=\bar y-b\bar x")
    add_text(
        doc,
        "Коэффициент b — наклон: на сколько единиц в среднем изменяется y при увеличении x на 1 "
        "при сохранении линейной формы. Свободный член a — значение ŷ при x = 0; содержательно "
        "интерпретируется только если x = 0 лежит в области данных.",
    )
    add_text(doc, "Выборочные среднее, ковариация и дисперсия, входящие в формулы:")
    add_formula(doc, r"\bar x=\dfrac{1}{n}\sum_{i=1}^{n}x_i,\qquad \bar y=\dfrac{1}{n}\sum_{i=1}^{n}y_i")
    add_formula(
        doc,
        r"\mathrm{cov}(x,y)=\dfrac{1}{n-1}\sum_{i=1}^{n}(x_i-\bar x)(y_i-\bar y)",
    )
    add_formula(doc, r"s_x^2=\dfrac{1}{n-1}\sum_{i=1}^{n}(x_i-\bar x)^2")

    add_heading_ru(doc, "5.2. Множественная регрессия: матричная формула МНК", 2)
    add_text(doc, "Критерий:")
    add_formula(doc, r"S(\beta)=(y-X\beta)^{T}(y-X\beta)\to\min_{\beta}")
    add_text(doc, "Нормальное уравнение и оценка:")
    add_formula(doc, r"X^{T}X\,\hat\beta=X^{T}y")
    add_formula(doc, r"\hat\beta=(X^{T}X)^{-1}X^{T}y")
    add_text(
        doc,
        "Оценка существует и единственна тогда и только тогда, когда XᵀX обратима, то есть "
        "столбцы X линейно независимы. Прогноз и остатки:",
    )
    add_formula(doc, r"\hat y=X\hat\beta=Hy,\qquad H=X(X^{T}X)^{-1}X^{T}")
    add_formula(doc, r"e=y-\hat y=(I-H)y")
    add_text(
        doc,
        "Матрица H называется проекционной (hat-matrix). Если в модели есть свободный член, "
        "сумма сырых остатков тождественно равна нулю: Σ e_i = 0. Это алгебраическое свойство МНК, "
        "а не признак качества модели.",
    )

    add_heading_ru(doc, "5.3. Смысл каждого коэффициента β_j", 2)
    add_text(
        doc,
        "В линейной модели ŷ = β₀ + β₁x₁ + … + β_k x_k коэффициент β_j (при количественном факторе) — "
        "это частная производная:",
    )
    add_formula(doc, r"\beta_j=\dfrac{\partial E(Y\mid X)}{\partial x_j}")
    add_text(
        doc,
        "Интерпретация: при увеличении x_j на 1 единицу своего измерения и при фиксированных "
        "остальных регрессорах условное среднее Y изменяется на β_j единиц. Знак β_j — направление "
        "связи, модуль — сила в единицах Y на единицу X_j. Сравнивать «кто важнее» по сырым β_j "
        "нельзя, если факторы в разных шкалах: для этого служат эластичности, β- и Δ-коэффициенты.",
    )
    add_text(
        doc,
        "Свободный член β₀ = E(Y | все x_j = 0 и все фиктивные = 0), то есть среднее для базовой "
        "комбинации категорий при нулевых количественных факторах. Часто это вне области данных "
        "(возраст 0, ИМТ 0) и тогда β₀ нужен как калибровка уровня, а не как содержательный эффект.",
    )

    add_heading_ru(doc, "6. Разложение суммы квадратов", 1)
    add_text(doc, "Полный разброс фактических значений относительно среднего:")
    add_formula(doc, r"T_{SS}=\mathrm{SST}=\sum_{i=1}^{n}(y_i-\bar y)^2")
    add_text(doc, "Разброс, обусловленный регрессией (объяснённая сумма квадратов):")
    add_formula(doc, r"R_{SS}=\mathrm{SSR}=\sum_{i=1}^{n}(\hat y_i-\bar y)^2")
    add_text(doc, "Разброс за счёт случайных отклонений (остаточная сумма квадратов, SSE / ESS / Q_E):")
    add_formula(doc, r"E_{SS}=\mathrm{SSE}=\sum_{i=1}^{n}(y_i-\hat y_i)^2")
    add_text(
        doc,
        "Основное тождество дисперсионного разложения для модели со свободным членом:",
    )
    add_formula(doc, r"T_{SS}=R_{SS}+E_{SS}\qquad(\mathrm{SST}=\mathrm{SSR}+\mathrm{SSE})")
    add_text(
        doc,
        "Полный разброс равен сумме разбросов за счёт регрессии и за счёт случайных отклонений. "
        "На лекции SSR также обозначают Q_R, SSE — Q_E.",
    )
    add_text(doc, "Несмещённая оценка дисперсии ошибки (остаточная дисперсия) для вывода t и F:")
    add_formula(doc, r"s^2=\hat\sigma^{2}_{\mathrm{OLS}}=\dfrac{\mathrm{SSE}}{n-k-1}=\dfrac{\sum e_i^2}{n-k-1}")
    add_text(
        doc,
        "Знаменатель n − k − 1 — число степеней свободы остатков: n наблюдений минус k наклонов "
        "минус 1 свободный член. В ноутбуке это печатается как df = n − k − 1. Не путать с k = 2 "
        "или с df числителя F-критерия.",
    )

    add_heading_ru(doc, "7. Коэффициент детерминации R²", 1)
    add_text(
        doc,
        "R² — доля вариации значений y_i, обусловленная регрессией (доля закономерной изменчивости "
        "в общей изменчивости). Это характеристика точности уравнения в целом: насколько вариация "
        "зависимой переменной Y объясняется изменением независимых переменных X.",
    )
    add_formula(doc, r"R^2=\dfrac{R_{SS}}{T_{SS}}=\dfrac{\mathrm{SSR}}{\mathrm{SST}}=1-\dfrac{\mathrm{SSE}}{\mathrm{SST}}")
    add_formula(
        doc,
        r"R^2=\dfrac{\sum_{i=1}^{n}(\hat y_i-\bar y)^2}{\sum_{i=1}^{n}(y_i-\bar y)^2},"
        r"\qquad \mathrm{SSE}=\sum_{i=1}^{n}(y_i-\hat y_i)^2",
    )
    add_formula(doc, r"0\leq R^2\leq 1")
    add_text(
        doc,
        "Порог со слайда диагностики: R² ≥ 0,75 — модель приемлема. Если R² < 0,75, доля объяснённой "
        "вариации меньше 75 %, модель по этому критерию слабая (это не отменяет возможную значимость по F).",
    )
    add_text(
        doc,
        "Для линейной зависимости R² связан с бета-коэффициентами и парными корреляциями:",
    )
    add_formula(doc, r"R^{2}_{yx_1,\ldots,x_p}=\sum_{i=1}^{p}\beta_i\, r_{yx_i}")
    add_text(doc, "Эквивалентная запись через определители корреляционной матрицы:")
    add_formula(doc, r"R_{yx_1 x_2\ldots x_p}=\sqrt{1-\dfrac{\Delta r}{\Delta r_{11}}}")
    add_text(
        doc,
        "где Δr — определитель полной корреляционной матрицы вектора (y, x₁, …, x_p), "
        "а Δr₁₁ — определитель корреляционной матрицы только факторов (x₁, …, x_p).",
    )
    add_text(doc, "Проблемы R² (лекция «Диагностика модели»):")
    add_bullets(
        doc,
        [
            "нельзя сравнивать величины R² для моделей с различными зависимыми переменными (разный SST);",
            "R² никогда не уменьшается при добавлении новых объясняющих переменных, даже бесполезных;",
            "R² малопригоден для оценки качества временных рядов: значение становится очень близким к 1, что затрудняет сравнение моделей.",
        ],
    )

    add_heading_ru(doc, "8. Скорректированный коэффициент детерминации", 1)
    add_text(
        doc,
        "Скорректированный (adjusted) R² штрафует за число параметров и может уменьшиться при "
        "добавлении лишнего фактора. На слайде диагностики используется формула",
    )
    add_formula(doc, r"\bar R^{2}=1-(1-R^{2})\dfrac{n}{n-k-1}")
    add_text(
        doc,
        "где n — число наблюдений, k — число независимых переменных x_i, i = 1…k "
        "(без свободного члена). Именно эту формулу считает ноутбук ЛР 1.2.",
    )
    add_text(
        doc,
        "На слайде «Коэффициент множественной детерминации» дана несмещённая версия со множителем (n−1):",
    )
    add_formula(
        doc,
        r"\bar R^{2}_{yx_1,\ldots,x_p}"
        r"=1-\dfrac{\sum(y-\hat y)^{2}/(n-p-1)}{\sum(y-\bar y)^{2}/(n-1)}"
        r"=1-(1-R^{2})\dfrac{n-1}{n-p-1}",
    )
    add_text(
        doc,
        "Здесь p — число параметров при переменных x (то же, что k). Пакет statsmodels использует "
        "вторую формулу. Для большого n различие мало. Смысл один: сравнивать модели с разным "
        "числом факторов нужно по R̄² (и по AIC/BIC), а не по сырому R².",
    )

    add_heading_ru(doc, "9. Фиктивные переменные", 1)
    add_text(
        doc,
        "Качественные признаки (пол, курение, регион, образование, сезон) имеют два или несколько "
        "уровней. Их нельзя подставлять в регрессию как «1, 2, 3»: коэффициенты тогда не имеют "
        "сдвига среднего. Используют фиктивные (манекенные, dummy) переменные — дихотомические "
        "индикаторы 0/1.",
    )
    add_heading_ru(doc, "9.1. Два уровня (бинарный фактор)", 2)
    add_text(doc, "Регрессионная модель заработной платы с фактором «пол»:")
    add_formula(doc, r"y_i=\beta_0+\beta_1 x_{i1}+\ldots+\beta_p x_{ip}+\alpha_1 z_{i1}+\varepsilon_i")
    add_formula(
        doc,
        r"z_{i1}=1\ \mathrm{if\ male},\quad z_{i1}=0\ \mathrm{if\ female}",
    )
    add_text(
        doc,
        "Средняя заработная плата у мужчин на α₁·1 = α₁ выше, чем у женщин, при неизменных "
        "значениях других параметров модели. Проверка H₀: α₁ = 0 — это проверка существования "
        "влияния фактора «пол». Коэффициент α₁ — не «эффект пола вообще», а сдвиг относительно "
        "базового уровня (здесь — женщины, z = 0).",
    )
    add_heading_ru(doc, "9.2. Несколько уровней: правило k − 1", 2)
    add_text(
        doc,
        "Если качественный признак имеет k градаций, в модель вводят (k − 1) бинарных переменных. "
        "Пропущенный уровень — базовая (референтная) категория: его эффект «сидит» в свободном члене. "
        "Если ввести все k индикаторов вместе со свободным членом, столбцы X линейно зависимы "
        "(dummy trap): сумма dummies равна 1, XᵀX вырождена.",
    )
    add_text(doc, "Пример: образование (высшее / среднее / начальное — база):")
    add_formula(
        doc,
        r"y_i=\beta_0+\beta_1 x_{i1}+\ldots+\beta_p x_{ip}+\alpha_1 z_{i1}+\alpha_{21}z_{i21}+\alpha_{22}z_{i22}+\varepsilon_i",
    )
    add_text(
        doc,
        "z_{i21} = 1 при высшем образовании, иначе 0; z_{i22} = 1 при среднем, иначе 0. "
        "Тогда α₂₁ — надбавка высшего образования относительно начального, α₂₂ — среднего "
        "относительно начального, α₁ — по-прежнему сдвиг по полу.",
    )
    add_heading_ru(doc, "9.3. Фиктивные переменные, меняющие наклон", 2)
    add_text(
        doc,
        "Dummy могут входить не только аддитивно (сдвиг свободного члена), но и в произведениях "
        "с количественным фактором (смена наклона):",
    )
    add_formula(
        doc,
        r"y_i=\beta_0+\beta_1 x_{i1}+\beta_{11}(z_{i11}x_{i1})+\beta_{12}(z_{i12}x_{i1})"
        r"+\alpha_{21}z_{i21}+\alpha_{22}z_{i22}+\varepsilon_i",
    )
    add_text(
        doc,
        "Пример лекции: потребление Y от дохода X, сезонность сдвигает свободный член, "
        "а уровень доходности хозяйства меняет параметр β₁ («склонность к потреблению»).",
    )
    add_heading_ru(doc, "9.4. Кодирование в модели insurance", 2)
    add_text(
        doc,
        "Базовые уровни: женщина (sex = female), не курит (smoker = no), регион northeast. "
        "В уравнение входят dummy: sex:male, smoker:yes, region:northwest, region:southeast, "
        "region:southwest. Коэффициент при smoker:yes — разница средних расходов курильщика и "
        "некурящего при тех же age, bmi, children, sex, region. Коэффициент при region:southeast — "
        "разница юго-востока относительно северо-востока, а не «эффект региона сам по себе».",
    )

    add_heading_ru(doc, "10. Коэффициент эластичности", 1)
    add_text(
        doc,
        "Если зависимость y = f(x), коэффициент эластичности",
    )
    add_formula(doc, r"\Theta=f'(x)\dfrac{x}{y}")
    add_text(
        doc,
        "показывает, на сколько процентов в среднем изменится результативный признак y при "
        "изменении фактора x на 1 % от своего номинального значения. Для линейной регрессии "
        "y = a + b x производная f'(x) = b, поэтому",
    )
    add_formula(doc, r"\Theta=b\dfrac{x}{y}")
    add_text(
        doc,
        "Эластичность зависит от точки x и является переменной величиной. Чтобы получить одну "
        "цифру по выборке, берут средний коэффициент эластичности (постоянную величину):",
    )
    add_formula(doc, r"\bar\Theta=f'(\bar x)\dfrac{\bar x}{\bar y}=b\dfrac{\bar x}{\bar y}")
    add_text(
        doc,
        "Он показывает, на сколько процентов в среднем по совокупности значений фактора x "
        "изменится y при изменении x на 1 %.",
    )

    add_heading_ru(doc, "10.1. Средние частные эластичности множественной регрессии", 2)
    add_formula(doc, r"\bar\Theta_{y x_i}=b_i\dfrac{\bar x_i}{\bar y}")
    add_text(
        doc,
        "Показывают, на сколько процентов изменится в среднем y при изменении x_i на 1 % от "
        "своего среднего уровня и неизменных (средних) значениях остальных факторов; характеризуют "
        "влияние каждого фактора на результат в среднем по всей выборке. В ноутбуке считается "
        "именно эта формула: elast = b · mean(x) / mean(y).",
    )

    add_heading_ru(doc, "10.2. Частный коэффициент эластичности в заданной точке", 2)
    add_text(
        doc,
        "Показывает, на сколько процентов изменится y при изменении x_i на 1 % от заданного "
        "значения x_i и средних значениях остальных факторов:",
    )
    add_formula(
        doc,
        r"\Theta_{y x_i}=b_i\dfrac{x_i}{\hat y_{x_i;\,x_1,\ldots,x_{i-1},x_{i+1},\ldots,x_p}}",
    )
    add_text(doc, "В общем виде для любой (в том числе нелинейной) функции множественной регрессии:")
    add_formula(
        doc,
        r"\Theta_{x_j}=\dfrac{\partial\hat y}{\partial x_j}\cdot\dfrac{x_j}{\hat y},\qquad j=1,\ldots,m",
    )
    add_formula(doc, r"\hat y=f(x_1,x_2,\ldots,x_m)")
    add_text(
        doc,
        "Эластичности сравнивают по модулю: чей |Θ| больше, тот фактор сильнее влияет в процентном "
        "выражении. Но эластичность не учитывает степень колебания факторов — для этого нужны "
        "бета-коэффициенты.",
    )

    add_heading_ru(doc, "10.3. Эластичность для разных видов функций", 2)
    add_text(
        doc,
        "Ниже — стандартные формулы эластичности, которые используют при выборе нелинейной спецификации "
        "(пункт 12 лабораторной). Для каждой функции указаны уравнение и Θ = f'(x)·x/y.",
    )
    add_table(
        doc,
        ["Вид функции", "Уравнение", "Эластичность Θ"],
        [
            ["Линейная", "y = a + b x", "b · x / y"],
            ["Средняя линейная", "в точке средних", "b · x̄ / ȳ"],
            ["Гипербола", "y = a + b / x", "−b / (x y)"],
            ["Полулогарифм (ln x)", "y = a + b ln x", "b / y"],
            ["Линейно-лог (ln y)", "ln y = a + b x", "b x"],
            ["Степенная", "y = a x^b", "b (постоянна)"],
            ["Показательная", "y = a · b^x", "x ln b"],
            ["Экспонента", "y = a e^{b x}", "b x"],
            ["Квадратичная", "y = a + b x + c x²", "(b + 2 c x) x / y"],
            ["Кубическая", "y = a + b x + c x² + d x³", "(b + 2 c x + 3 d x²) x / y"],
            ["Лог-лог", "ln y = a + b ln x", "b (постоянна)"],
        ],
    )
    add_text(
        doc,
        "Для полинома по возрасту в модели charges: ∂ŷ/∂age = β_age + 2 β_{age²} · age, поэтому "
        "эластичность по возрасту Θ_age = (β_age + 2 β_{age²} · age) · age / ŷ. Аналогично для ИМТ.",
    )

    add_heading_ru(doc, "11. Бета-коэффициенты", 1)
    add_text(
        doc,
        "Бета-коэффициент показывает, на какую часть величины среднеквадратического отклонения s_y "
        "изменится переменная y с изменением соответствующей независимой переменной x_j на величину "
        "своего среднеквадратического отклонения при фиксированном уровне значений остальных факторов.",
    )
    add_formula(doc, r"\beta_j=a_j\cdot\dfrac{s_{x_j}}{s_y}")
    add_text(
        doc,
        "где a_j — тот же коэффициент регрессии, который ранее обозначался b_j. "
        "Среднеквадратические отклонения:",
    )
    add_formula(
        doc,
        r"s_{x_j}=\sqrt{\dfrac{1}{n-1}\sum_{i=1}^{n}(x_{ij}-x_{j,cp})^2},"
        r"\qquad s_y=\sqrt{\dfrac{1}{n-1}\sum_{i=1}^{n}(y_i-y_{cp})^2}",
    )
    add_text(
        doc,
        "Бета-коэффициенты безразмерны: факторы приведены к шкале «число своих σ». "
        "Их сравнивают по модулю, чтобы упорядочить факторы по силе влияния независимо от единиц измерения. "
        "В стандартизованной регрессии (все переменные в z-оценках) оценки наклонов как раз равны β_j.",
    )

    add_heading_ru(doc, "12. Дельта-коэффициенты", 1)
    add_text(
        doc,
        "Долю влияния фактора в суммарном влиянии всех факторов оценивают по величине дельта-коэффициентов:",
    )
    add_formula(doc, r"\Delta_j=r(x_j,y)\cdot\dfrac{\beta_j}{R^2}")
    add_text(
        doc,
        "где r(x_j, y) — парный коэффициент корреляции Пирсона между фактором x_j и откликом y, "
        "β_j — бета-коэффициент, R² — коэффициент детерминации. В идеале по количественным факторам "
        "Σ Δ_j близка к 1 (на фиктивные переменные эта сумма не обязана раскладываться полностью, "
        "если Δ считают только по непрерывным x). В ноутбуке: delta = corr(x, y) · beta / R².",
    )
    add_text(doc, "Парный коэффициент корреляции Пирсона, входящий в Δ_j:")
    add_formula(
        doc,
        r"r(x,y)=\dfrac{\sum(x_i-\bar x)(y_i-\bar y)}"
        r"{\sqrt{\sum(x_i-\bar x)^2\sum(y_i-\bar y)^2}}=\dfrac{\mathrm{cov}(x,y)}{s_x s_y}",
    )

    add_heading_ru(doc, "13. Частная корреляция", 1)
    add_text(
        doc,
        "Частные коэффициенты корреляции характеризуют тесноту связи между результатом и "
        "соответствующим фактором при устранении влияния других факторов, включённых в уравнение регрессии.",
    )
    add_text(doc, "Через коэффициенты детерминации полной и укороченной моделей:")
    add_formula(
        doc,
        r"r_{yx_i\cdot x_1\ldots x_{i-1}x_{i+1}\ldots x_p}"
        r"=\sqrt{1-\dfrac{R^{2}_{yx_1 x_2\ldots x_p}}{R^{2}_{yx_1\ldots x_{i-1}x_{i+1}\ldots x_p}}}",
    )
    add_text(
        doc,
        "где в числителе под корнем — множественный R² всех факторов с результатом, в знаменателе — "
        "тот же показатель, но без фактора x_i. Порядок частного коэффициента равен числу факторов, "
        "влияние которых исключается.",
    )
    add_text(doc, "Рекуррентная формула (слайд «Частные коэффициенты корреляции»). При p = 2:")
    add_formula(
        doc,
        r"r_{yx_1\cdot x_2}=\dfrac{r_{yx_1}-r_{yx_2}\,r_{x_1 x_2}}"
        r"{\sqrt{(1-r_{yx_2}^{2})(1-r_{x_1 x_2}^{2})}}",
    )
    add_formula(
        doc,
        r"r_{yx_2\cdot x_1}=\dfrac{r_{yx_2}-r_{yx_1}\,r_{x_1 x_2}}"
        r"{\sqrt{(1-r_{yx_1}^{2})(1-r_{x_1 x_2}^{2})}}",
    )
    add_text(
        doc,
        "Именно формула для p = 2 используется в ноутбуке для charges–age при исключении bmi. "
        "Если частная корреляция заметно меньше парной, связь во многом косвенная: общий фактор "
        "(здесь ИМТ) раздувает парный коэффициент. Если частная остаётся большой, связь прямая.",
    )

    add_heading_ru(doc, "14. Значимость отдельного коэффициента: t-критерий", 1)
    add_text(doc, "Гипотезы для каждого коэффициента, включая свободный член:")
    add_formula(doc, r"H_0:\ \beta_i=0\qquad H_1:\ \beta_i\neq 0")
    add_text(
        doc,
        "Для проверки используется t-критерий Стьюдента с (n − k − 1) степенями свободы:",
    )
    add_formula(doc, r"t_{\hat\beta_i}=\dfrac{\hat\beta_i}{SE_{\hat\beta_i}}")
    add_text(
        doc,
        "SE_{β̂_i} — стандартная ошибка оценки. Ковариационная матрица МНК-оценок:",
    )
    add_formula(doc, r"\widehat{\mathrm{Var}}(\hat\beta)=s^2(X^{T}X)^{-1},\qquad s^2=\dfrac{\mathrm{SSE}}{n-k-1}")
    add_formula(doc, r"SE_{\hat\beta_i}=\sqrt{\widehat{\mathrm{Var}}(\hat\beta)_{ii}}=s\sqrt{c_{ii}}")
    add_text(
        doc,
        "где c_ii — i-й диагональный элемент (XᵀX)⁻¹. Критическое значение — двусторонний квантиль "
        "Стьюдента t_{α; n−k−1}; при α = 0,05 это t_{0,975; n−k−1}. Правило:",
    )
    add_formula(doc, r"|t_{\hat\beta_i}|>t_{\alpha,\,(n-k-1)}\ \Rightarrow\ H_0\ \mathrm{rejected}")
    add_text(
        doc,
        "Эквивалентно: p-value = 2 · P(T_{n−k−1} > |t|) < α. Если |t| меньше критического, "
        "коэффициент статистически не отличим от нуля: фактор можно считать лишним (при данной спецификации), "
        "либо эффект слишком мал относительно шума. Значимость — не то же самое, что практическая важность: "
        "при огромном n даже крошечный β становится «значимым».",
    )
    add_text(doc, "Доверительный интервал для коэффициента:")
    add_formula(
        doc,
        r"\hat\beta_i\pm t_{\alpha,\,(n-k-1)}\,SE_{\hat\beta_i}",
    )

    add_heading_ru(doc, "15. Значимость уравнения в целом: F-критерий Фишера", 1)
    add_text(
        doc,
        "Необходимо проверить значимость уравнения регрессии в целом и каждого коэффициента в отдельности. "
        "Значимость уравнения в целом проверяется по F-критерию Фишера.",
    )
    add_heading_ru(doc, "15.1. Парная регрессия (слайд с Q_R и Q_E)", 2)
    add_text(doc, "Гипотеза H₀: все коэффициенты регрессии (кроме, по смыслу, возможности β₁ = 0) равны 0. Уравнение значимо, если")
    add_formula(doc, r"F=\dfrac{\sigma_R^{2}(n-2)}{\sigma_E^{2}}=\dfrac{Q_R(n-2)}{Q_E}>F_{\alpha;1;n-2}")
    add_formula(doc, r"Q_R=\sum_{k=1}^{n}(\hat y_k-\bar y)^2,\qquad Q_E=\sum_{k=1}^{n}(y_k-\hat y_k)^2")
    add_formula(doc, r"\hat y_k=a_0+a_1 x_k")
    add_text(
        doc,
        "F_{α;1;n−2} — табличное значение F-критерия на уровне α (обычно 0,05) при 1 степени свободы "
        "числителя и n−2 знаменателя. (На слайде в текстовой расшифровке степени свободы числителя и "
        "знаменателя переставлены; по формуле и стандартной теории числитель имеет df = 1, знаменатель df = n−2.)",
    )

    add_heading_ru(doc, "15.2. Множественная регрессия (рабочая формула лабораторной)", 2)
    add_formula(doc, r"H_0:\ \beta_1=\beta_2=\ldots=\beta_k=0\qquad H_1:\ \beta_i\neq 0\ \mathrm{for\ some}\ i")
    add_text(
        doc,
        "F-критерий с k и (n − k − 1) степенями свободы:",
    )
    add_formula(
        doc,
        r"F=\dfrac{\mathrm{SSR}/k}{\mathrm{SSE}/(n-k-1)}=\dfrac{R^{2}}{1-R^{2}}\cdot\dfrac{n-k-1}{k}",
    )
    add_formula(doc, r"F>F_{\alpha,\,k,\,n-k-1}\ \Rightarrow\ H_0\ \mathrm{rejected}")
    add_text(
        doc,
        "Если F > F_крит, уравнение в целом значимо: H₀ (все наклоны равны 0) отвергается. "
        "Модель лучше константы. p-value F-теста — вероятность получить такое или большее F при верной H₀. "
        "В ноутбуке F сравнивается с квантилем stats.f.ppf(0.95, k, n−k−1).",
    )

    add_heading_ru(doc, "16. Анализ остатков", 1)
    add_text(doc, "Остаток наблюдения:")
    add_formula(doc, r"\hat\varepsilon_i=e_i=y_i-\hat\beta_0-\hat\beta_1 x_{i1}-\ldots-\hat\beta_k x_{ik}=y_i-\hat y_i")
    add_text(
        doc,
        "По графику остатков проверяют: постоянство среднего и равенство 0; постоянство дисперсии; "
        "отсутствие зависимости / корреляции остатков. «Хороший» график — облако вокруг нуля без "
        "воронки, без дуги и без серий одного знака. «Плохой» — тренд, циклы, рост размаха с номером "
        "или с ŷ.",
    )

    add_heading_ru(doc, "16.1. Среднее, MAE и RMSE", 2)
    add_formula(doc, r"\bar e=\dfrac{1}{n}\sum_{i=1}^{n}e_i")
    add_text(
        doc,
        "Для МНК со свободным членом ē = 0 с точностью машинного округления. Это не доказательство "
        "качества. Качество прогноза смотрят по средним абсолютным и квадратичным ошибкам:",
    )
    add_formula(doc, r"\mathrm{MAE}=\dfrac{1}{n}\sum_{i=1}^{n}|e_i|")
    add_formula(doc, r"\mathrm{RMSE}=\sqrt{\dfrac{1}{n}\sum_{i=1}^{n}e_i^{2}}=\sqrt{\dfrac{\mathrm{SSE}}{n}}")
    add_text(
        doc,
        "MAE измеряет типичную ошибку в единицах Y и устойчивее к большим выбросам, чем RMSE. "
        "RMSE сильнее штрафует крупные отклонения и связан с SSE.",
    )

    add_heading_ru(doc, "16.2. Стандартизированные остатки и гистограмма", 2)
    add_formula(doc, r"e_i^{\mathrm{std}}=\dfrac{e_i-\bar e}{s_e},\qquad s_e=\sqrt{\dfrac{1}{n-1}\sum(e_i-\bar e)^2}")
    add_text(
        doc,
        "Гистограмма стандартизированных остатков строится для проверки похожести на нормальный закон. "
        "Число интервалов — по формуле Стерджесса:",
    )
    add_formula(doc, r"k=\mathrm{ceil}(1+3.322\log_{10} n)")
    add_text(
        doc,
        "При нормальных ошибках гистограмма колоколообразная, без тяжёлого хвоста. "
        "Теоретический стьюдентизированный остаток учитывает рычаг h_ii:",
    )
    add_formula(doc, r"t_i=\dfrac{e_i}{s\sqrt{1-h_{ii}}},\qquad h_{ii}=(H)_{ii}")

    add_heading_ru(doc, "16.3. Гомоскедастичность", 2)
    add_text(
        doc,
        "Дисперсия остатков должна быть примерно постоянна вдоль прогноза. Практический приём "
        "лабораторной: разбить ŷ на три части по квантилям 1/3 и 2/3 и сравнить выборочные дисперсии "
        "остатков. Если отношение max/min дисперсий третей больше 2, фиксируют гетероскедастичность.",
    )
    add_formula(
        doc,
        r"Q_1=q_{1/3}(\hat y),\quad Q_2=q_{2/3}(\hat y),\quad"
        r"s^{2}_{\mathrm{low}}=\mathrm{Var}(e\mid \hat y\leq Q_1)",
    )

    add_heading_ru(doc, "16.4. Независимость остатков. Дарбин–Уотсон", 2)
    add_text(
        doc,
        "Проверка независимости: график остатков против предикторов; тест серий (слишком много или "
        "мало серий одинакового знака); тест Дарбина–Уотсона",
    )
    add_formula(
        doc,
        r"DW=\dfrac{\sum_{i=2}^{n}(\hat\varepsilon_i-\hat\varepsilon_{i-1})^{2}}{\sum_{i=1}^{n}\hat\varepsilon_i^{2}}",
    )
    add_text(
        doc,
        "DW ≈ 2 — остатки некоррелированы; DW → 0 — положительная автокорреляция; DW → 4 — отрицательная. "
        "Наблюдения insurance не являются временным рядом, поэтому DW вспомогательный: порядок строк в таблице "
        "не задаёт время.",
    )

    add_heading_ru(doc, "17. Мультиколлинеарность, ридж и лассо", 1)
    add_text(
        doc,
        "Мультиколлинеарность — сильная линейная связь между регрессорами. Тогда (XᵀX) почти вырождена, "
        "оценки β̂ неустойчивы, стандартные ошибки раздуваются, знаки коэффициентов могут «переворачиваться». "
        "Методы устранения по лекции:",
    )
    add_bullets(
        doc,
        [
            "метод главных компонент;",
            "ридж-регрессия;",
            "удаление одного из коррелированных факторов (выбор исходя из экономической интерпретации);",
            "переход от несмещённых МНК-оценок к смещённым с меньшим рассеянием: меньшим M(b_j − β_j)².",
        ],
    )
    add_text(doc, "Ридж- и лассо-регрессия — штрафы в критерии МНК:")
    add_formula(
        doc,
        r"\mathrm{ridge:}\quad\min_{\beta}\sum_{i=1}^{n}(y_i-\hat y_i)^{2}+\lambda(\hat\beta_1^{2}+\ldots+\hat\beta_k^{2})",
    )
    add_formula(
        doc,
        r"\mathrm{LASSO:}\quad\min_{\beta}\sum_{i=1}^{n}(y_i-\hat y_i)^{2}+\lambda(|\hat\beta_1|+\ldots+|\hat\beta_k|)",
    )
    add_text(
        doc,
        "λ ≥ 0 — параметр регуляризации. Ридж сжимает коэффициенты к нулю, но не обнуляет; лассо может "
        "обнулить часть β̂ и тем самым отобрать факторы. Свободный член обычно не штрафуют.",
    )

    add_heading_ru(doc, "18. Нелинейная регрессия, линейная по коэффициентам", 1)
    add_text(
        doc,
        "Пункт 12 задания: построить нелинейную модель с фиктивными переменными. Допускается линейная "
        "свёртка нелинейных элементов, например y = a₀ x³ + a₁ x² log x. В работе на charges взяты "
        "квадраты возраста и ИМТ:",
    )
    add_formula(
        doc,
        r"\mathrm{charges}=\beta_0+\beta_1\mathrm{age}+\beta_2\mathrm{age}^{2}+\beta_3\mathrm{bmi}+\beta_4\mathrm{bmi}^{2}"
        r"+\beta_5\mathrm{children}+\alpha^{\!T}z+\varepsilon",
    )
    add_text(
        doc,
        "Модель нелинейна по факторам, но линейна по коэффициентам, поэтому оценивается обычным МНК: "
        "столбцы age² и bmi² просто добавляются в X. Частный эффект возраста уже не константа:",
    )
    add_formula(doc, r"\dfrac{\partial\widehat{\mathrm{charges}}}{\partial\mathrm{age}}=\beta_1+2\beta_2\cdot\mathrm{age}")
    add_text(
        doc,
        "Если β₂ > 0, рост расходов с возрастом ускоряется; если β₂ < 0 — замедляется. Сравнение с линейной "
        "моделью — по R̄², AIC/BIC и графикам остатков, а не только по тому, что «добавили квадраты».",
    )

    add_heading_ru(doc, "19. Нормальная модель ошибок: правдоподобие, AIC и BIC", 1)
    add_text(
        doc,
        "Для линейной и полиномиальной МНК в лабораторной явно принимается нормальный закон ошибок "
        "ε_i ~ N(0, σ²), независимо. Плотность одного наблюдения:",
    )
    add_formula(
        doc,
        r"f(y_i\mid x_i,\beta,\sigma^{2})=\dfrac{1}{\sqrt{2\pi\sigma^{2}}}"
        r"\exp\left(-\dfrac{(y_i-x_i^{T}\beta)^{2}}{2\sigma^{2}}\right)",
    )
    add_text(doc, "Логарифм правдоподобия выборки:")
    add_formula(
        doc,
        r"\ln L(\beta,\sigma^{2})=-\dfrac{n}{2}\ln(2\pi)-\dfrac{n}{2}\ln\sigma^{2}"
        r"-\dfrac{1}{2\sigma^{2}}\sum_{i=1}^{n}(y_i-x_i^{T}\beta)^{2}",
    )
    add_text(
        doc,
        "При фиксированном σ² максимум по β совпадает с МНК: SSE = Σ(y − Xβ)² минимизируется той же "
        "оценкой β̂. ММП-оценка дисперсии (знаменатель n, не n−k−1):",
    )
    add_formula(doc, r"\hat\sigma^{2}_{\mathrm{MLE}}=\dfrac{\mathrm{SSE}}{n}=\dfrac{1}{n}\sum_{i=1}^{n}e_i^{2}")
    add_text(doc, "Подстановка даёт достигнутый логарифм правдоподобия:")
    add_formula(
        doc,
        r"\ln\hat L=-\dfrac{n}{2}\ln(2\pi)-\dfrac{n}{2}\ln\hat\sigma^{2}-\dfrac{n}{2}"
        r"=-\dfrac{n}{2}\left(\ln(2\pi)+\ln(\mathrm{SSE}/n)+1\right)",
    )
    add_text(
        doc,
        "Информационные критерии (p — число оцениваемых параметров, включая свободный член и, в гауссовой "
        "модели statsmodels, иногда σ; в ручном расчёте лабораторной p = k + 1, как у коэффициентов β):",
    )
    add_formula(doc, r"\mathrm{AIC}=-2\ln\hat L+2p")
    add_formula(doc, r"\mathrm{BIC}=-2\ln\hat L+p\ln n")
    add_text(
        doc,
        "Меньше AIC/BIC — лучше баланс качества подгонки и сложности. BIC сильнее штрафует лишние параметры, "
        "потому что ln n > 2 при n > 8. AIC и BIC сравнимы только у моделей одной и той же зависимой "
        "переменной и одного семейства правдоподобия. Гауссовы AIC линейной и полиномиальной МНК сравнивать "
        "можно. AIC пуассоновской модели — на другой шкале (другое ln L), с нормальной МНК не сравнивают.",
    )

    add_heading_ru(doc, "20. Обобщённые линейные модели (GLM)", 1)
    add_text(
        doc,
        "GLM задаёт семейство распределений p(y | μ(x)), в котором от x зависит математическое ожидание μ(x). "
        "В семейство входят нормальное распределение с фиксированной дисперсией, экспоненциальное, "
        "пуассоновское, гамма, биномиальное. Компоненты GLM:",
    )
    add_bullets(
        doc,
        [
            "случайная: Y независимы, Y ~ экспоненциальное семейство с средним μ;",
            "систематическая: линейный предиктор η = Xβ = β₀ + β₁X₁ + … + β_k X_k;",
            "функция связи: g(μ) = η, откуда μ = g⁻¹(Xβ).",
        ],
    )
    add_text(doc, "Канонические связи: нормаль — тождественная (μ = η); Пуассон — логарифм (ln μ = η); гамма — обратная или логарифм; Бернулли — логит.")

    add_heading_ru(doc, "21. Пуассоновская регрессия", 1)
    add_text(
        doc,
        "При построении линейной регрессии предполагается, что Y непрерывна и нормальна. Если исследуемая "
        "переменная принимает только неотрицательные целочисленные значения (счётное множество) — число "
        "посетителей, лайков, обращений, заболевших, событий, баллы шкалы Лайкерта — линейная модель может "
        "дать отрицательный прогноз, а условие постоянства дисперсии остатков противоречит свойству Пуассона: "
        "дисперсия равна математическому ожиданию.",
    )
    add_heading_ru(doc, "21.1. Распределение Пуассона", 2)
    add_formula(doc, r"Y\sim\mathrm{Poisson}(\mu),\qquad f(y)=\dfrac{\mu^{y}e^{-\mu}}{y!},\quad y\in\mathbb{N}_0")
    add_formula(doc, r"\mathbb{E}(Y)=\mathrm{Var}(Y)=\mu")
    add_text(
        doc,
        "При малых μ плотность сильно асимметрична, при больших — почти симметрична (приближение нормалью). "
        "Для i-го наблюдения с регрессорами X_i:",
    )
    add_formula(
        doc,
        r"P(Y_i=y_i\mid X_i)=\dfrac{\lambda_i^{y_i}e^{-\lambda_i}}{y_i!},\quad y_i=0,1,2,\ldots",
    )
    add_formula(doc, r"M[Y_i]=D[Y_i]=\lambda_i,\qquad \lambda_i>0")

    add_heading_ru(doc, "21.2. Функция регрессии и прогноз", 2)
    add_formula(doc, r"\ln(\lambda_i)=X_i\beta=\beta_0+\beta_1 X_{1i}+\beta_2 X_{2i}+\ldots+\beta_k X_{ki}")
    add_formula(doc, r"\lambda_i=e^{X_i\beta}=e^{\beta_0+\beta_1 X_{1i}+\ldots+\beta_k X_{ki}}")
    add_text(
        doc,
        "Коэффициент β_j в пуассоновской модели — это не сдвиг среднего в штуках, а сдвиг логарифма среднего. "
        "Интерпретация: при увеличении x_j на 1 (и прочих равных) среднее λ умножается на e^{β_j}, то есть "
        "меняется на (e^{β_j} − 1)·100 процентов. Для dummy «курильщик» e^{β} — во сколько раз среднее число "
        "(или, формально, средний уровень) у курильщика больше, чем у некурящего.",
    )

    add_heading_ru(doc, "21.3. Оценивание: ММП, не МНК", 2)
    add_text(
        doc,
        "Параметры вектора β нельзя оценить базовым МНК из-за нелинейности модели. Используется метод "
        "максимального правдоподобия; максимизация — численными методами (Ньютон, IRLS). Логарифмическая "
        "функция правдоподобия:",
    )
    add_formula(
        doc,
        r"\ln L(\beta)=\sum_{i=1}^{n}\left(y_i(X_i\beta)-e^{X_i\beta}-\ln(y_i!)\right)",
    )
    add_text(
        doc,
        "Слагаемое ln(y_i!) не зависит от β и на оценку не влияет. Score-уравнение ММП для модели со "
        "свободным членом влечёт Σ (y_i − λ̂_i) = 0, поэтому среднее сырых остатков снова около нуля — "
        "и снова это не мера качества.",
    )

    add_heading_ru(doc, "21.4. Качество и значимость пуассоновской модели", 2)
    add_text(
        doc,
        "Как правило, отсутствует обычный R² и F-статистика. Значимость коэффициентов тестируется "
        "z-тестами Вальда:",
    )
    add_formula(doc, r"z_j=\dfrac{\hat\beta_j}{SE(\hat\beta_j)}\ \overset{H_0}{\approx}\ N(0,1)")
    add_text(
        doc,
        "H₀: β_j = 0 отвергается, если |z_j| > z_{1−α/2} (при α = 0,05 критическое ≈ 1,96) либо p = 2(1 − Φ(|z|)) < α.",
    )
    add_text(doc, "Значимость модели в целом — сравнение с нулевой моделью (только константа) по отношению правдоподобий:")
    add_formula(doc, r"\mathrm{LR}=2\left(\ln L_{\mathrm{full}}-\ln L_{\mathrm{null}}\right)\ \sim\ \chi^{2}_{k}")
    add_text(
        doc,
        "Модель значима относительно константы, если LR > χ²_{кр}(0,05; k). Псевдо-R² Макфаддена:",
    )
    add_formula(doc, r"R^{2}_{\mathrm{McFadden}}=1-\dfrac{\ln L_{\mathrm{full}}}{\ln L_{\mathrm{null}}}")
    add_text(
        doc,
        "Он не обязан лежать в тех же «привычных» диапазонах, что гауссов R², и порог 0,75 к нему не применяют.",
    )

    add_heading_ru(doc, "21.5. Сверхдисперсия, квази-Пуассон, отрицательная биномиальная, ZIP", 2)
    add_text(
        doc,
        "Требование равенства среднего и дисперсии часто нарушается (выбросы, неоднородность, "
        "зависимость наблюдений, неверное семейство, пропущенный предиктор): дисперсия превышает "
        "матожидание в φ раз. Оценка",
    )
    add_formula(
        doc,
        r"\hat\varphi=\dfrac{1}{n-p}\sum_{i=1}^{n}\dfrac{(y_i-\hat\mu_i)^{2}}{\hat\mu_i}",
    )
    add_text(
        doc,
        "Для Пуассона нужно φ ≈ 1. Если φ ≫ 1, z-тесты завышают значимость (SE занижены). Решения лекции:",
    )
    add_bullets(
        doc,
        [
            "квази-пуассоновская регрессия: уравнение ln μ = Xβ то же, но Var(y_i) = φ μ_i, φ оценивается по данным;",
            "отрицательная биномиальная регрессия: дополнительный параметр избыточной дисперсии;",
            "ZIP (zero-inflated Poisson): смесь структурных нулей и пуассоновских отсчётов, если нулей слишком много.",
        ],
    )
    add_formula(doc, r"y_i\sim\mathrm{QuasiPoisson}(\mu_i),\quad \mathbb{E}(y_i)=\mu_i,\ \mathrm{var}(y_i)=\varphi\mu_i,\ \ln\mu_i=\eta_i")
    add_text(
        doc,
        "У квази-Пуассона меняются стандартные ошибки, доверительные интервалы, расчёт правдоподобия, "
        "тесты значимости (t и F вместо z и χ²); AIC не вычисляется. В лабораторной на charges "
        "сверхдисперсия огромна: непрерывные расходы Пуассону не подчиняются, линейная (или гамма) модель уместнее.",
    )

    add_heading_ru(doc, "22. Гамма-регрессия (дополнительно, в методичках ЛР 1.2 нет)", 1)
    add_text(
        doc,
        "Гамма-GLM подходит для строго положительного непрерывного скошенного отклика. Дисперсия растёт "
        "с квадратом среднего: Var(Y) = φ μ². Плотность в параметризации среднего μ и дисперсионного параметра φ:",
    )
    add_formula(
        doc,
        r"f(y\mid\mu,\varphi)=\dfrac{1}{\Gamma(1/\varphi)\,(\varphi\mu)^{1/\varphi}}\,y^{1/\varphi-1}"
        r"\exp\left(-y/(\phi\mu)\right),\quad y>0",
    )
    add_text(
        doc,
        "Эквивалентно: форма (shape) α = 1/φ, масштаб (scale) θ = φμ, при этом E(Y) = αθ = μ, "
        "Var(Y) = αθ² = φμ². Каноническая связь гаммы — обратная 1/μ = Xβ, но для расходов удобнее "
        "логарифмическая:",
    )
    add_formula(doc, r"\ln\mu_i=X_i\beta\qquad\Rightarrow\qquad \mu_i=e^{X_i\beta}")
    add_text(
        doc,
        "Оценки — ММП (IRLS). Интерпретация β_j такая же, как у логарифмической связи Пуассона: "
        "e^{β_j} — множитель к среднему Y при росте x_j на 1. AIC и BIC гамма-модели сравнимы с "
        "гауссовой МНК, потому что оба — ln L одной и той же выборки charges (разные семейства, но "
        "один и тот же y); меньший AIC означает, что данное семейство лучше описывает закон данных. "
        "Для скошенных положительных расходов гамма обычно выигрывает у нормали.",
    )

    add_heading_ru(doc, "23. Квантильная регрессия (дополнительно, в методичках ЛР 1.2 нет)", 1)
    add_text(
        doc,
        "МНК оценивает условное среднее E(Y|X) и чувствителен к тяжёлому правому хвосту charges. "
        "Квантильная регрессия оценивает условный квантиль уровня τ ∈ (0, 1):",
    )
    add_formula(doc, r"Q_{\tau}(Y\mid X=x)=x^{T}\beta(\tau)")
    add_text(doc, "Критерий — минимизация суммы пинбол-потерь (check function) ρ_τ:")
    add_formula(
        doc,
        r"\rho_{\tau}(u)=u\left(\tau-\mathbf{1}\{u<0\}\right)"
        r"=\tau u\cdot 1_{u\geq 0}+(\tau-1)u\cdot 1_{u<0}",
    )
    add_formula(doc, r"\hat\beta(\tau)=\mathrm{argmin}_{\beta}\sum_{i=1}^{n}\rho_{\tau}(y_i-x_i^{T}\beta)")
    add_text(
        doc,
        "При τ = 0,5 получаем медиану: ρ_{1/2}(u) = |u|/2, то есть метод наименьших модулей. "
        "Коэффициент β_j(τ) — на сколько сдвигается квантиль уровня τ при росте x_j на 1. "
        "У курения β(0,9) обычно больше β(0,5): в верхнем хвосте расходов эффект сильнее, распределение неоднородно.",
    )
    add_text(doc, "Псевдо-R² Кенкера:")
    add_formula(
        doc,
        r"R^{1}(\tau)=1-\dfrac{\sum\rho_{\tau}(y_i-x_i^{T}\hat\beta(\tau))}"
        r"{\sum\rho_{\tau}(y_i-\hat Q_{\tau})}",
    )
    add_text(
        doc,
        "где в знаменателе — квантильная модель только со свободным членом. Гауссовские AIC/BIC "
        "для квантильной модели не считают: правдоподобие не нормальное. Сравнивают по MAE "
        "(медиана минимизирует MAE) и по интерпретации коэффициентов на разных τ.",
    )

    add_heading_ru(doc, "24. Сравнение моделей: какой критерий к чему относится", 1)
    add_table(
        doc,
        ["Модель", "Что оценивает", "Главные критерии", "С чем сравнивать AIC"],
        [
            ["Линейная МНК", "E(Y|X), ε ~ N(0,σ²)", "R², adj R², F, t, AIC/BIC, MAE, RMSE, остатки", "с другой нормальной МНК"],
            ["Полиномиальная МНК", "то же, нелинейные факторы", "то же", "с линейной МНК"],
            ["Пуассон ММП", "E(Y|X)=λ, Y~Pois", "LR, z, псевдо-R², φ", "только с другими пуассоновскими"],
            ["Гамма-GLM log", "E(Y|X)=μ>0, Var=φμ²", "AIC/BIC, z, φ", "с гауссовой МНК по AIC тех же y"],
            ["Квантиль τ=0,5", "медиана Q_{0,5}(Y|X)", "MAE, псевдо-R² Кенкера", "AIC не используется"],
        ],
    )
    add_text(
        doc,
        "Правило лабораторной: сначала смотрят, значимо ли уравнение (F или LR), затем — значимы ли "
        "коэффициенты (t или z), затем — R² относительно порога 0,75 (только для нормальной МНК), "
        "затем — график и гистограмму остатков, затем — AIC/BIC внутри одного семейства.",
    )

    add_heading_ru(doc, "25. Словарь всех коэффициентов и статистик лабораторной", 1)
    add_table(
        doc,
        ["Обозначение", "Что это", "Формула / смысл"],
        [
            ["β₀, Intercept", "свободный член", "E(Y| все x=0, все dummy=0)"],
            ["β_j", "наклон количественного фактора", "∂E(Y|X)/∂x_j в линейной модели"],
            ["α, dummy", "сдвиг категории vs база", "разница средних при прочих равных"],
            ["b, a_j", "выборочная оценка β", "МНК или ММП"],
            ["e_i, ε̂_i", "остаток", "y_i − ŷ_i"],
            ["SST, T_SS", "полная сумма квадратов", "Σ(y_i − ȳ)²"],
            ["SSR, R_SS, Q_R", "объяснённая СК", "Σ(ŷ_i − ȳ)²"],
            ["SSE, E_SS, Q_E", "остаточная СК", "Σ(y_i − ŷ_i)²"],
            ["R²", "коэффициент детерминации", "SSR/SST = 1 − SSE/SST"],
            ["R̄²", "скорректированный R²", "1−(1−R²)·n/(n−k−1) по слайду диагностики"],
            ["k", "число факторов (наклонов)", "df модели; не путать с df остатков"],
            ["n−k−1", "df остатков", "знаменатель s² и t"],
            ["s²", "оценка дисперсии ошибки", "SSE/(n−k−1)"],
            ["SE(β̂_j)", "стандартная ошибка", "s √c_jj"],
            ["t", "t-статистика коэффициента", "β̂_j / SE"],
            ["t_кр", "критическое t", "t_{1−α/2; n−k−1}"],
            ["F", "F-статистика уравнения", "(SSR/k)/(SSE/(n−k−1))"],
            ["F_кр", "критическое F", "F_{α; k, n−k−1}"],
            ["p-value", "достигнутый уровень", "P(получить ≥ наблюд. статист. | H₀)"],
            ["Θ, Э", "эластичность", "f'(x)·x/y ; средняя b·x̄/ȳ"],
            ["β_j (бета)", "стандартизованный наклон", "a_j · s_{x_j}/s_y"],
            ["Δ_j", "дельта-коэффициент", "r(x_j,y)·β_j / R²"],
            ["r_{yx}", "парная корреляция", "cov/(s_x s_y)"],
            ["r_{yx1·x2}", "частная корреляция", "см. раздел 13"],
            ["DW", "Дарбин–Уотсон", "Σ(e_i−e_{i−1})² / Σ e_i²"],
            ["MAE", "средняя абс. ошибка", "mean |e|"],
            ["RMSE", "среднеквадр. ошибка", "√mean(e²)"],
            ["ln L", "лог-правдоподобие", "зависит от семейства"],
            ["AIC", "критерий Акаике", "−2 ln L + 2p"],
            ["BIC", "критерий Шварца", "−2 ln L + p ln n"],
            ["σ²_MLE", "ММП-дисперсия нормали", "SSE/n"],
            ["λ, μ", "среднее GLM", "e^{Xβ} при log-связи"],
            ["φ", "дисперсионный параметр", "у Пуассона D/M; у гаммы Var/μ²"],
            ["z", "статистика Вальда", "β̂/SE для GLM"],
            ["LR", "отношение правдоподобий", "2(ℓ_full − ℓ_null)"],
            ["R²_McFadden", "псевдо-R²", "1 − ℓ_full/ℓ_null"],
            ["τ", "уровень квантиля", "0,5 — медиана, 0,9 — верхний хвост"],
            ["ρ_τ", "пинбол-потеря", "u(τ − 1_{u<0})"],
            ["H", "проекционная матрица", "X(XᵀX)⁻¹Xᵀ"],
            ["λ_ridge", "штраф риджа/лассо", "сила регуляризации"],
        ],
    )

    add_heading_ru(doc, "26. Как читать коэффициенты модели charges", 1)
    add_text(
        doc,
        "Линейная МНК: charges = β₀ + β_male·sex:male + β_smoke·smoker:yes + β_nw·northwest "
        "+ β_se·southeast + β_sw·southwest + β_age·age + β_bmi·bmi + β_ch·children + ε.",
    )
    add_bullets(
        doc,
        [
            "β₀ — расчётный расход для женщины, не курящей, из northeast, при age = bmi = children = 0; сам по себе не интерпретируется как реальный тариф.",
            "β_male — на сколько долларов среднее charges у мужчины отличается от женщины при тех же возрасте, ИМТ, детях, курении и регионе.",
            "β_smoke — надбавка курильщика относительно некурящего; обычно главный по |t| и по вкладу фактор.",
            "β_nw, β_se, β_sw — сдвиги регионов относительно northeast.",
            "β_age — доллары на каждый дополнительный год возраста (линейная модель).",
            "β_bmi — доллары на единицу индекса массы тела.",
            "β_ch — доллары на одного дополнительного ребёнка.",
        ],
    )
    add_text(
        doc,
        "В гамма- и пуассоновской моделях с log-связью те же имена коэффициентов читаются как логарифмы множителей: "
        "средний расход курильщика ≈ e^{β_smoke} средних некурящего. В квантильной модели при τ = 0,5 "
        "β_smoke — сдвиг медианы, а не среднего.",
    )

    add_heading_ru(doc, "27. Ответы на вопросы к защите (теория)", 1)
    add_text(
        doc,
        "1. Регрессионные модели. Уравнение регрессии — функция h_β(x), описывающая систематическую "
        "часть связи Y с X. Линейная: y = Xβ + ε, оценка МНК, ошибки нормальны с постоянной дисперсией. "
        "Пуассоновская: Y ~ Poisson(λ), ln λ = Xβ, оценка ММП.",
        first_line=True,
    )
    add_text(
        doc,
        "2. Качество: R² (доля объяснённой вариации, порог 0,75), график остатков (нулевое среднее, "
        "постоянная дисперсия, нет структуры), гистограмма стандартизированных остатков (похожесть на нормаль). "
        "Дополнительно MAE, RMSE, AIC, BIC.",
    )
    add_text(
        doc,
        "3. Значимость уравнения — F (или LR для Пуассона) против модели «только константа». "
        "Значимость коэффициентов — t (МНК) или z Вальда (GLM): |статистика| > критического.",
    )
    add_text(
        doc,
        "4. Анализ остатков: e = y − ŷ; проверяют M(e)=0, гомоскедастичность, некоррелированность (DW), "
        "нормальность по гистограмме. Нулевое среднее МНК — тождество, не вывод о качестве.",
    )
    add_text(
        doc,
        "5. R² = SSR/SST = 1 − SSE/SST. Скорректированный R̄² = 1 − (1−R²)·n/(n−k−1) "
        "(версия слайда диагностики) штрафует за число факторов.",
    )
    add_text(
        doc,
        "6. МНК: минимизация Σ(y_i − ŷ_i)²; решение β̂ = (XᵀX)⁻¹ Xᵀy. Для парной регрессии "
        "b = cov(x,y)/var(x), a = ȳ − b x̄.",
    )
    add_text(
        doc,
        "7. Виды: парная и множественная; линейная и нелинейная по факторам; МНК / МНМ / ММП; "
        "GLM (нормаль, Пуассон, гамма); квантильная; штрафные (ридж, лассо); логистическая для бинарного Y.",
    )

    add_heading_ru(doc, "28. Краткая шпаргалка вычислений «от данных к выводу»", 1)
    add_text(doc, "Шаг 1. Собрать y и матрицу X (единицы + количественные + k−1 dummy).")
    add_text(doc, "Шаг 2. β̂ = (XᵀX)⁻¹ Xᵀy, ŷ = Xβ̂, e = y − ŷ.")
    add_text(doc, "Шаг 3. SST, SSR, SSE → R², R̄², s², F, t, SE.")
    add_text(doc, "Шаг 4. Эластичности b_j x̄_j / ȳ, бета a_j s_{x_j}/s_y, дельта r_j β_j / R².")
    add_text(doc, "Шаг 5. Частная корреляция по формуле для p = 2 при необходимости.")
    add_text(doc, "Шаг 6. График e, гистограмма e^{std}, дисперсии по третям ŷ, DW, MAE, RMSE.")
    add_text(doc, "Шаг 7. Для нормальной модели: σ² = SSE/n, ln L, AIC, BIC.")
    add_text(doc, "Шаг 8. Пуассон / гамма / квантиль — по своим критериям; сравнить модели честно.")
    add_text(
        doc,
        "Шаг 9. Записать уравнение, назвать значимые коэффициенты, интерпретировать знаки и главный фактор, "
        "сделать вывод о приемлемости модели и о том, какое семейство лучше описывает закон Y.",
    )

    add_heading_ru(doc, "Литература (из задания ЛР 1.2 и лекций)", 1)
    add_bullets(
        doc,
        [
            "Айвазян С. А. Методы эконометрики. — М.: Магистр: ИНФРА-М, 2010.",
            "Лекции курса: регрессионный анализ; фиктивные переменные; диагностика модели; эластичность, β, Δ; частная корреляция; пуассоновская регрессия.",
            "machinelearning.ru: одномерная линейная регрессия; коэффициент детерминации; анализ регрессионных остатков.",
        ],
    )

    doc.save(OUT)
    print("saved", OUT, "size", OUT.stat().st_size)


if __name__ == "__main__":
    build()
