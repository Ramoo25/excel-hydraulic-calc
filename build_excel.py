# -*- coding: utf-8 -*-
"""
build_excel.py — генератор hydr_calc.xlsx

Калькулятор потерь давления (Дарси-Вейсбах) на участке трубы с водой
в диапазоне температур 5..80 °C с шагом 5 °C.

Все расчёты выполняются live-формулами Excel: пользователь меняет ввод на
листе «Ввод» — таблица пересчитывается.

Методика:
  v  = Q / (3600 * f),  f = pi*d^2/4
  Re = v * d_m / nu
  lambda = 64/Re                      (Re <= 2300, ламинарный)
  lambda = 0.11*(k/d + 68/Re)^0.25    (Re > 2300, Альтшуль)
  R  = lambda / d_m * rho * v^2 / 2   [Па/м]
  dp = R * L                          [кПа]
  dh = dp_Pa / (rho * g)              [м]
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT_FILE = "hydr_calc.xlsx"

# Свойства теплоносителей при атмосферном давлении.
# Диапазон расчёта: -20 … +80 °C, шаг 5 °C.
# Вода: rho [кг/м³], nu [м²/с] (T ≥ 5 °C; ниже — лёд, ячейки пустые).
# ПГ 40%: KasperCalc/Dow данные (μ, ρ в °F → μ/ρ → ν), точка кристаллизации ≈ -24 °C.
TEMP_LIST = list(range(-20, 85, 5))          # -20, -15, ..., 80  (21 значение)

WATER_PROPS = {
    5:  (999.96, 1.5189e-6),
    10: (999.70, 1.3069e-6),
    15: (999.10, 1.1394e-6),
    20: (998.21, 1.0037e-6),
    25: (997.05, 0.8927e-6),
    30: (995.65, 0.8007e-6),
    35: (994.03, 0.7246e-6),
    40: (992.22, 0.6579e-6),
    45: (990.21, 0.6017e-6),
    50: (988.04, 0.5531e-6),
    55: (985.69, 0.5102e-6),
    60: (983.19, 0.4740e-6),
    65: (980.55, 0.4434e-6),
    70: (977.76, 0.4131e-6),
    75: (974.79, 0.3879e-6),
    80: (971.80, 0.3647e-6),
}

# (T, rho, nu_CSt)  для 40% раствора пропиленгликоля; nu в мм²/с (cSt)
PG40 = [
    (-20, 1048.0, 46.66),
    (-15, 1046.3, 31.60),
    (-10, 1044.6, 22.12),
    (-5,  1042.8, 15.94),
    (0,   1041.0, 11.81),
    (5,   1039.1,  8.968),
    (10,  1037.1,  6.954),
    (15,  1035.0,  5.507),
    (20,  1032.7,  4.444),
    (25,  1030.2,  3.640),
    (30,  1027.5,  3.035),
    (35,  1024.6,  2.557),
    (40,  1021.5,  2.192),
    (45,  1018.3,  1.927),
    (50,  1015.0,  1.655),
    (55,  1011.6,  1.463),
    (60,  1008.2,  1.299),
    (65,  1004.7,  1.174),
    (70,  1001.2,  1.059),
    (75,   997.63, 0.9623),
    (80,   994.0,  0.8853),
]

# -------------------------------------------------------------------- стили
F_TITLE = Font(size=14, bold=True)
F_HEAD  = Font(size=10, bold=True, color="FFFFFF")
F_UNIT  = Font(size=9, italic=True, color="555555")
F_NOTE  = Font(size=9, italic=True, color="777777")
F_LABEL = Font(size=11)
F_INPUT = Font(size=11, bold=True)

FILL_HEAD = PatternFill("solid", fgColor="305496")
FILL_INP  = PatternFill("solid", fgColor="FFF2CC")
FILL_T    = PatternFill("solid", fgColor="D9E1F2")

THIN  = Side(style="thin", color="B0B0B0")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

C  = Alignment(horizontal="center", vertical="center")
CW = Alignment(horizontal="center", vertical="center", wrap_text=True)
LW = Alignment(horizontal="left", vertical="center", wrap_text=True)

wb = Workbook()

# =================================================================== ЛИСТ 1
# «Ввод» — исходные данные + именованные ячейки
ws1 = wb.active
ws1.title = "Ввод"

ws1["A1"] = "Ввод исходных данных"
ws1["A1"].font = F_TITLE
ws1.merge_cells("A1:E1")

FILL_OFF = PatternFill("solid", fgColor="E0E0E0")   # серый — неиспользуемая ячейка
F_OFF    = Font(size=11, italic=True, color="909090")

MODE_Q  = "Расход (Q)"
MODE_V  = "Скорость (v)"

FLUID_W  = "Вода"
FLUID_PG = "ПГ 40%"

rows = [
    ("Теплоноситель",                FLUID_W, "—",
     "выпадающий список: Вода или 40% раствор пропиленгликоля (защита ≈ −24 °C)"),
    ("Внутренний диаметр трубы d", 50.0, "мм",
     "внутренний (проходной) диаметр"),
    ("Эквивалентная шероховатость k", 0.1, "мм",
     "нерж. сталь ~0.02-0.05; сталь новая ~0.05-0.1; сталь б/у 0.1-0.3; чугун ~0.3-0.5"),
    ("Расход воды Q", 5.0, "м³/ч",
     "используется, если режим = «Расход (Q)»"),
    ("Скорость воды v", None, "м/с",
     "используется, если режим = «Скорость (v)»; при пустой ячейке задайте значение"),
    ("РЕЖИМ ЗАДАНИЯ", MODE_Q, "—",
     "выберите из списка: задавать расход Q или скорость v"),
    ("Длина участка L", 100.0, "м", "полная длина прямого участка"),
]
# строки 3..9: теплоноситель, d, k, Q, v, режим, L
for i, (label, val, unit, note) in enumerate(rows):
    r = 3 + i
    ws1[f"A{r}"] = label
    ws1[f"A{r}"].font = F_LABEL
    if val is not None:
        ws1[f"C{r}"] = val
    ws1[f"C{r}"].font = F_INPUT
    ws1[f"C{r}"].fill = FILL_INP
    ws1[f"C{r}"].alignment = C
    ws1[f"C{r}"].border = B_ALL
    ws1[f"D{r}"] = unit
    ws1[f"D{r}"].font = F_UNIT
    ws1[f"E{r}"] = note
    ws1[f"E{r}"].font = F_NOTE

ws1["A11"] = "Расчётные (используемые в таблице) значения:"
ws1["A11"].font = F_LABEL
ws1["A12"] = "• расход  Q*"
ws1["C12"] = (f'=IF(Regim="{MODE_Q}",Q_m3h,'
              '3600*v_ms*(PI()*(d_mm/1000)^2/4))')
ws1["D12"] = "м³/ч"
ws1["A13"] = "• скорость v*"
ws1["C13"] = (f'=IF(Regim="{MODE_Q}",'
              'Q_m3h/3600/(PI()*(d_mm/1000)^2/4),'
              'v_ms)')
ws1["D13"] = "м/с"
for rr in (12, 13):
    ws1[f"A{rr}"].font = F_LABEL
    ws1[f"C{rr}"].number_format = "0.000"
    ws1[f"C{rr}"].alignment = C
    ws1[f"C{rr}"].border = B_ALL
    ws1[f"D{rr}"].font = F_UNIT

# именованные диапазоны
for name, cell in (("Fluid", "$C$3"), ("d_mm", "$C$4"), ("k_mm", "$C$5"),
                   ("Q_m3h", "$C$6"), ("v_ms", "$C$7"), ("Regim", "$C$8"),
                   ("L_m", "$C$9"), ("Q_used", "$C$12"), ("v_used", "$C$13")):
    wb.defined_names.add(DefinedName(name, attr_text=f"Ввод!{cell}"))

# валидация ввода
dv_pos = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                        allow_blank=False, showErrorMessage=True,
                        errorTitle="Ошибка ввода",
                        error="Значение должно быть положительным числом.")
for addr in ("C4", "C6", "C7", "C9"):
    dv_pos.add(addr)
ws1.add_data_validation(dv_pos)

dv_nonneg = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                           allow_blank=False, showErrorMessage=True,
                           errorTitle="Ошибка ввода",
                           error="Шероховатость не может быть отрицательной.")
dv_nonneg.add("C5")
ws1.add_data_validation(dv_nonneg)

dv_mode = DataValidation(type="list", formula1=f'"{MODE_Q},{MODE_V}"',
                         allow_blank=False, showErrorMessage=True,
                         errorTitle="Ошибка режима",
                         error="Выберите вариант из списка.")
dv_mode.add("C8")
ws1.add_data_validation(dv_mode)

dv_fluid = DataValidation(type="list", formula1=f'"{FLUID_W},{FLUID_PG}"',
                          allow_blank=False, showErrorMessage=True,
                          errorTitle="Ошибка выбора",
                          error="Выберите теплоноситель из списка.")
dv_fluid.add("C3")
ws1.add_data_validation(dv_fluid)

# подсветка неиспользуемой ячейки (серым)
ws1.conditional_formatting.add(
    "C6", FormulaRule(formula=[f'$C$8<>"{MODE_Q}"'], fill=FILL_OFF, font=F_OFF))
ws1.conditional_formatting.add(
    "C7", FormulaRule(formula=[f'$C$8="{MODE_Q}"'], fill=FILL_OFF, font=F_OFF))

ws1["A15"] = "Примечания:"
ws1["A16"] = "• жёлтые ячейки — редактируемые исходные данные;"
ws1["A17"] = ('• режим задаёт, что вводить: расход Q (м³/ч) или скорость v (м/с); '
              'неиспользуемая ячейка серееет автоматически, заполнять её не надо;')
ws1["A18"] = "• серые ячейки Q* и v* — вычисляемые (то, что реально пойдёт в таблицу);"
ws1["A19"] = "• шероховатость — эквивалентная (по состоянию трубы);"
ws1["A20"] = ('• теплоноситель: «Вода» (свойства IAPWS, T ≥ 5 °C) или «ПГ 40%» '
              '(40% раствор пропиленгликоля, свойства KasperCalc/Dow, T = −20…+80 °C);')
ws1["A21"] = "• результат — лист «Расчёт»: потери давления по температуре."
for r in range(16, 22):
    ws1[f"A{r}"].font = F_NOTE

ws1.column_dimensions["A"].width = 32
ws1.column_dimensions["C"].width = 12
ws1.column_dimensions["D"].width = 10
ws1.column_dimensions["E"].width = 70

# =================================================================== ЛИСТ 2
# «Свойства воды» — справочник: rho(T), nu(T) для воды и ПГ 40% + расчётные ρ*, ν*
ws2 = wb.create_sheet("Свойства воды")
ws2["A1"] = "Свойства теплоносителей при атмосферном давлении"
ws2["A1"].font = F_TITLE
ws2.merge_cells("A1:G1")
ws2["A2"] = ("Вода: IAPWS (T ≥ 5 °C, ниже — лёд).  ПГ 40%: данные KasperCalc/Dow "
             "(интерполяция к шагу 5 °C), кристаллизация ≈ −24 °C.")
ws2["A2"].font = F_NOTE

hdr = ["T, °C", "rho воды, кг/м³", "nu воды, мм²/с",
       "rho ПГ 40%, кг/м³", "nu ПГ 40%, мм²/с",
       "rho* (расч.), кг/м³", "nu* (расч.), м²/с"]
for j, h in enumerate(hdr):
    c = ws2.cell(row=4, column=1 + j, value=h)
    c.font = F_HEAD
    c.fill = FILL_HEAD
    c.alignment = CW
    c.border = B_ALL

RES_ROW0 = 5
N_RES = len(TEMP_LIST)
for i in range(N_RES):
    r = RES_ROW0 + i
    t = TEMP_LIST[i]
    ws2.cell(row=r, column=1, value=t).fill = FILL_T
    wp = WATER_PROPS.get(t)         # (rho, nu) или None (T < 5 → лёд)
    pg = PG40[i]
    ws2.cell(row=r, column=2, value=wp[0] if wp else None).number_format = "0.00"
    ws2.cell(row=r, column=3,
             value=wp[1] * 1e6 if wp else None).number_format = "0.000"
    ws2.cell(row=r, column=4, value=pg[1]).number_format = "0.00"
    ws2.cell(row=r, column=5, value=pg[2]).number_format = "0.000"
    # расчётные значения по выбранному теплоносителю
    ws2.cell(row=r, column=6,
             value=f'=IF(Fluid="{FLUID_W}",B{r},D{r})').number_format = "0.00"
    ws2.cell(row=r, column=7,
             value=f'=IF(Fluid="{FLUID_W}",C{r},E{r})*0.000001'
             ).number_format = "0.0000E+00"
    for j in range(1, 8):
        cell = ws2.cell(row=r, column=j)
        cell.alignment = C
        cell.border = B_ALL

ws2["A28"] = ("Выбор теплоносителя — ячейка «Теплоноситель» на листе «Ввод»; "
              "колонки rho*/nu* — то, что реально идёт в расчёт.")

for col, w in zip("ABCDEFG", (8, 13, 13, 14, 14, 15, 15)):
    ws2.column_dimensions[col].width = w
ws2.freeze_panes = "A5"

# =================================================================== ЛИСТ 3
# «Расчёт» — главная таблица 5…80 °C
ws3 = wb.create_sheet("Расчёт")

ws3["A1"] = "Потери давления на участке трубы (вода) — расчёт по температурам"
ws3["A1"].font = F_TITLE
ws3.merge_cells("A1:O1")

ws3["A2"] = ('=" Исходные данные:  теплоноситель: "&Fluid&";  d = "&d_mm&" мм;   '
             'k = "&k_mm&" мм;   режим: "&Regim'
             '&";   Q* = "&ROUND(Q_used,3)&" м³/ч;   v* = "&ROUND(v_used,3)&" м/с;   '
             'L = "&L_m&" м"')
ws3["A2"].font = F_UNIT

# заголовки (две строки: группа + колонка)
props = [
    ("T", "°C"),        ("rho", "кг/м³"),   ("nu", "м²/с"),
    ("v", "м/с"),       ("Re", "—"),        ("k/d", "—"),
    ("lambda", "—"),    ("R", "Па/м"),      ("dP", "кПа"),   ("dH", "м вод. ст."),
    ("lam_cw1", "—"),   ("lam_cw2", "—"),   ("lam_cw3", "—"), ("d_lam", "%"),
    ("d_R", "Па/м"),
]
# строка 3 — названия, строка 4 — единицы
for j, (name, unit) in enumerate(props, start=1):
    c3 = ws3.cell(row=3, column=j, value=name)
    c4 = ws3.cell(row=4, column=j, value=unit)
    for c in (c3, c4):
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.alignment = CW
        c.border = B_ALL
# «человеческие» названия в строке 3
titles = ["T", "Плотность ρ", "Вязкость ν", "Скорость v", "Рейнольдс Re",
          "Отн. шероховатость k/d", "Коэффициент трения λ", "Удельная потеря R",
          "Потеря давления Δp", "Потеря напора Δh",
          "λ Колбрука (ит. 1)", "λ Колбрука (ит. 2)", "λ Колбрука (ит. 3)",
          "Разность λ (Альтшуль vs Колбрук)",
          "Δ R (Альтшуль − Колбрук)"]
for j, t in enumerate(titles, start=1):
    ws3.cell(row=3, column=j, value=t).font = F_HEAD

DM = "(d_mm/1000)"   # диаметр в метрах — для краткости формул

formulas = {
    "A": "={t}",
    "B": "='Свойства воды'!F{sr}",
    "C": "='Свойства воды'!G{sr}",
    "D": "=v_used",
    "E": f'=IF(C{{r}}="","",D{{r}}*{DM}/C{{r}})',
    "F": "=k_mm/d_mm",
    "G": '=IF(E{r}="","",IF(E{r}<=2300,64/E{r},0.11*(F{r}+68/E{r})^0.25))',
    "H": f'=IF(G{{r}}="","",G{{r}}/{DM}*(B{{r}}*D{{r}}^2)/2)',
    "I": '=IF(H{r}="","",H{r}*L_m/1000)',
    "J": '=IF(I{r}="","",I{r}*1000/(B{r}*9.80665))',
    # --- Колбрук-Уайт: 3 итерации простых приближений от λ Альтшуля ---
    # λ_next = (1 / (-2·LOG10(k/(3.7·d) + 2.51/(Re·√λ_prev))))²
    "K": '=IF(E{r}="","",IF(E{r}<=2300,64/E{r},'
         '(1/(-2*LOG10(k_mm/(3.7*d_mm)+2.51/(E{r}*SQRT(G{r})))))^2))',
    "L": '=IF(E{r}="","",IF(E{r}<=2300,64/E{r},'
         '(1/(-2*LOG10(k_mm/(3.7*d_mm)+2.51/(E{r}*SQRT(K{r})))))^2))',
    "M": '=IF(E{r}="","",IF(E{r}<=2300,64/E{r},'
         '(1/(-2*LOG10(k_mm/(3.7*d_mm)+2.51/(E{r}*SQRT(L{r})))))^2))',
    "N": '=IF(M{r}="","",(G{r}-M{r})/M{r})',
    # разность удельных потерь в Па/м: Альтшуль минус Колбрук
    "O": f'=IF(M{{r}}="","",(G{{r}}-M{{r}})/{DM}*(B{{r}}*D{{r}}^2)/2)',
}
numfmt = {
    "A": "0",    "B": "0.00", "C": "0.0000E+00", "D": "0.000",
    "E": "0",    "F": "0.00000", "G": "0.0000", "H": "0",
    "I": "0.00", "J": "0.000", "K": "0.0000", "L": "0.0000",
    "M": "0.0000", "N": "0.0%",
    "O": "0",
}

for i in range(N_RES):
    r = 5 + i                 # строка на листе «Расчёт»
    sr = RES_ROW0 + i         # строка на листе «Свойства воды»
    ws3.cell(row=r, column=1, value=TEMP_LIST[i]).fill = FILL_T
    for col in "ABCDEFGHIJKLMNO":
        cell = ws3[f"{col}{r}"]
        if col == "A":
            continue  # уже записано
        cell.value = formulas[col].format(r=r, sr=sr)
    for col in "ABCDEFGHIJKLMNO":
        cell = ws3[f"{col}{r}"]
        cell.number_format = numfmt[col]
        cell.alignment = C
        cell.border = B_ALL

# скрываем промежуточные итерации Колбрука
for c in ("K", "L", "M"):
    ws3.column_dimensions[c].hidden = True

tot = ws3.cell(row=5 + N_RES, column=1,
               value="min/max Δp по диапазону")
tot.font = F_NOTE
ws3[f"I{5 + N_RES}"] = f'=IF(MAX(I5:I{4+N_RES})=MIN(I5:I{4+N_RES}),I5,"см. таблицу")'
ws3[f"I{5 + N_RES}"].font = F_NOTE

widths = {"A": 6, "B": 11, "C": 13, "D": 9, "E": 10, "F": 9,
          "G": 11, "H": 12, "I": 12, "J": 13,
          "K": 12, "L": 12, "M": 12, "N": 13, "O": 14}
for col, w in widths.items():
    ws3.column_dimensions[col].width = w
ws3.freeze_panes = "A5"

# =================================================================== ЛИСТ 4
# «Методика» — описание формул
ws4 = wb.create_sheet("Методика")
ws4["A1"] = "Методика расчёта"
ws4["A1"].font = F_TITLE

lines = [
    "",
    "0. Режим задания режима потока:",
    "   • «Расход (Q)» — в таблице используется введённый расход Q, м³/ч,",
    "     скорость вычисляется: v = Q / (3600 · f), где f = π·d²/4 — площадь сечения;",
    "   • «Скорость (v)» — используется введённая скорость v, м/с,",
    "     расход вычисляется: Q = 3600 · v · f.",
    "   Выбор — ячейка «РЕЖИМ ЗАДАНИЯ» на листе «Ввод» (выпадающий список);",
    "   неиспользуемая из двух ячеек подсвечивается серым автоматически.",
    "",
    "1. Скорость потока:",
    "   f = π·d²/4 — площадь сечения; d — внутренний диаметр, м",
    "   v* и Q* (расчётные) формируются по выбранному режиму (см. выше).",
    "",
    "2. Число Рейнольдса:",
    "   Re = v · d / ν   (d — м; ν — кинематическая вязкость воды при данной температуре)",
    "",
    "3. Коэффициент гидравлического трения λ:",
    "   • ламинарный режим (Re ≤ 2300):        λ = 64 / Re",
    "   • турбулентный режим  (Re > 2300):     λ = 0.11 · (k/d + 68/Re)^0.25   (формула Альтшуля)",
    "",
    "4. Потери давления по Дарси-Вейсбаху (трения по длине):",
    "   R  = λ · (1/d) · (ρ·v²)/2      [Па/м]  — удельная",
    "   Δp = R · L                      [кПа]  — на участке",
    "",
    "5. Потеря напора:",
    "   Δh = Δp / (ρ·g), g = 9.80665 м/с²   [м вод. ст.]",
    "",
    "5а. Сравнение с формулой Колбрука-Уайта (это эталон, уравнение неявное):",
    "    1/√λ = −2·log10( k/(3.7·d) + 2.51/(Re·√λ) )",
    "    В Excel решается последовательными приближениями (3 итерации):",
    "    λ(i+1) = (1/(-2·log10( k/(3.7·d) + 2.51/(Re·√λ(i)) )))²,",
    "    старт λ(0) = λ Альтшуля. В ламинарной зоне — просто 64/Re.",
    "    Колонка «Разность λ» = (λ_Альтшуль − λ_Колбрука)/λ_Колбрука;",
    "    поскольку Δp ∝ λ, тот же процент — разность Δp между методами.",
    "Свойства воды: стандартные таблицы теплофизических свойств воды (IAPWS),",
    "значения ρ и ν при атмосферном давлении; диапазон 5…80 °C с шагом 5 °C.",
    "",
    "Теплоносители:",
    "• «Вода» — стандартные таблицы IAPWS, расчёт возможен при T ≥ 5 °C (ниже — лёд);",
    "• «ПГ 40%» — водный раствор пропиленгликоля 40% масс., свойства по данным KasperCalc",
    "  (кривые Dow, приведены к шагу 5 °C), диапазон −20…+80 °C, кристаллизация ≈ −24 °C.",
    "  Вязкость ПГ кратно выше воды: при температурах ниже 0 °C режим течения часто",
    "  становится ламинарным — потери давления могут в разы превышать водные.",
    "",
    "Примечания:",
    "• местные сопротивления (отводы, арматура) не учитываются — считается только трение по длине;",
    "• при малых Re формула Альтшуля плавно переходит в зависимость для гладких труб (Блазиус);",
    "• все вычисления — формулами Excel: изменение исходных данных на листе «Ввод»",
    "  пересчитывает таблицу автоматически.",
]
for i, s in enumerate(lines):
    cell = ws4.cell(row=2 + i, column=1, value=s if s else None)
    if s.startswith(("1.", "2.", "3.", "4.", "5.")):
        cell.font = Font(size=11, bold=True)
    elif s:
        cell.font = F_LABEL
ws4.column_dimensions["A"].width = 100

wb.save(OUT_FILE)
print(f"OK: {OUT_FILE} создан (4 листа: Ввод, Свойства воды, Расчёт, Методика)")

