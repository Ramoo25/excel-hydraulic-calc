# -*- coding: utf-8 -*-
"""Повторная проверка ссылок после исправления RES_ROW0."""
from openpyxl import load_workbook

wb = load_workbook("hydr_calc.xlsx")
ws = wb["Расчёт"]
ws2 = wb["Свойства воды"]

ok = True
for r in range(5, 21):
    sr = r  # расшифровка ссылки: строка свойств должна совпадать с r
    b = ws[f"B{r}"].value
    expect = f"='Свойства воды'!B{sr}"
    t_calc = ws[f"A{r}"].value
    t_prop = ws2.cell(row=sr, column=1).value
    if b != expect or t_calc != t_prop:
        ok = False
        print(f"MISMATCH строка {r}: {b} != {expect}; T {t_calc} != {t_prop}")
    rho = ws2.cell(row=sr, column=2).value
    nu = ws2.cell(row=sr, column=3).value
    if not (isinstance(rho, float) and isinstance(nu, float)):
        ok = False
        print(f"ПУСТО в свойствах строка {sr}")

print("Все 16 строк согласованы OK" if ok else "ОШИБКИ ЕСТЬ")
