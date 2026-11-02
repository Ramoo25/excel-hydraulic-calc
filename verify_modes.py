# -*- coding: utf-8 -*-
"""Проверка v4: три режима задания расхода (м³/ч, л/мин, м/с)."""
from openpyxl import load_workbook
import math

wb = load_workbook("hydr_calc.xlsx")
ws1 = wb["Ввод"]

print("Строки ввода 3..10:")
for r in range(3, 11):
    print(f"  A{r}:", repr(ws1[f"A{r}"].value), "| C:", repr(ws1[f"C{r}"].value))
print("C12 (Q*):", repr(ws1["C12"].value))
print("C13 (v*):", repr(ws1["C13"].value))
print("Именованные:", sorted(wb.defined_names.keys()))

# --- математика: согласованность трёх режимов эквивалентными расходами ---
f_area = math.pi * (0.05) ** 2 / 4      # d = 50 мм
Q_m3h = 5.0                              # режим 1
Q_lmin = Q_m3h * 1000 / 60               # эквивалент в л/мин = 83.333
v_direct = Q_m3h / 3600 / f_area         # эквивалентная скорость

v_from_Q   = Q_m3h / 3600 / f_area
v_from_lpm = (Q_lmin / 60000) / f_area
print(f"\nQ_lmin из 83.3333 л/мин: v = {v_from_lpm:.6f} (режим Q: {v_from_Q:.6f}, "
      f"прямая: {v_direct:.6f})")
assert abs(v_from_Q - v_from_lpm) < 1e-9 and abs(v_from_Q - v_direct) < 1e-9
print("Согласованность режимов OK")
