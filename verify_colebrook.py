# -*- coding: utf-8 -*-
"""Проверка колонок Колбрука: формулы в файле + независимый расчёт сходимости."""
from openpyxl import load_workbook
import math

wb = load_workbook("hydr_calc.xlsx")
ws = wb["Расчёт"]

print("Заголовки (строка 3):")
for c in range(11, 15):
    print(f"  {ws.cell(row=3, column=c).coordinate}: {ws.cell(row=3, column=c).value}")
print("\nK5:", ws["K5"].value)
print("L5:", ws["L5"].value)
print("M5:", ws["M5"].value)
print("N5:", ws["N5"].value)

# --- независимая проверка сходимости итераций, d=50, k=0.1, Q=5, T=20 C ---
d_m, k_mm, Q, L = 0.05, 0.1, 5.0, 100.0
v = Q / 3600 / (math.pi * d_m**2 / 4)
rho, nu = 998.21, 1.0037e-6
Re = v * d_m / nu
lam_a = 0.11 * (k_mm / 50.0 + 68 / Re) ** 0.25

def cw(lam_prev):
    X = k_mm / (3.7 * 50.0) + 2.51 / (Re * math.sqrt(lam_prev))
    return (1 / (-2 * math.log10(X))) ** 2

it1, it2, it3 = cw(lam_a), cw(cw(lam_a)), cw(cw(cw(lam_a)))
print(f"\nRe = {Re:.0f}")
print(f"lambda Альтшуль:      {lam_a:.5f}")
print(f"Колбрук ит.1:         {it1:.5f}")
print(f"Колбрук ит.2:         {it2:.5f}")
print(f"Колбрук ит.3:         {it3:.5f}")
print(f"Разность (Альт-CW3):  {(lam_a - it3) / it3 * 100:+.2f}%")

# точное решение по одной Коltбруку после многих итераций
lam = lam_a
for _ in range(50):
    lam = cw(lam)
print(f"Колбрук сходящийся:   {lam:.5f}")
print(f"ит.3 vs точное:       {(it3 - lam) / lam * 100:+.4f}%")
