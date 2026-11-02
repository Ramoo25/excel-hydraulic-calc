# -*- coding: utf-8 -*-
"""Проверка v3: теплоноситель Вода/ПГ 40%, строки -20..80."""
from openpyxl import load_workbook
import math

wb = load_workbook("hydr_calc.xlsx")
ws1, ws2, ws3 = wb["Ввод"], wb["Свойства воды"], wb["Расчёт"]

print("Именованные:", sorted(wb.defined_names.keys()))
print("Ввод!C3 (теплоноситель):", repr(ws1["C3"].value))
print("Ввод!C4..C9:", [ws1[f"C{n}"].value for n in range(4, 10)])
print("Строк свойств:", ws2.max_row, "Строк расчёта:", ws3.max_row)

# убедимся, что -20..80 заполнены
temps = [ws3.cell(row=r, column=1).value for r in range(5, 26)]
print("T в «Расчёт»:", temps)
for r in range(5, 26):
    b, c = ws3[f"B{r}"].value, ws3[f"C{r}"].value
    print(f"  r={r} T={r-4}: B={b} C={c}")
print("B формула:", ws3["B5"].value, "| C:", ws3["C5"].value, "| E:", ws3["E5"].value)

# --- контрольная математика ПГ 40% при -20 °C (должен быть ламинарный) ---
print("\n=== ПГ 40%, T=-20: nu=46.66e-6, rho=1048 ===")
d_m, k_mm, Q, L = 0.05, 0.1, 5.0, 100.0
v = Q / 3600 / (math.pi * d_m**2 / 4)
nu, rho = 46.66e-6, 1048.0
Re = v * d_m / nu
lam = 64 / Re if Re <= 2300 else 0.11 * (0.002 + 68 / Re) ** 0.25
R = lam / d_m * rho * v**2 / 2
print(f"v={v:.4f} Re={Re:.0f} lambda={lam:.4f} R={R:.1f} Па/м dP={R*L/1000:.2f} кПа")

print("\n=== ПГ 40%, T=+20: nu=4.444e-6, rho=1032.7 ===")
nu, rho = 4.444e-6, 1032.7
Re = v * d_m / nu
lam = 64 / Re if Re <= 2300 else 0.11 * (0.002 + 68 / Re) ** 0.25
R = lam / d_m * rho * v**2 / 2
print(f"Re={Re:.0f} lambda={lam:.4f} R={R:.1f} Па/м dP={R*L/1000:.2f} кПа")

print("\n=== Вода, T=20 (регрессия) ===")
nu, rho = 1.0037e-6, 998.21
Re = v * d_m / nu
lam = 0.11 * (0.002 + 68 / Re) ** 0.25
R = lam / d_m * rho * v**2 / 2
print(f"Re={Re:.0f} lambda={lam:.4f} R={R:.1f} Па/м dP={R*L/1000:.2f} кПа")
