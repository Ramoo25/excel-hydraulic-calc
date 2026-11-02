# -*- coding: utf-8 -*-
"""Проверка v2: режим задания расхода/скорости + контрольная математика."""
from openpyxl import load_workbook
import math

wb = load_workbook("hydr_calc.xlsx")
ws1 = wb["Ввод"]
ws3 = wb["Расчёт"]

print("Именованные:", sorted(wb.defined_names.keys()))
for addr in ("C4", "C6", "C7", "C8", "C9", "C12", "C13"):
    print(f"Ввод!{addr}:", repr(ws1[addr].value))
print("Расчёт!A2:", ws3["A2"].value)
print("Расчёт!D5:", ws3["D5"].value)

# --- контрольная математика режима «Скорость (v)»: v = 0.5 м/с ---
print("\n=== Режим «Скорость (v)»: v = 0.5 м/с, d = 50 мм, k = 0.1, L = 100 ===")
d_m, k_mm = 0.05, 0.1
f = math.pi * d_m**2 / 4
v = 0.5
Q = 3600 * v * f
print(f"Q* = {Q:.3f} м³/ч (ожидание 3.534)")
rho, nu = 998.21, 1.0037e-6          # 20 °C
Re = v * d_m / nu
lam = 0.11 * (k_mm / 50.0 + 68 / Re) ** 0.25 if Re > 2300 else 64 / Re
R = lam / d_m * rho * v**2 / 2
print(f"Re = {Re:.0f}, lambda = {lam:.4f}, R = {R:.2f} Па/м, dP = {R*L if False else R*100/1000:.2f} кПа")
L = 100.0
print(f"точный расчёт при L=100: dP = {R*L/1000:.2f} кПа")

# --- контрольная математика режима «Расход»: Q = 5 (как раньше) ---
print("\n=== Режим «Расход»: Q = 5 м³/ч (регрессионный тест) ===")
v2 = 5.0 / 3600 / f
print(f"v* = {v2:.4f} м/с (ожидание 0.7074)")
