# -*- coding: utf-8 -*-
"""Проверка hydr_calc.xlsx: структура + независимый контрольный расчёт."""
from openpyxl import load_workbook
import math

wb = load_workbook("hydr_calc.xlsx")
print("Листы:", wb.sheetnames)
print("Именованные:", list(wb.defined_names.keys()))

ws = wb["Расчёт"]
print("A2:", repr(ws["A2"].value))
for r in (5, 9, 20):
    print(f"--- строка {r} ---")
    for col in "ABCDEFGHIJ":
        print(f"  {col}{r}:", repr(ws[f"{col}{r}"].value))

# --- независимый контрольный расчёт (те же формулы, что в плане) ---
print("\n=== Контрольный расчёт: d=50, k=0.1, Q=5, L=100 ===")
d_m, k_mm, Q, L = 0.05, 0.1, 5.0, 100.0
v = Q / 3600 / (math.pi * d_m**2 / 4)
print(f"v = {v:.4f} м/с")
water = {  # T: (rho, nu)
    5: (999.96, 1.5189e-6), 10: (999.70, 1.3069e-6), 15: (999.10, 1.1394e-6),
    20: (998.21, 1.0037e-6), 25: (997.05, 0.8927e-6), 30: (995.65, 0.8007e-6),
    35: (994.03, 0.7246e-6), 40: (992.22, 0.6579e-6), 45: (990.21, 0.6017e-6),
    50: (988.04, 0.5531e-6), 55: (985.69, 0.5102e-6), 60: (983.19, 0.4740e-6),
    65: (980.55, 0.4434e-6), 70: (977.76, 0.4131e-6), 75: (974.79, 0.3879e-6),
    80: (971.80, 0.3647e-6),
}
print(f"{'T':>4} {'v':>7} {'Re':>8} {'lambda':>8} {'R,Па/м':>9} {'dP,кПа':>8} {'dH,м':>8}")
for t, (rho, nu) in water.items():
    Re = v * d_m / nu
    lam = 64 / Re if Re <= 2300 else 0.11 * (k_mm / 50.0 + 68 / Re) ** 0.25
    R = lam / d_m * rho * v**2 / 2
    dp = R * L / 1000
    dh = R * L / (rho * 9.80665)
    print(f"{t:>4} {v:>7.3f} {Re:>8.0f} {lam:>8.4f} {R:>9.2f} {dp:>8.2f} {dh:>8.3f}")
