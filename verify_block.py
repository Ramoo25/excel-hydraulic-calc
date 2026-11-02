# -*- coding: utf-8 -*-
"""Проверка v5: блок оценки тёплого пола/снеготаяния + контрольная математика."""
from openpyxl import load_workbook
import math

wb = load_workbook("hydr_calc.xlsx")
ws = wb["Ввод"]

print("Блок на «Ввод»:")
for r in range(23, 36):
    print(f"  A{r}:", repr(ws[f"A{r}"].value), "| C:", repr(ws[f"C{r}"].value))

ws2 = wb["Свойства воды"]
print("\nСвойства (T, cp воды, cp ПГ) строки 5,10,14,25:")
for r in (5, 10, 14, 25):
    print(f"  r={r}:", [ws2.cell(row=r, column=c).value for c in (1, 8, 9)])

# --- контрольная математика ---
print("\n=== Вода: Q=5 м³/ч, dT=10, Tср=35, S=100/0.2=500 м² ===")
cp = 4178.8 + 0.01016 * (35 - 37.5) ** 2   # min строк точной таблицы ~35 °C
rho = 994.03                                # свойства при 35 °C
Q_m3h, dT = 5.0, 10.0
Phi = cp * rho * (Q_m3h / 3600) * dT
S = 100 / (200 / 1000)
print(f"cp={cp:.1f} rho={rho} Ф={Phi:.1f} Вт ({Phi/1000:.2f} кВт), "
      f"S={S} м²  →  {Phi/S:.2f} Вт/м²  →  {Phi/100:.2f} Вт/м трубы")

print("\n=== ПГ 40%, те же параметры ===")
cp_pg = 3560 + 2.8 * 35
rho_pg = 1024.6                              # свойства ПГ при 35 °C
Phi_pg = cp_pg * rho_pg * (Q_m3h / 3600) * dT
print(f"cp={cp_pg:.1f} rho={rho_pg} Ф={Phi_pg:.1f} Вт ({Phi_pg/1000:.2f} кВт), "
      f"→ {Phi_pg/S:.2f} Вт/м²  →  {Phi_pg/100:.2f} Вт/м трубы")
