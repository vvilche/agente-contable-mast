#!/usr/bin/env python3
"""Verificación ad-hoc del tablero con pestañas y del master prompt de Copilot.

Comprueba que TODA cifra declarada en el master prompt exista y cuadre contra el
Excel: un prompt que le enseña a Copilot un número equivocado es peor que ninguno.
"""
import json
import re
import sys

import openpyxl

RAG = "/Users/victorvilche/orca/workspaces/Clases Usach/EjercicioClases/rag"
XLSX = "/Users/victorvilche/Documents/EjercicioJose/MASTER - Venta-Sub-GM_Flujos.xlsx"
fails = []


def check(c, m):
    print(("  OK    " if c else "  FALLA ") + m)
    if not c:
        fails.append(m)


# ── 1. El tablero embebió los datos completos ────────────────────────────
print("\n[1] Tablero: datos embebidos")
html = open(f"{RAG}/tablero-pagos-mast.html", encoding="utf-8").read()
m = re.search(r"const PROV = (\[.*?\]), PROY = (\[.*?\]), SEM = (\[.*?\]), CAB = (\{.*?\}), "
              r"VENTAS = (\[.*?\]), CRUCE = (\[.*?\]);", html, re.S)
check(bool(m), "los 6 bloques de datos están en el HTML")
if not m:
    print("\nNo se puede seguir: bloques no encontrados")
    sys.exit(1)
PROV = json.loads(m.group(1))
PROY = json.loads(m.group(2))
SEM = json.loads(m.group(3))
CAB = json.loads(m.group(4))
VENTAS = json.loads(m.group(5))
CRUCE = json.loads(m.group(6))

check(len(PROV) == 663, f"663 pagos a proveedores ({len(PROV)})")
check(len(SEM) == 16, f"16 semanas de flujo ({len(SEM)})")
check(len(VENTAS) == 331, f"331 líneas de venta ({len(VENTAS)})")

# ── 2. Cifras contra el Excel ───────────────────────────────────────────
print("\n[2] Cifras vs Excel")
wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["Subcontratos"]
hdr = {str(c.value).strip(): c.column for c in openpyxl.load_workbook(XLSX)["Subcontratos"][1] if c.value}

bruto_excel = 0.0
pend_excel = 0.0
for r in range(2, ws.max_row + 1):
    p = ws.cell(r, hdr["Proveedor"]).value
    q = ws.cell(r, hdr["Pago Bruto"]).value
    e = str(ws.cell(r, hdr["Estado del pago"]).value or "").strip().lower()
    if not p or str(p).upper().startswith("TOTAL") or not isinstance(q, (int, float)):
        continue
    bruto_excel += q
    if e == "sin pago":
        pend_excel += q

bruto_tab = sum(p["bruto"] or 0 for p in PROV)
pend_tab = sum(p["bruto"] or 0 for p in PROV if p["estado_pago"].lower() == "sin pago")
check(abs(bruto_excel - bruto_tab) < 1, f"bruto total ${bruto_tab:,.0f} = Excel ${bruto_excel:,.0f}")
check(abs(pend_excel - pend_tab) < 1, f"pendiente ${pend_tab:,.0f} = Excel ${pend_excel:,.0f}")

# proveedores: 89 = 88 con RUT + 1 sin RUT ("Cristian Auger A"), agrupado por nombre
ruts = {p["rut"] for p in PROV if p["rut"]}
sin_rut = {p["proveedor"] for p in PROV if not p["rut"]}
check(len(ruts) == 88, f"88 proveedores con RUT ({len(ruts)})")
check(len(sin_rut) == 1, f"1 proveedor sin RUT ({len(sin_rut)}): {sorted(sin_rut)}")
check(len(ruts) + len(sin_rut) == 89, f"total 89 proveedores ({len(ruts) + len(sin_rut)})")

# ── 3. Flujo de caja contra las celdas del Excel ────────────────────────
print("\n[3] Flujo de caja vs celdas")
wf = wb["Flujo de caja"]
check(abs((CAB["ingresos_16"] or 0) - wf.cell(4, 6).value) < 1, "ingresos 16 semanas = F4")
check(abs((CAB["egresos_16"] or 0) - wf.cell(4, 9).value) < 1, "egresos 16 semanas = I4")
check(abs((CAB["neto_16"] or 0) - wf.cell(4, 12).value) < 1, "flujo neto = L4")
check(abs((CAB["saldo_final"] or 0) - wf.cell(4, 15).value) < 1, "saldo final proyectado = O4")
check(abs((CAB["saldo_inicial"] or 0) - wf.cell(5, 2).value) < 1, "saldo bancario inicial = B5")
check(abs((CAB["saldo_final"] or 0) - (SEM[-1]["saldo_final"] or 0)) < 1,
      "saldo final del panel = última semana de la tabla")

# ── 4. El cruce: 12 de 31 y el conteo de proveedores ───────────────────
print("\n[4] Cruce pagos ↔ facturación")
con_pagos = [p for p in CRUCE if p["n_pagos"] > 0]
cruzan = [p for p in con_pagos if p["cruza"]]
check(len(con_pagos) == 31, f"31 proyectos con egresos ({len(con_pagos)})")
check(len(cruzan) == 12, f"12 cruzan con facturación ({len(cruzan)})")

# ── 5. El master prompt no afirma nada que no esté medido ──────────────
print("\n[5] Master prompt contra los datos")
mp = open(f"{RAG}/08-master-prompt-copilot.md", encoding="utf-8").read()
check("$89.675.467" in mp or "$89,7 millones" in mp, "declara el pendiente correcto")
check("−$56.122.693" in mp or "-$56.122.693" in mp, "declara el saldo final correcto")
check("$56.122.694" in mp, "declara la necesidad de caja correcta")
check("solo 12 de 31" in mp, "declara la limitación del cruce")
check("89" in mp and "91" in mp, "explica 89 por RUT vs 91 por nombre")
check("no está en la fuente" in mp, "instruye responder 'no está'")

# los números del prompt deben cuadrar con lo medido
check(f"{round(pend_tab):,}".replace(",", ".") in mp, "el pendiente del prompt = el medido")
check(f"{round(abs(CAB['saldo_final'])):,}".replace(",", ".") in mp, "el saldo final del prompt = el medido")

print("\n" + "=" * 54)
if fails:
    print(f"{len(fails)} FALLA(S):")
    for f in fails:
        print("  -", f)
    sys.exit(1)
print("TODO OK")
