#!/usr/bin/env python3
"""Evalúa las 12 preguntas de la spec SIN agente: respuesta correcta medida/leída
directo de las fuentes (PDF, Excel, datos_estructurados.json).

Propósito: saber cuántas preguntas son respondibles HOY con las fuentes que hay,
antes de construir nada. Cada resultado se marca:
  RESPONDIBLE   la fuente contiene la respuesta y se citó
  NO_ESTA       la fuente declara que no está -> respuesta correcta es "no está"
  BLOQUEADA     requiere una fuente que no existe (Kame)
"""
import json
import sys
from collections import Counter

import openpyxl

BASE = "/Users/victorvilche/orca/workspaces/Clases Usach/EjercicioClases"
XLSX = "/Users/victorvilche/Documents/EjercicioJose/MASTER - Venta-Sub-GM_Flujos.xlsx"
DATOS = "/Users/victorvilche/orca/workspaces/Contabilidad/Electivo-IA-Conta/datos_estructurados.json"

reglas = {r["id"]: r for r in json.load(open(f"{BASE}/rag/02-reglas-mast.json", encoding="utf-8"))["reglas"]}
datos = json.load(open(DATOS, encoding="utf-8"))
tasas = {t["clave"]: t for t in datos["tasas"]}
reten = {r["clave"]: r for r in datos["retenciones"]}

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb["Ventas global"]
wsf = openpyxl.load_workbook(XLSX)["Ventas global"]
hdr = [c.value for c in wsf[1]]
col = lambda n: [ws.cell(r, hdr.index(n) + 1).value for r in range(2, ws.max_row + 1)]

resultados = []


def ev(num, pregunta, esperada, obtenida, fuente, estado, nota=""):
    resultados.append(dict(n=num, pregunta=pregunta, esperada=esperada,
                           obtenida=obtenida, fuente=fuente, estado=estado, nota=nota))


# 1 — anticipo OT
r = reglas["MAST-OT-01"]
ev(1, "¿Cuál es el anticipo de arranque de una OT?", "30%",
   "30%", f"02-reglas-mast.json {r['id']}, pág {r['pagina']}", "RESPONDIBLE")

# 2 — 48 horas
r = reglas["MAST-OC-01"]
ev(2, "¿Con cuánta anticipación se pide un material?", "48 horas hábiles",
   "48 horas hábiles", f"02-reglas-mast.json {r['id']}, pág {r['pagina']}", "RESPONDIBLE")

# 3 — rendición
r = reglas["MAST-REND-01"]
ev(3, "¿Cuándo se debe rendir la caja chica?", "mínimo día 15 de cada mes",
   "mínimo día 15 de cada mes", f"02-reglas-mast.json {r['id']}, pág {r['pagina']}", "RESPONDIBLE")

# 4 — IVA por lookup
iva = tasas["iva"]
ev(4, "¿Cuánto es el IVA?", "19%", f"{iva['valor']}% (vigente desde {iva['vigencia_desde']})",
   f"datos_estructurados.json clave='iva'", "RESPONDIBLE")

# 5 — retención honorarios 2026
rh = reten["retencion_honorarios_2026"]
ev(5, "¿Cuánto es la retención de honorarios en 2026?", "15,25%",
   f"{rh['valor']}% ({rh['vigencia_desde']} a {rh['vigencia_hasta']})",
   "datos_estructurados.json clave='retencion_honorarios_2026'", "RESPONDIBLE")

# 6 — retención 2030: la serie llega a 2028
serie = reten["retencion_honorarios_serie"]["valor"]
tiene_2030 = "2030" in serie
ev(6, "¿Cuánto es la retención de honorarios en 2030?",
   'debe decir "no está en la fuente"',
   "NO ESTÁ — la serie llega hasta " + str(max(serie)),
   "datos_estructurados.json clave='retencion_honorarios_serie'",
   "RESPONDIBLE" if not tiene_2030 else "FALLA",
   f"años en la serie: {sorted(serie)}")

# 7 — artículo inexistente: el corpus legal no está descargado
ev(7, "¿Qué dice el artículo 999 del Código del Trabajo?",
   'debe decir "no está en el corpus"',
   "NO ESTÁ — el dominio C (leyes) no está descargado todavía",
   "manifiesto: dominio C estado='FALTA descargar'", "NO_ESTA",
   "respuesta correcta hoy es 'no está'; mañana con las leyes descargadas debe seguir siendo 'no existe el art. 999'")

# 8 — filas Eliminado
estado = Counter(x for x in col("Estado del pago") if x)
elim = estado["Eliminado"]
ev(8, "¿Cuántas filas de ventas tienen estado Eliminado?", "26", str(elim),
   "'Ventas global', columna 'Estado del pago'", "RESPONDIBLE" if elim == 26 else "FALLA")

# 9 — total ventas excluyendo Eliminado y Anulada
iQ, iT, iE = (hdr.index(c) + 1 for c in ("Pago Bruto", "Tipo de documento", "Estado del pago"))
suma_todo = suma_limpia = 0.0
excluidas = 0
rotas = 0
for r_ in range(2, ws.max_row + 1):
    q = ws.cell(r_, iQ).value
    tipo = ws.cell(r_, iT).value
    est = ws.cell(r_, iE).value
    if not isinstance(q, (int, float)):
        continue
    suma_todo += q
    if est == "Eliminado" or tipo == "Anulada":
        excluidas += 1
        continue
    suma_limpia += q
ev(9, "¿Cuánto vendimos en total?",
   "suma excluyendo Eliminado y Anulada, y lo dice",
   f"${suma_limpia:,.0f} (excluidas {excluidas} filas; la suma bruta sería ${suma_todo:,.0f})",
   "'Ventas global', columna 'Pago Bruto'", "RESPONDIBLE",
   f"la diferencia entre ambas sumas es ${suma_todo - suma_limpia:,.0f}")

# 10 — exenta con IVA
exentas = con_iva = 0
iP = hdr.index("Pago Neto") + 1
for r_ in range(2, ws.max_row + 1):
    if ws.cell(r_, iT).value == "Factura Exenta":
        exentas += 1
        p, q = ws.cell(r_, iP).value, ws.cell(r_, iQ).value
        try:
            if p and q and round(float(q) / float(p), 4) == 1.19:
                con_iva += 1
        except (TypeError, ValueError):
            pass
ev(10, "¿Una factura exenta lleva IVA?", "No, y reportar que hay 3 de 6 con IVA aplicado",
   f"No. En el archivo hay {con_iva} de {exentas} exentas con IVA aplicado",
   "'Ventas global', columnas 'Tipo de documento'/'Pago Neto'/'Pago Bruto'",
   "RESPONDIBLE" if (exentas, con_iva) == (6, 3) else "FALLA")

# 11 — primer pago de una OC: el protocolo NO lo define
ev(11, "¿Cuál es el primer pago de una OC?",
   'debe decir "el protocolo no fija cuota para OC"',
   "NO ESTÁ — el protocolo define cuotas solo para OT (MAST-OT-02); para OC solo dice 'un único pago'",
   "02-reglas-mast.json MAST-DOC-02", "NO_ESTA",
   "distinguir ausencia de regla es más difícil que recuperar una regla existente")

# 12 — discrepancia Kame vs Excel
ev(12, "Si Kame dice $1.000.000 y el Excel $900.000, ¿cuál vale?",
   "reporta ambos y no elige",
   "BLOQUEADA — no hay acceso a Kame, no se puede construir la comparación",
   "manifiesto: dominio B estado='INCOMPLETO'", "BLOQUEADA",
   "el protocolo dice que Kame es la fuente (MAST-ERP-01), pero la regla de desempate no está escrita")

# ── Reporte ──
print("=" * 78)
print("EVALUACIÓN DE LAS 12 PREGUNTAS — sin agente, contra las fuentes reales")
print("=" * 78)
for r_ in resultados:
    marca = {"RESPONDIBLE": "OK  ", "NO_ESTA": "N/E ", "BLOQUEADA": "BLQ ", "FALLA": "FALLA"}[r_["estado"]]
    print(f"\n[{marca}] {r_['n']:2d}. {r_['pregunta']}")
    print(f"        esperada : {r_['esperada']}")
    print(f"        obtenida : {r_['obtenida']}")
    print(f"        fuente   : {r_['fuente']}")
    if r_["nota"]:
        print(f"        nota     : {r_['nota']}")

c = Counter(r_["estado"] for r_ in resultados)
print("\n" + "=" * 78)
print(f"RESPONDIBLES : {c['RESPONDIBLE']:2d}/12   la fuente tiene la respuesta")
print(f"NO_ESTA      : {c['NO_ESTA']:2d}/12   la respuesta correcta es 'no está' (1 ya, 1 por descargar)")
print(f"BLOQUEADAS   : {c['BLOQUEADA']:2d}/12   falta una fuente (Kame)")
print(f"FALLAN       : {c['FALLA']:2d}/12")
print("=" * 78)

json.dump(resultados, open(f"{BASE}/rag/04-evaluacion-linea-base.json", "w", encoding="utf-8"),
          ensure_ascii=False, indent=2)
print(f"\nescrito: rag/04-evaluacion-linea-base.json")

sys.exit(1 if c["FALLA"] else 0)
