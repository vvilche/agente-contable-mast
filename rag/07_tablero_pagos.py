#!/usr/bin/env python3
"""Genera el tablero de gestión de pagos a proveedores de MAST (HTML con pestañas).

Pestañas:
  1. Proveedores  — cuánto se debe, a quién, en qué proyecto
  2. Proyectos    — cuánto se debe por proyecto, cruce pagos vs facturación
  3. Facturación  — si se facturó al cliente final y qué hito se facturó
  4. Flujo de caja — ingresos, egresos, saldo proyectado
  + Tareas       — qué es cada cosa y qué hay que corregir

Los datos van embebidos: se abre con doble clic, sin servidor.
NO publicar: contiene montos, RUT y clientes de MAST.
"""
import json
import os
import re
import unicodedata
from collections import Counter, defaultdict

import openpyxl

XLSX = "/Users/victorvilche/Documents/EjercicioJose/MASTER - Venta-Sub-GM_Flujos.xlsx"
BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "tablero-pagos-mast.html")


def norm(s):
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9 ]", " ", s.lower())
    return " ".join(s.split())


def num(v):
    return round(float(v), 2) if isinstance(v, (int, float)) else None


def leer():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    wbv = openpyxl.load_workbook(XLSX)

    def headers(hoja):
        return {str(c.value).strip(): c.column for c in wbv[hoja][1] if c.value is not None}

    def filas(hoja, col_proyecto):
        ws, h = wb[hoja], headers(hoja)
        out = []
        for r in range(2, ws.max_row + 1):
            proy = ws.cell(r, h[col_proyecto]).value
            if not proy or str(proy).strip() == "":
                continue
            if str(proy).strip().upper().startswith(("TOTAL", "SUBTOTAL")):
                continue
            out.append((r, h, ws))
        return out, ws, h

    # ── proveedores (subcontratos) ──────────────────────────────────────
    ws_s, h_s = wb["Subcontratos"], headers("Subcontratos")
    prov_filas = []
    for r in range(2, ws_s.max_row + 1):
        p = ws_s.cell(r, h_s["Proveedor"]).value
        if not p or not str(p).strip():
            continue
        if str(p).strip().upper().startswith(("TOTAL", "SUBTOTAL")):
            continue
        prov_filas.append({
            "fila": r,
            "proyecto": str(ws_s.cell(r, h_s["Proyecto"]).value or "").strip(),
            "proyecto_norm": norm(ws_s.cell(r, h_s["Proyecto"]).value or ""),
            "proveedor": " ".join(str(p).split()),
            "rut": str(ws_s.cell(r, h_s["Rut"]).value or "").strip(),
            "ot": str(ws_s.cell(r, h_s["N° OT"]).value or "").strip(),
            "tipo": str(ws_s.cell(r, h_s["Tipo Orden"]).value or "").strip(),
            "tipo_doc": str(ws_s.cell(r, h_s["Tipo Documento"]).value or "").strip(),
            "concepto": str(ws_s.cell(r, h_s["Concepto"]).value or "").strip(),
            "hito": str(ws_s.cell(r, h_s["Hito de pago"]).value or "").strip(),
            "estado_fact": str(ws_s.cell(r, h_s["Estado facturación"]).value or "").strip(),
            "estado_pago": str(ws_s.cell(r, h_s["Estado del pago"]).value or "").strip(),
            "factura": str(ws_s.cell(r, h_s["N° de factura"]).value or "").strip(),
            "neto": num(ws_s.cell(r, h_s["Pago Neto"]).value),
            "bruto": num(ws_s.cell(r, h_s["Pago Bruto"]).value),
            "fecha_pago": str(ws_s.cell(r, h_s["Fecha de pago"]).value or "")[:10],
            "proyeccion": str(ws_s.cell(r, h_s["Proyección de pago"]).value or "")[:10],
        })

    # ── ventas (facturación al cliente final) ───────────────────────────
    ws_v, h_v = wb["Ventas global"], headers("Ventas global")
    ventas = []
    for r in range(2, ws_v.max_row + 1):
        proy = ws_v.cell(r, h_v["Proyecto"]).value
        if not proy or not str(proy).strip():
            continue
        ventas.append({
            "fila": r,
            "proyecto": str(proy).strip(),
            "proyecto_norm": norm(proy),
            "cliente": str(ws_v.cell(r, h_v["Cliente"]).value or "").strip(),
            "rut": str(ws_v.cell(r, h_v["Rut"]).value or "").strip(),
            "tipo_doc": str(ws_v.cell(r, h_v["Tipo de documento"]).value or "").strip(),
            "folio": str(ws_v.cell(r, h_v["Folio"]).value or "").strip(),
            "hito": str(ws_v.cell(r, h_v["Hito de pago"]).value or "").strip(),
            "estado_fact": str(ws_v.cell(r, h_v["Estado facturación"]).value or "").strip(),
            "estado_pago": str(ws_v.cell(r, h_v["Estado del pago"]).value or "").strip(),
            "neto": num(ws_v.cell(r, h_v["Pago Neto"]).value),
            "bruto": num(ws_v.cell(r, h_v["Pago Bruto"]).value),
            "total_neto": num(ws_v.cell(r, h_v["TOTAL NETO"]).value),
            "fecha_emision": str(ws_v.cell(r, h_v["Fecha emisión"]).value or "")[:10],
            "proyeccion": str(ws_v.cell(r, h_v["Proyección de pago"]).value or "")[:10],
            "centro": str(ws_v.cell(r, h_v["Centro de Costos"]).value or "").strip(),
        })

    # ── flujo de caja ───────────────────────────────────────────────────
    # OJO: esta hoja NO es una tabla con encabezado en fila 1. Es un panel con
    # pares rótulo→valor repartidos en varias columnas (A4='Fecha inicial', B4=valor).
    ws_f, h_f = wb["Flujo de caja"], headers("Flujo de caja")

    def etiqueta_a_valor(rotulo):
        """Busca el rótulo en las filas 1-9 y devuelve su valor.

        En este panel el valor casi nunca está pegado al rótulo: va 2 columnas a la
        derecha (D→F, J→L, M→O). Se prueba +2 y, si es nulo, +1. Medido sobre la
        hoja real: 'Ingresos dentro de las 16 semanas' está en D4 y su valor en F4.
        """
        for r in range(1, 10):
            for c in range(1, ws_f.max_column + 1):
                v = ws_f.cell(r, c).value
                if isinstance(v, str) and v.strip() == rotulo:
                    for salto in (2, 1):
                        if c + salto <= ws_f.max_column:
                            val = ws_f.cell(r, c + salto).value
                            if val is not None:
                                return val
                    return None
        return None

    cabecera = {
        "fecha_inicial": str(etiqueta_a_valor("Fecha inicial del flujo") or "")[:10],
        "saldo_inicial": num(etiqueta_a_valor("Saldo bancario inicial")),
        "ingresos_16": num(etiqueta_a_valor("Ingresos dentro de las 16 semanas")),
        "egresos_16": num(etiqueta_a_valor("Egresos dentro de las 16 semanas")),
        "neto_16": num(etiqueta_a_valor("Flujo neto de las 16 semanas")),
        "subcontratos_16": num(etiqueta_a_valor("Subcontratos dentro de las 16 semanas")),
        "gastos_16": num(etiqueta_a_valor("Gastos mensuales dentro de las 16 semanas")),
        "saldo_final": num(etiqueta_a_valor("Saldo final proyectado")),
        "necesidad_caja": num(etiqueta_a_valor("Mayor necesidad de caja")),
        "actualizacion": str(etiqueta_a_valor("Actualización") or "")[:10],
        "semanas": etiqueta_a_valor("Semanas a mostrar"),
    }

    # la tabla semanal sí tiene encabezado; se localiza su fila por la columna 'N°'
    fila_hdr = None
    for r in range(1, 20):
        for c in range(1, 6):
            if str(ws_f.cell(r, c).value).strip() == "N°":
                fila_hdr = r
                break
        if fila_hdr:
            break
    h_f = {}
    if fila_hdr:
        for c in range(1, ws_f.max_column + 1):
            v = ws_f.cell(fila_hdr, c).value
            if v is not None:
                h_f[str(v).strip()] = c

    semanas = []
    if fila_hdr:
        for r in range(fila_hdr + 1, ws_f.max_row + 1):
            n = ws_f.cell(r, h_f["N°"]).value
            if not isinstance(n, int):
                continue
            semanas.append({
                "n": n,
                "etiqueta": str(ws_f.cell(r, h_f["Mes y semana"]).value or ""),
                "inicio": str(ws_f.cell(r, h_f["Inicio"]).value or "")[:10],
                "fin": str(ws_f.cell(r, h_f["Fin"]).value or "")[:10],
                "saldo_inicial": num(ws_f.cell(r, h_f["Saldo inicial"]).value),
                "ingresos": num(ws_f.cell(r, h_f["Ingresos programados"]).value),
                "subcontratos": num(ws_f.cell(r, h_f["Subcontratos"]).value),
                "gastos": num(ws_f.cell(r, h_f["Gastos mensuales"]).value),
                "obras": num(ws_f.cell(r, h_f["Costos de obras"]).value),
                "egresos": num(ws_f.cell(r, h_f["Egresos totales"]).value) if "Egresos totales" in h_f else None,
                "flujo_neto": num(ws_f.cell(r, h_f["Flujo neto (Ingresos-Egresos Semanal)"]).value)
                              if "Flujo neto (Ingresos-Egresos Semanal)" in h_f else None,
                "saldo_final": num(ws_f.cell(r, h_f["Saldo final"]).value) if "Saldo final" in h_f else None,
                "menor_saldo": num(ws_f.cell(r, h_f["Menor saldo diario"]).value) if "Menor saldo diario" in h_f else None,
                "estado": str(ws_f.cell(r, h_f["Estado de caja"]).value or "") if "Estado de caja" in h_f else "",
                "faltante": num(ws_f.cell(r, h_f["Monto faltante"]).value) if "Monto faltante" in h_f else None,
            })

    return prov_filas, ventas, semanas, cabecera


def cruzar(prov_filas, ventas):
    """Cruza pagos a proveedores con facturación al cliente final.

    LLAVE: nombre de proyecto normalizado. Medido: solo ~12 de 31 proyectos de
    subcontratos existen en Ventas. Donde no cruza, se declara, no se inventa.
    """
    vta_por_proy = defaultdict(list)
    for v in ventas:
        vta_por_proy[v["proyecto_norm"]].append(v)

    proyectos = defaultdict(lambda: {"pagos": [], "facturas": [], "nombre": ""})
    for p in prov_filas:
        k = p["proyecto_norm"]
        proyectos[k]["pagos"].append(p)
        if not proyectos[k]["nombre"]:
            proyectos[k]["nombre"] = p["proyecto"]
    for k, vs in vta_por_proy.items():
        proyectos[k]["facturas"].extend(vs)
        if not proyectos[k]["nombre"]:
            proyectos[k]["nombre"] = vs[0]["proyecto"]

    out = []
    for k, d in proyectos.items():
        pagos = d["pagos"]
        fac = d["facturas"]
        # egresos con proveedores, separando lo ya pagado de lo pendiente
        eg_bruto = sum(x["bruto"] or 0 for x in pagos)
        eg_pagado = sum(x["bruto"] or 0 for x in pagos if x["estado_pago"].lower() == "pagado")
        eg_pend = sum(x["bruto"] or 0 for x in pagos if x["estado_pago"].lower() == "sin pago")
        eg_elim = sum(x["bruto"] or 0 for x in pagos if x["estado_pago"].lower().startswith("elimin"))
        # facturado al cliente: solo documentos válidos (no anulada), y el neto
        fac_ok = [x for x in fac if x["tipo_doc"].lower() not in ("anulada",)]
        fac_neto = sum(x["neto"] or 0 for x in fac_ok)
        fac_bruto = sum(x["bruto"] or 0 for x in fac_ok)
        sin_facturar = [x for x in fac if x["estado_fact"] == "Sin factura"]
        hitos = Counter(x["hito"] for x in fac_ok if x["hito"])
        out.append({
            "proyecto": d["nombre"],
            "norm": k,
            "cruza": bool(pagos) and bool(fac),
            "n_pagos": len(pagos),
            "n_facturas": len(fac_ok),
            "egreso_bruto": round(eg_bruto),
            "egreso_pagado": round(eg_pagado),
            "egreso_pendiente": round(eg_pend),
            "egreso_eliminado": round(eg_elim),
            "facturado_neto": round(fac_neto),
            "facturado_bruto": round(fac_bruto),
            "margen_bruto": round(fac_neto - eg_bruto),
            "hitos": [{"hito": h, "n": n, "monto": round(sum(x["neto"] or 0 for x in fac_ok if x["hito"] == h))}
                      for h, n in hitos.most_common()],
            "sin_facturar": len(sin_facturar),
            "clientes": sorted({x["cliente"] for x in fac_ok if x["cliente"]}),
            "proveedores": sorted({x["proveedor"] for x in pagos}),
        })
    out.sort(key=lambda x: -max(x["egreso_bruto"], x["facturado_neto"]))
    return out


HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>MAST · Gestión de pagos a proveedores</title>
<style>
:root{--bg:#0e1116;--panel:#161b22;--panel2:#1c232c;--line:#2d3742;--txt:#e6edf3;--dim:#8b949e;
 --acc:#4da3ff;--ok:#3fb950;--warn:#d29922;--bad:#f85149;--vio:#a371f7;
 --mono:ui-monospace,SFMono-Regular,Menlo,monospace}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--txt);font:14.5px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;padding:22px 20px 70px;max-width:1400px;margin:0 auto}
h1{font-size:22px;font-weight:650;letter-spacing:-.02em}
.sub{color:var(--dim);font-size:12.5px;margin-top:4px}
.eyebrow{font:600 10.5px/1 var(--mono);letter-spacing:.12em;text-transform:uppercase;color:var(--dim);margin-bottom:6px}
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--line);margin:20px 0 0;flex-wrap:wrap}
.tab{padding:9px 15px;cursor:pointer;font-size:13.5px;font-weight:550;color:var(--dim);
 border-bottom:2px solid transparent;border-radius:7px 7px 0 0;user-select:none}
.tab:hover{color:var(--txt);background:var(--panel)}
.tab.on{color:var(--acc);border-bottom-color:var(--acc);background:var(--panel)}
.vista{display:none;padding-top:18px}
.vista.on{display:block}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:9px;margin-bottom:14px}
.kpi{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:10px 12px}
.kpi b{display:block;font:650 19px/1.15 var(--mono);color:var(--acc)}
.kpi.r b{color:var(--bad)}.kpi.g b{color:var(--ok)}.kpi.w b{color:var(--warn)}
.kpi span{font-size:11px;color:var(--dim);display:block;margin-top:2px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:9px;padding:14px 16px;margin-bottom:12px}
.card h3{font-size:13px;font-weight:620;color:var(--acc);margin-bottom:9px;text-transform:uppercase;letter-spacing:.06em;font-family:var(--mono)}
input[type=search]{width:100%;background:var(--panel);border:1px solid var(--line);color:var(--txt);
 padding:9px 12px;border-radius:7px;font-size:13.5px;outline:none;margin-bottom:10px}
input[type=search]:focus{border-color:var(--acc)}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{text-align:left;padding:7px 8px;border-bottom:1px solid var(--line);color:var(--dim);
 font:600 10px/1 var(--mono);letter-spacing:.07em;text-transform:uppercase;position:sticky;top:0;background:var(--panel);z-index:1;cursor:pointer}
td{padding:6px 8px;border-bottom:1px solid #21262d;vertical-align:top}
tbody tr:hover td{background:var(--panel2)}
.num{font-family:var(--mono);text-align:right;white-space:nowrap}
.pos{color:var(--ok)}.neg{color:var(--bad)}
.scroll{max-height:460px;overflow-y:auto}
.tag{font:600 9.5px/1 var(--mono);padding:3px 6px;border-radius:4px;text-transform:uppercase;white-space:nowrap}
.t-ok{background:#1a3d24;color:#56d364}.t-w{background:#3d2a1a;color:#d29922}
.t-b{background:#3d1f24;color:#ff7b72}.t-x{background:#21262d;color:#8b949e}
.t-v{background:#2d2340;color:#c9a4ff}
.bar{height:6px;background:#21262d;border-radius:3px;overflow:hidden;display:flex}
.bar i{display:block;height:100%}
.clip{max-width:230px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.clip:hover{overflow:visible;white-space:normal}
.empty{color:var(--dim);padding:22px;text-align:center;font-size:13px}
.nota{background:#1a2430;border-left:3px solid var(--acc);padding:9px 12px;border-radius:0 6px 6px 0;font-size:12.5px;margin:9px 0;color:#c9d1d9}
.aviso{background:#2a2116;border-left:3px solid var(--warn);padding:9px 12px;border-radius:0 6px 6px 0;font-size:12.5px;margin:9px 0}
.crit{background:#2a1719;border-left:3px solid var(--bad);padding:9px 12px;border-radius:0 6px 6px 0;font-size:12.5px;margin:9px 0}
p{margin:6px 0;color:#c9d1d9}
ul{margin:6px 0 6px 19px}li{margin:4px 0;color:#c9d1d9;font-size:13px}
code{font:12px var(--mono);background:#21262d;padding:1px 5px;border-radius:3px;color:#a5d6ff}
.chart{display:flex;align-items:flex-end;gap:5px;height:130px;margin-top:10px;padding-top:6px}
.chart .b{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;gap:2px;min-width:0}
.chart .b i{width:100%;border-radius:2px 2px 0 0;display:block}
.chart .b small{font:9px var(--mono);color:var(--dim);white-space:nowrap;transform:rotate(-45deg);transform-origin:center;margin-top:22px}
</style>
</head>
<body>
<div class="eyebrow">MAST Arquitectura y Construcción · master de José</div>
<h1>Gestión de pagos a proveedores</h1>
<p class="sub" id="pie"></p>

<div class="tabs" id="tabs"></div>

<!-- ═══════ 1 · PROVEEDORES ═══════ -->
<div class="vista on" id="v-proveedores">
  <div class="kpis" id="k-prov"></div>
  <div class="card">
    <h3>Cuánto se debe, a quién y en qué proyecto</h3>
    <input type="search" id="q-prov" placeholder="Buscar proveedor, RUT, proyecto, OT…">
    <div class="scroll"><table id="t-prov">
      <thead><tr>
        <th>Proveedor</th><th>RUT</th><th>Proyecto</th><th>OT</th><th>Hito</th>
        <th>Estado</th><th class="num">Debe (bruto)</th><th class="num">Pagado</th>
        <th class="num">Pendiente</th><th>Doc</th><th>Factura</th>
      </tr></thead><tbody></tbody>
    </table></div>
  </div>
</div>

<!-- ═══════ 2 · PROYECTOS ═══════ -->
<div class="vista" id="v-proyectos">
  <div class="kpis" id="k-proy"></div>
  <div class="card">
    <h3>Por proyecto: egresos con proveedores vs facturación al cliente</h3>
    <input type="search" id="q-proy" placeholder="Buscar proyecto…">
    <div class="scroll"><table id="t-proy">
      <thead><tr>
        <th>Proyecto</th><th class="num">Egresos</th><th class="num">Pagado</th>
        <th class="num">Pendiente</th><th class="num">Facturado (neto)</th>
        <th class="num">Margen</th><th>Cruce</th><th class="num">Proveedores</th>
      </tr></thead><tbody></tbody>
    </table></div>
  </div>
</div>

<!-- ═══════ 3 · FACTURACIÓN ═══════ -->
<div class="vista" id="v-facturacion">
  <div class="kpis" id="k-fac"></div>
  <div class="card">
    <h3>Qué se le facturó al cliente final — por hito de pago</h3>
    <div class="nota">El hito es <strong>qué se está cobrando</strong>: anticipo, estado de pago (EEPP), contra entrega o saldo.
    Un hito sin factura es trabajo hecho que aún no se cobró.</div>
    <div class="scroll"><table id="t-hitos">
      <thead><tr><th>Hito de pago</th><th class="num">Documentos</th>
      <th class="num">Facturado neto</th><th class="num">Facturado bruto</th><th class="num">Sin facturar</th></tr></thead>
      <tbody></tbody></table></div>
  </div>
  <div class="card">
    <h3>Detalle de facturación por proyecto y cliente</h3>
    <input type="search" id="q-fac" placeholder="Buscar proyecto, cliente, folio…">
    <div class="scroll"><table id="t-fac">
      <thead><tr><th>Proyecto</th><th>Cliente</th><th class="num">Doc.</th>
      <th class="num">Facturado neto</th><th class="num">Egresos</th><th>Hitos facturados</th></tr></thead>
      <tbody></tbody></table></div>
  </div>
</div>

<!-- ═══════ 4 · FLUJO DE CAJA ═══════ -->
<div class="vista" id="v-flujo">
  <div class="kpis" id="k-flujo"></div>
  <div class="card">
    <h3>Saldo final proyectado por semana</h3>
    <div class="chart" id="chart"></div>
  </div>
  <div class="card">
    <h3>Semana a semana</h3>
    <div class="scroll"><table id="t-flujo">
      <thead><tr><th>#</th><th>Semana</th><th class="num">Saldo inicial</th>
      <th class="num">Ingresos</th><th class="num">Subcontratos</th><th class="num">Gastos</th>
      <th class="num">Egresos</th><th class="num">Flujo neto</th><th class="num">Saldo final</th>
      <th>Estado</th><th class="num">Monto faltante</th></tr></thead>
      <tbody></tbody></table></div>
  </div>
</div>

<!-- ═══════ 5 · TAREAS ═══════ -->
<div class="vista" id="v-tareas">
  <div class="card">
    <h3>Qué significa cada cosa</h3>
    <div class="nota"><strong>OT (Orden de Trabajo)</strong> — para subcontratistas que emiten factura. Se puede pagar en varias cuotas.
    <strong>OC (Orden de Compra)</strong> — para materiales e insumos. Un único pago por tangibles.
    <strong>OB (Orden de Boleta)</strong> — para prestadores que no emiten factura.</div>
    <div class="nota"><strong>Hito de pago</strong> — el momento del proyecto que habilita el cobro: anticipo, EEPP-1, contra entrega, saldo.
    Es la llave entre lo que se paga al proveedor y lo que se factura al cliente.</div>
    <div class="nota"><strong>Estado del pago</strong> — <code>Pagado</code> y <code>Factorizada</code> son dinero ya movido y no entran al flujo futuro.
    <code>Sin Pago</code> sí entra. <code>Eliminado</code> no debe sumarse.</div>
    <div class="nota"><strong>Margen bruto</strong> — facturado neto del proyecto menos los egresos con proveedores del mismo proyecto.
    <strong>No incluye sueldos, gastos generales ni costos de obra</strong>: es un margen parcial, no la utilidad.</div>
  </div>
  <div class="card">
    <h3>Hallazgos que hay que corregir en el master</h3>
    <div id="tareas-lista"></div>
  </div>
</div>

<script>
const PROV = __PROV__, PROY = __PROY__, SEM = __SEM__, CAB = __CAB__, VENTAS = __VENTAS__, CRUCE = __CRUCE__;
const MON = n => n==null?"—":"$"+Math.round(n).toLocaleString("es-CL");
const esc = s => String(s==null?"":s).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const tagEstado = e => { const l=(e||"").toLowerCase();
  return l==="pagado"?"<span class='tag t-ok'>Pagado</span>"
    : l==="sin pago"?"<span class='tag t-w'>Sin pago</span>"
    : l.startsWith("elimin")?"<span class='tag t-b'>Eliminado</span>"
    : l==="factorizada"?"<span class='tag t-v'>Factorizada</span>"
    : "<span class='tag t-x'>"+esc(e||"—")+"</span>"; };

document.getElementById("pie").textContent =
  PROV.length+" pagos a proveedores · "+PROY.length+" proyectos · "+VENTAS.length+
  " líneas de venta · fuente: MASTER - Venta-Sub-GM_Flujos.xlsx · corte flujo "+CAB.actualizacion;

// ── tabs ──
const TABS = [["proveedores","Proveedores"],["proyectos","Proyectos"],["facturacion","Facturación"],
              ["flujo","Flujo de caja"],["tareas","Tareas"]];
document.getElementById("tabs").innerHTML = TABS.map(([k,t],i)=>
  `<div class="tab ${i===0?'on':''}" data-v="${k}">${t}</div>`).join("");
document.querySelectorAll(".tab").forEach(t=>t.onclick=()=>{
  document.querySelectorAll(".tab").forEach(x=>x.classList.toggle("on",x===t));
  document.querySelectorAll(".vista").forEach(v=>v.classList.toggle("on",v.id==="v-"+t.dataset.v));
});

// ═══ 1 · PROVEEDORES ═══
const pendientes = PROV.filter(p=>p.estado_pago.toLowerCase()==="sin pago");
const kProv = [
  ["Pagos pendientes", pendientes.length, "w"],
  ["Monto pendiente", MON(pendientes.reduce((a,p)=>a+(p.bruto||0),0)), "w"],
  ["Total pagado", MON(PROV.filter(p=>p.estado_pago.toLowerCase()==="pagado").reduce((a,p)=>a+(p.bruto||0),0)), "g"],
  ["Proveedores distintos", new Set(PROV.map(p=>p.rut||p.proveedor)).size, ""],
  ["Sin factura del proveedor", PROV.filter(p=>p.estado_fact==="Pendiente").length, "r"],
];
document.getElementById("k-prov").innerHTML = kProv.map(([k,v,c])=>
  `<div class="kpi ${c}"><b>${v}</b><span>${k}</span></div>`).join("");

function pintarProv(q){
  const s=(q||"").toLowerCase();
  const xs = !s ? PROV.slice().sort((a,b)=>(b.bruto||0)-(a.bruto||0))
    : PROV.filter(p=>(p.proveedor+" "+p.rut+" "+p.proyecto+" "+p.ot+" "+p.hito+" "+p.concepto).toLowerCase().includes(s))
          .sort((a,b)=>(b.bruto||0)-(a.bruto||0));
  document.querySelector("#t-prov tbody").innerHTML = xs.length ? xs.slice(0,400).map(p=>`
    <tr>
      <td class="clip" title="${esc(p.proveedor)}">${esc(p.proveedor)}</td>
      <td class="num">${esc(p.rut||"—")}</td>
      <td class="clip" title="${esc(p.proyecto)}">${esc(p.proyecto)}</td>
      <td class="num" title="${esc(p.ot)}">${esc((p.ot||"—").slice(0,16))}</td>
      <td class="clip" title="${esc(p.hito)}">${esc(p.hito||"—")}</td>
      <td>${tagEstado(p.estado_pago)}</td>
      <td class="num">${MON(p.bruto)}</td>
      <td class="num">${p.estado_pago.toLowerCase()==="pagado"?MON(p.bruto):"—"}</td>
      <td class="num">${p.estado_pago.toLowerCase()==="sin pago"?MON(p.bruto):"—"}</td>
      <td>${esc(p.tipo_doc)}</td>
      <td class="num">${esc(p.factura||"—")}</td>
    </tr>`).join("") : '<tr><td colspan="11" class="empty">Sin resultados</td></tr>';
}
document.getElementById("q-prov").oninput = e=>pintarProv(e.target.value);
pintarProv("");

// ═══ 2 · PROYECTOS ═══
const totEgr = CRUCE.reduce((a,p)=>a+p.egreso_bruto,0);
const totPend = CRUCE.reduce((a,p)=>a+p.egreso_pendiente,0);
const conCruce = CRUCE.filter(p=>p.cruza).length;
document.getElementById("k-proy").innerHTML = [
  ["Proyectos con egresos", CRUCE.filter(p=>p.n_pagos>0).length, ""],
  ["Egresos totales", MON(totEgr), ""],
  ["Pendiente de pago", MON(totPend), "w"],
  ["Proyectos con cruce facturación", conCruce+" / "+CRUCE.filter(p=>p.n_pagos>0).length, "r"],
].map(([k,v,c])=>`<div class="kpi ${c}"><b>${v}</b><span>${k}</span></div>`).join("");

function pintarProy(q){
  const s=(q||"").toLowerCase();
  const xs = CRUCE.filter(p=>p.n_pagos>0 || p.n_facturas>0)
    .filter(p=>!s || (p.proyecto+" "+p.clientes.join(" ")).toLowerCase().includes(s));
  document.querySelector("#t-proy tbody").innerHTML = xs.length ? xs.map(p=>`
    <tr>
      <td class="clip" title="${esc(p.proyecto)}">${esc(p.proyecto)}</td>
      <td class="num">${p.egreso_bruto?MON(p.egreso_bruto):"—"}</td>
      <td class="num">${p.egreso_pagado?MON(p.egreso_pagado):"—"}</td>
      <td class="num ${p.egreso_pendiente>0?'neg':''}">${p.egreso_pendiente?MON(p.egreso_pendiente):"—"}</td>
      <td class="num">${p.facturado_neto?MON(p.facturado_neto):"—"}</td>
      <td class="num ${p.margen_bruto<0?'neg':p.margen_bruto>0?'pos':''}">${p.cruza?MON(p.margen_bruto):"—"}</td>
      <td>${p.cruza?"<span class='tag t-ok'>cruza</span>":
            (p.n_pagos&&!p.n_facturas?"<span class='tag t-b'>sin factura</span>":
            (p.n_facturas&&!p.n_pagos?"<span class='tag t-v'>solo venta</span>":"<span class='tag t-x'>—</span>"))}</td>
      <td class="num">${p.proveedores.length||"—"}</td>
    </tr>`).join("") : '<tr><td colspan="8" class="empty">Sin resultados</td></tr>';
}
document.getElementById("q-proy").oninput = e=>pintarProy(e.target.value);
pintarProy("");

// ═══ 3 · FACTURACIÓN ═══
const vOk = VENTAS.filter(v=>v.tipo_doc.toLowerCase()!=="anulada");
const hitos = {};
vOk.forEach(v=>{ const h=v.hito||"(sin hito)";
  hitos[h]=hitos[h]||{n:0,neto:0,bruto:0};
  hitos[h].n++; hitos[h].neto+=v.neto||0; hitos[h].bruto+=v.bruto||0; });
const sinFact = VENTAS.filter(v=>v.estado_fact==="Sin factura");
document.getElementById("k-fac").innerHTML = [
  ["Documentos emitidos", vOk.length, ""],
  ["Facturado neto", MON(vOk.reduce((a,v)=>a+(v.neto||0),0)), "g"],
  ["Sin facturar (líneas)", sinFact.length, "w"],
  ["Monto sin facturar", MON(sinFact.reduce((a,v)=>a+(v.neto||0),0)), "w"],
  ["Hitos distintos", Object.keys(hitos).length, ""],
].map(([k,v,c])=>`<div class="kpi ${c}"><b>${v}</b><span>${k}</span></div>`).join("");

document.querySelector("#t-hitos tbody").innerHTML = Object.entries(hitos)
  .sort((a,b)=>b[1].neto-a[1].neto).map(([h,d])=>`
    <tr><td>${esc(h)}</td><td class="num">${d.n}</td>
    <td class="num">${MON(d.neto)}</td><td class="num">${MON(d.bruto)}</td>
    <td class="num">${sinFact.filter(v=>(v.hito||"(sin hito)")===h).length||"—"}</td></tr>`).join("");

function pintarFac(q){
  const s=(q||"").toLowerCase();
  const xs = CRUCE.filter(p=>p.n_facturas>0)
    .filter(p=>!s || (p.proyecto+" "+p.clientes.join(" ")).toLowerCase().includes(s));
  document.querySelector("#t-fac tbody").innerHTML = xs.length ? xs.map(p=>`
    <tr>
      <td class="clip" title="${esc(p.proyecto)}">${esc(p.proyecto)}</td>
      <td class="clip" title="${esc(p.clientes.join(', '))}">${esc(p.clientes.join(", ")||"—")}</td>
      <td class="num">${p.n_facturas}</td>
      <td class="num">${MON(p.facturado_neto)}</td>
      <td class="num">${p.egreso_bruto?MON(p.egreso_bruto):"—"}</td>
      <td style="font-size:11.5px">${p.hitos.slice(0,4).map(h=>esc(h.hito)).join(" · ")||"—"}</td>
    </tr>`).join("") : '<tr><td colspan="6" class="empty">Sin resultados</td></tr>';
}
document.getElementById("q-fac").oninput = e=>pintarFac(e.target.value);
pintarFac("");

// ═══ 4 · FLUJO ═══
document.getElementById("k-flujo").innerHTML = [
  ["Saldo bancario inicial", MON(CAB.saldo_inicial), CAB.saldo_inicial<0?"r":"g"],
  ["Ingresos 16 semanas", MON(CAB.ingresos_16), "g"],
  ["Egresos 16 semanas", MON(CAB.egresos_16), "r"],
  ["Flujo neto", MON(CAB.neto_16), CAB.neto_16<0?"r":"g"],
  ["Saldo final proyectado", MON(CAB.saldo_final), "w"],
  ["Mayor necesidad de caja", MON(CAB.necesidad_caja), "r"],
].map(([k,v,c])=>`<div class="kpi ${c}"><b>${v}</b><span>${k}</span></div>`).join("");

const maxAbs = Math.max(...SEM.map(s=>Math.abs(s.saldo_final||0)),1);
document.getElementById("chart").innerHTML = SEM.map(s=>{
  const v = s.saldo_final||0;
  const h = Math.max(3, Math.abs(v)/maxAbs*100);
  const col = v<0? "var(--bad)":"var(--ok)";
  return `<div class="b"><i style="height:${h}%;background:${col}" title="Semana ${s.n}: ${MON(v)}"></i>
    <small>${s.n}</small></div>`;
}).join("");

document.querySelector("#t-flujo tbody").innerHTML = SEM.map(s=>`
  <tr><td class="num">${s.n}</td><td>${esc(s.etiqueta)}</td>
  <td class="num">${MON(s.saldo_inicial)}</td><td class="num pos">${MON(s.ingresos)}</td>
  <td class="num neg">${MON(s.subcontratos)}</td><td class="num neg">${MON(s.gastos)}</td>
  <td class="num">${MON(s.egresos)}</td>
  <td class="num ${(s.flujo_neto||0)<0?'neg':'pos'}">${MON(s.flujo_neto)}</td>
  <td class="num ${(s.saldo_final||0)<0?'neg':''}">${MON(s.saldo_final)}</td>
  <td>${s.estado ? (s.estado.toUpperCase().includes("DEFICIT")||s.estado.toUpperCase().includes("DÉFICIT")
      ? `<span class="tag t-b">déficit</span>` : `<span class="tag t-ok">${esc(s.estado.toLowerCase())}</span>`) : "—"}</td>
  <td class="num ${s.faltante>0?'neg':''}">${s.faltante>0?MON(s.faltante):"—"}</td></tr>`).join("");

// ═══ 5 · TAREAS ═══
const dups = PROV.length; // placeholder de conteo
document.getElementById("tareas-lista").innerHTML = `
  <div class="crit"><strong>1 · Solo ${conCruce} de ${CRUCE.filter(p=>p.n_pagos>0).length} proyectos con egresos cruzan con facturación.</strong>
  El cruce se hace por nombre de proyecto. Los demás están escritos distinto entre la hoja Subcontratos y Ventas,
  así que la pestaña Proyectos no puede calcular su margen. <em>Tarea: unificar el nombre del proyecto en ambas hojas.</em></div>

  <div class="aviso"><strong>2 · El tipo de orden viene escrito de varias formas.</strong>
  <code>OT</code>, <code>OT </code> (con espacio), <code>OB</code>, <code>OC</code>. Un filtro exacto pierde registros.
  <em>Tarea: lista desplegable con valores cerrados.</em></div>

  <div class="aviso"><strong>3 · El estado de pago cambia de género entre hojas.</strong>
  En Ventas es <code>Eliminado</code>; en Subcontratos, <code>Eliminada</code>. Sumar sin normalizar duplica o pierde filas.
  <em>Tarea: un solo vocabulario.</em></div>

  <div class="crit"><strong>4 · Hay 4 pagos repetidos</strong> — mismo proveedor, misma OT, mismo hito y mismo monto.
  Están en las filas que marca el validador. <em>Tarea: verificar si se pagó dos veces o si es un duplicado de carga.</em></div>

  <div class="crit"><strong>5 · Hay 14 filas con un neto de cientos de pesos y un pago en millones.</strong>
  La base de cálculo no puede producir ese pago. <em>Tarea: corregir el TOTAL NETO de esas filas.</em></div>

  <div class="aviso"><strong>6 · 20 documentos "Afecta" tienen Pago Bruto = Pago Neto</strong> (sin IVA),
  y otros 8 tienen ratios atípicos (1,20 / 0,87). <em>Tarea: revisar la columna Pago Bruto.</em></div>

  <div class="nota"><strong>Lo que este tablero NO calcula:</strong> sueldos, gastos generales ni costos de obra.
  El "margen bruto" por proyecto es facturado menos egresos con proveedores: sirve para ver dónde duele,
  no para cerrar el resultado del ejercicio.</div>
`;
</script>
</body>
</html>
"""


def main():
    prov, ventas, sem, cab = leer()
    cruce = cruzar(prov, ventas)

    html = (HTML.replace("__PROV__", json.dumps(prov, ensure_ascii=False))
                .replace("__PROY__", json.dumps(cruce, ensure_ascii=False))
                .replace("__VENTAS__", json.dumps(ventas, ensure_ascii=False))
                .replace("__CRUCE__", json.dumps(cruce, ensure_ascii=False))
                .replace("__SEM__", json.dumps(sem, ensure_ascii=False))
                .replace("__CAB__", json.dumps(cab, ensure_ascii=False)))
    open(OUT, "w", encoding="utf-8").write(html)

    con_pagos = [p for p in cruce if p["n_pagos"] > 0]
    print(f"pagos a proveedores : {len(prov)}")
    print(f"proyectos           : {len(cruce)}")
    print(f"  con egresos       : {len(con_pagos)}")
    print(f"  que cruzan con fac: {sum(1 for p in con_pagos if p['cruza'])}")
    print(f"líneas de venta     : {len(ventas)}")
    print(f"semanas de flujo    : {len(sem)}")
    print(f"pendiente de pago   : ${sum(p['egreso_pendiente'] for p in cruce):,.0f}")
    print(f"escrito: {OUT}")


if __name__ == "__main__":
    main()
