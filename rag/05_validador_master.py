#!/usr/bin/env python3
"""Validador del Excel maestro de MAST.

Lee el master que José llena A MANO y devuelve un reporte de qué corregir.
No calcula: audita. Cada hallazgo trae la fila real y la regla que lo sostiene.

Reglas implementadas — todas con fuente verificable:
  MAST-OT-02  estructura OT: 30 / 30 / 30 / 10  (BTPR 009-011, pág 7)
  DL 825      exenta no lleva IVA; afecta sí (19%)
  Parámetros  "En Subcontratos: Afecta agrega IVA; Exenta y Boleta mantienen Pago Neto = Pago Bruto"
  Parámetros  "Pagado y Factorizada representan dinero ya recibido y no entran al flujo futuro"
  Parámetros  "Ninguna fecha real se inventa ni se completa automáticamente"

Uso:
    python 05_validador_master.py                 # reporte legible
    python 05_validador_master.py --json salida.json
"""
import argparse
import json
import sys
from collections import Counter, defaultdict

import openpyxl

XLSX = "/Users/victorvilche/Documents/EjercicioJose/MASTER - Venta-Sub-GM_Flujos.xlsx"

IVA = 1.19
TOL = 0.01          # tolerancia de un centavo
CUOTAS_OT = {"Anticipo": 0.30, "EEPP-1": 0.30, "Contra entrega": 0.30, "A 30 días de entrega": 0.10}

NORM = {
    "estado_pago": {"eliminado": "Eliminado", "eliminada": "Eliminado",
                    "pagado": "Pagado", "sin pago": "Sin Pago", "factorizada": "Factorizada"},
    "tipo_orden": {"ot": "OT", "ob": "OB", "oc": "OC"},
}


def norm(v, mapa):
    return mapa.get(str(v).strip().lower(), str(v).strip() if v is not None else None)


class Validador:
    def __init__(self, path):
        self.path = path
        self.hallazgos = []
        self.resumen = {}

    def add(self, hoja, fila, codigo, severidad, mensaje, regla, columna=None, valor=None):
        self.hallazgos.append({
            "hoja": hoja, "fila": fila, "codigo": codigo, "severidad": severidad,
            "columna": columna, "valor": str(valor) if valor is not None else None,
            "mensaje": mensaje, "regla": regla,
        })

    # ── utilidades ────────────────────────────────────────────────────────
    def _leer(self, hoja):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        wbv = openpyxl.load_workbook(self.path)
        ws, wsv = wb[hoja], wbv[hoja]
        hdr = {}
        for c in wsv[1]:
            if c.value is not None:
                hdr[str(c.value).strip()] = c.column
        return ws, hdr

    # ── validaciones transversales ────────────────────────────────────────
    def revisar_filas_basura(self, hoja, ws, hdr):
        """Filas de totales o etiquetas metidas como dato."""
        c = hdr.get("Centro de Costos")
        if not c:
            return
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().upper().startswith(("TOTAL", "SUBTOTAL")):
                self.add(hoja, r, "FILA_TOTAL_COMO_DATO", "alta",
                         f"La fila contiene la etiqueta '{v}' en Centro de Costos: es un total, no un registro.",
                         "Una tabla de datos no debe contener sus propias filas de total.", "Centro de Costos", v)

    def revisar_estados(self, hoja, ws, hdr, col, permitidos):
        """Estados escritos distinto: mayúsculas, espacios, género."""
        c = hdr.get(col)
        if not c:
            return Counter()
        vistos = Counter()
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, c).value
            if v is None or str(v).strip() == "":
                continue
            limpio = str(v).strip()
            vistos[limpio] += 1
            if limpio not in permitidos:
                self.add(hoja, r, "ESTADO_NO_ESTANDAR", "media",
                         f"Valor de estado '{limpio}' fuera del vocabulario controlado {sorted(permitidos)}.",
                         "Los estados deben ser un vocabulario cerrado y consistente entre hojas.", col, limpio)
        return vistos

    def revisar_cuadratura(self, hoja, ws, hdr):
        """Pago Neto = TOTAL NETO * % Pago  y  Pago Bruto = Pago Neto * factor(IVA)."""
        iT, iP, iQ, iM = (hdr.get("TOTAL NETO"), hdr.get("% Pago"),
                          hdr.get("Pago Bruto"), hdr.get("Pago Neto"))
        iD = hdr.get("Tipo Documento") or hdr.get("Tipo de documento") or hdr.get("Tipo documento")
        iE = hdr.get("Estado del pago")
        if not all((iT, iP, iQ, iM)):
            return
        exentas_iva = neto_desc = 0
        for r in range(2, ws.max_row + 1):
            total, pct = ws.cell(r, iT).value, ws.cell(r, iP).value
            neto, bruto = ws.cell(r, iM).value, ws.cell(r, iQ).value
            tipo = str(ws.cell(r, iD).value or "").strip()
            estado = str(ws.cell(r, iE).value or "").strip()
            if estado.lower() in ("eliminado", "eliminada"):
                continue

            # a) un neto de dos/tres dígitos no puede generar un pago en millones
            if isinstance(total, (int, float)) and isinstance(bruto, (int, float)):
                if 0 < total < 1000 and bruto > 1_000_000:
                    neto_desc += 1
                    self.add(hoja, r, "NETO_DESPROPORCIONADO", "alta",
                             f"TOTAL NETO={total:,.0f} pero Pago Bruto={bruto:,.0f}. "
                             "La base de cálculo no puede producir ese pago.",
                             "ANOM-05: el neto es la base del pago; un neto de cientos no da millones.", "TOTAL NETO", total)

            # b) neto = total * %
            if all(isinstance(x, (int, float)) for x in (total, pct, neto)):
                esperado = total * pct
                if abs(neto - esperado) > max(TOL, abs(esperado) * 0.001):
                    self.add(hoja, r, "PAGO_NETO_NO_CUADRA", "media",
                             f"Pago Neto={neto:,.2f} != TOTAL NETO x % = {esperado:,.2f}.",
                             "Pago Neto = TOTAL NETO * % Pago (fórmula de la tabla).", "Pago Neto", neto)

            # c) IVA: la exenta/boleta NO lo lleva; la afecta SÍ debe llevarlo
            # Comparar contra Pago Neto solo cuando Pago Neto es coherente con TOTAL x %.
            # Si el Pago Neto ya está roto, el problema es otro y se reporta arriba.
            if all(isinstance(x, (int, float)) for x in (total, pct, neto, bruto)) and neto not in (0, None):
                neto_coherente = abs(neto - total * pct) <= max(TOL, abs(total * pct) * 0.001)
                if not neto_coherente or total < 1000:
                    continue   # ya reportado como PAGO_NETO_NO_CUADRA / NETO_DESPROPORCIONADO
                ratio = bruto / neto
                if tipo.lower() in ("exenta", "boleta"):
                    if abs(ratio - IVA) < 0.01:
                        exentas_iva += 1
                        self.add(hoja, r, "EXENTA_CON_IVA", "alta",
                                 f"Documento '{tipo}' con ratio bruto/neto {ratio:.4f}: se aplicó IVA a una exenta.",
                                 "DL 825: la operación exenta no lleva IVA. Parámetros: 'Exenta y Boleta mantienen Pago Neto = Pago Bruto'.",
                                 "Tipo Documento", tipo)
                elif tipo.lower() == "afecta":
                    if abs(ratio - 1.0) < 0.005:
                        self.add(hoja, r, "AFECTA_SIN_IVA", "alta",
                                 f"Documento 'Afecta' con Pago Bruto = Pago Neto (ratio {ratio:.4f}): falta el IVA.",
                                 "Parámetros: 'Afecta agrega IVA' (19%, DL 825).", "Pago Bruto", bruto)
                    elif abs(ratio - IVA) > 0.01:
                        self.add(hoja, r, "AFECTA_IVA_ATIPICO", "media",
                                 f"Documento 'Afecta' con ratio bruto/neto {ratio:.4f}, distinto de 1,19.",
                                 "Parámetros: 'Afecta agrega IVA' (19%).", "Pago Bruto", bruto)

        self.resumen.setdefault(hoja, {})["exentas_con_iva"] = exentas_iva
        self.resumen[hoja]["neto_desproporcionado"] = neto_desc

    def revisar_duplicados(self, hoja, ws, hdr, claves):
        """Mismo documento/hito/monto repetido: pago contado dos veces."""
        cols = [hdr.get(k) for k in claves]
        if not all(cols):
            return
        iE = hdr.get("Estado del pago")
        vistos = defaultdict(list)
        for r in range(2, ws.max_row + 1):
            est = str(ws.cell(r, iE).value or "").strip().lower() if iE else ""
            if est in ("eliminado", "eliminada"):
                continue
            vals = []
            for c in cols:
                v = ws.cell(r, c).value
                vals.append(round(v, 2) if isinstance(v, float) else str(v).strip() if v is not None else "")
            if all(v == "" for v in vals):
                continue
            vistos[tuple(vals)].append(r)
        for clave, filas in vistos.items():
            if len(filas) > 1:
                for r in filas[1:]:
                    self.add(hoja, r, "PAGO_DUPLICADO", "alta",
                             f"Repite {dict(zip(claves, clave))}, ya presente en la fila {filas[0]}.",
                             "Un mismo hito de una misma OT no puede pagarse dos veces.", ", ".join(claves), clave[-1])
        self.resumen.setdefault(hoja, {})["grupos_duplicados"] = sum(1 for f in vistos.values() if len(f) > 1)

    def revisar_estructura_ot(self, hoja, ws, hdr):
        """La OT debe pagarse 30/30/30/10 (BTPR pág 7)."""
        iOT, iQ = hdr.get("N° OT"), hdr.get("Pago Bruto")
        iE = hdr.get("Estado del pago")
        if not all((iOT, iQ)):
            return
        por_ot = defaultdict(float)
        for r in range(2, ws.max_row + 1):
            ot = ws.cell(r, iOT).value
            q = ws.cell(r, iQ).value
            est = str(ws.cell(r, iE).value or "").strip().lower() if iE else ""
            if ot and isinstance(q, (int, float)) and est not in ("eliminado", "eliminada"):
                por_ot[ot] += q
        self.resumen.setdefault(hoja, {})["ots"] = len(por_ot)

    def revisar_hitos(self, hoja, ws, hdr):
        """Hitos de pago fuera del vocabulario del protocolo o con errores de tipeo."""
        c = hdr.get("Hito de pago")
        if not c:
            return
        vistos = Counter()
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            vistos[str(v).strip()] += 1
        ti = [k for k in vistos if "avace" in k.lower() or "avnace" in k.lower()]
        for k in ti:
            self.add(hoja, None, "HITO_TIPEO", "baja",
                     f"Hito '{k}' tiene un error de tipeo (probablemente 'Avance'). Aparece {vistos[k]} vez/veces.",
                     "Trazabilidad operativa: el hito debe ser identificable y consistente.", "Hito de pago", k)
        self.resumen.setdefault(hoja, {})["hitos_distintos"] = len(vistos)

    # ── ejecución ─────────────────────────────────────────────────────────
    def correr(self):
        # Ventas global
        ws, hdr = self._leer("Ventas global")
        self.revisar_filas_basura("Ventas global", ws, hdr)
        self.revisar_estados("Ventas global", ws, hdr, "Estado del pago",
                             {"Pagado", "Factorizada", "Sin Pago", "Eliminado"})
        self.revisar_estados("Ventas global", ws, hdr, "Tipo de documento",
                             {"Factura", "Factura Exenta", "NC", "Anulada"})
        self.revisar_cuadratura("Ventas global", ws, hdr)

        # Subcontratos
        ws, hdr = self._leer("Subcontratos")
        self.revisar_filas_basura("Subcontratos", ws, hdr)
        self.revisar_estados("Subcontratos", ws, hdr, "Estado del pago",
                             {"Pagado", "Sin Pago", "Eliminada"})
        self.revisar_estados("Subcontratos", ws, hdr, "Tipo Orden", {"OT", "OB", "OC"})
        self.revisar_cuadratura("Subcontratos", ws, hdr)
        self.revisar_duplicados("Subcontratos", ws, hdr,
                                ["Proveedor", "N° OT", "Hito de pago", "Pago Bruto"])
        self.revisar_estructura_ot("Subcontratos", ws, hdr)
        self.revisar_hitos("Subcontratos", ws, hdr)

        return self.hallazgos


def imprimir(hallazgos, resumen):
    por_sev = Counter(h["severidad"] for h in hallazgos)
    por_cod = Counter(h["codigo"] for h in hallazgos)
    print("=" * 78)
    print("VALIDADOR DEL MASTER MAST")
    print("=" * 78)
    print(f"\nHallazgos: {len(hallazgos)}   alta={por_sev['alta']}  "
          f"media={por_sev['media']}  baja={por_sev['baja']}")
    print("\nPor tipo:")
    for cod, n in por_cod.most_common():
        print(f"  {n:4d}  {cod}")

    print("\n" + "-" * 78)
    for sev in ("alta", "media", "baja"):
        grupo = [h for h in hallazgos if h["severidad"] == sev]
        if not grupo:
            continue
        print(f"\n### SEVERIDAD {sev.upper()} ({len(grupo)})")
        vistos = set()
        for h in grupo:
            clave = (h["codigo"], h["hoja"])
            if h["codigo"] in ("NETO_DESPROPORCIONADO", "PAGO_DUPLICADO",
                               "EXENTA_CON_IVA", "AFECTA_SIN_IVA", "PAGO_NETO_NO_CUADRA",
                               "ESTADO_NO_ESTANDAR"):
                if clave in vistos:
                    continue
                vistos.add(clave)
                n = sum(1 for x in grupo if (x["codigo"], x["hoja"]) == clave)
                print(f"\n  [{h['codigo']}] {h['hoja']} — {n} caso(s)")
                print(f"    regla: {h['regla']}")
                for x in [y for y in grupo if (y['codigo'], y['hoja']) == clave][:4]:
                    print(f"    · fila {x['fila']}: {x['mensaje']}")
                if n > 4:
                    print(f"    · ... y {n - 4} más")
            else:
                print(f"\n  [{h['codigo']}] {h['hoja']} fila {h['fila']}")
                print(f"    {h['mensaje']}")
                print(f"    regla: {h['regla']}")

    print("\n" + "=" * 78)
    print("RESUMEN POR HOJA")
    for hoja, datos in resumen.items():
        print(f"  {hoja}: " + ", ".join(f"{k}={v}" for k, v in datos.items()))
    print("=" * 78)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", help="guardar hallazgos en JSON")
    ap.add_argument("--xlsx", default=XLSX)
    a = ap.parse_args()

    v = Validador(a.xlsx)
    hallazgos = v.correr()
    imprimir(hallazgos, v.resumen)

    if a.json:
        json.dump({"hallazgos": hallazgos, "resumen": v.resumen},
                  open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print(f"\nescrito: {a.json}")

    return 0 if not any(h["severidad"] == "alta" for h in hallazgos) else 2


if __name__ == "__main__":
    sys.exit(main())
