#!/usr/bin/env python3
"""Verifica que cada cita_literal de 02-reglas-mast.json exista REALMENTE
en la página que el JSON declara. Si una cita no aparece, la regla es humo."""
import json, re, sys
import pymupdf

PDF = "/Users/victorvilche/Documents/EjercicioJose/PROTOCOLO DE GESTIÓN INTERNA.pdf"
JSON = "/Users/victorvilche/orca/workspaces/Clases Usach/EjercicioClases/rag/02-reglas-mast.json"


def norm(s):
    s = s.lower()
    s = re.sub(r'[“”"\'’]', '', s)
    s = re.sub(r'[^a-z0-9%áéíóúñü ]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def main():
    doc = pymupdf.open(PDF)
    reglas = json.load(open(JSON, encoding="utf-8"))["reglas"]
    fallos = []
    for r in reglas:
        pg = r["pagina"]
        if not (1 <= pg <= doc.page_count):
            fallos.append((r["id"], f"página {pg} fuera de rango"))
            continue
        texto_pg = norm(doc[pg - 1].get_text())
        trozos = [t for t in r["cita_literal"].split(" / ") if len(t.strip()) > 10]
        faltan = []
        for t in trozos:
            ancla = " ".join(norm(t).split()[:6])
            if ancla not in texto_pg:
                faltan.append(ancla[:70])
        if faltan:
            fallos.append((r["id"], f"p{pg}: no encontrado -> {faltan}"))
        else:
            print(f"OK     {r['id']:14s} p{pg:02d}  ({len(trozos)} trozo/s)")

    print(f"\n{len(reglas) - len(fallos)}/{len(reglas)} reglas con cita verificada")
    if fallos:
        print("\nFALLAN:")
        for rid, msg in fallos:
            print(f"  {rid}: {msg}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
