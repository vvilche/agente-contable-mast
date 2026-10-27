# PUNTO DE PARTIDA — Agente Contable MAST (Ejercicio Clases USACH)

Guardado: 2026-10-27 · Estado: **fase de plan, nada construido aún**

## Qué es esto

Ejercicio de la clase de **contabilidad agéntica** (Electivo IV IA Agéntica, FAE USACH).
Automatizar tareas de José, cuya empresa **MAST ARQUITECTURA Y CONSTRUCCIÓN** usa el ERP
**Kame** y maneja su control financiero en un Excel maestro.

Objetivo de arquitectura: construir un agente que respete **los 7 subsistemas + el bucle**
(harness engineering), con seguridad, skills y un **RAG** que garantice que el agente
responda según las normas de MAST y las leyes chilenas vigentes.

## Material de entrada (verificado en disco)

| Archivo | Qué es | Medido |
|---|---|---|
| `~/Documents/EjercicioJose/PROTOCOLO DE GESTIÓN INTERNA.pdf` | BTPR 009-011 v01, fecha 01-08-2026 | 19 págs, 34.795 chars de texto |
| `~/Documents/EjercicioJose/MASTER - Venta-Sub-GM_Flujos.xlsx` | Excel maestro MAST | 7 hojas, 5.600+ fórmulas |

### Hojas del Excel
- `Ventas global` — 331 filas, 20 columnas, 230 fórmulas
- `Subcontratos` — 686 filas, 22 columnas, 1.249 fórmulas
- `Gastos mensuales` — 229 filas, 853 fórmulas
- `Costos proyectados de obras` — 77 filas, 116 fórmulas
- `Flujo de caja` — 22 filas, 327 fórmulas
- `Parámetros` — 480 filas, 1.920 fórmulas (contiene las REGLAS DE NEGOCIO escritas)
- `Detalle diario` — 112 filas, 1.344 fórmulas

### Páginas del protocolo sin texto suficiente (formularios escaneados)
Págs 11-18: anexos 9.1 (entrega de herramientas), 9.2 (OT/OC), 9.3 (orden de boleta),
9.4 (factura de proveedor), 9.5 (formularios de rendición). Requieren OCR si se quieren
en el corpus.

## Fuentes que EXISTEN en disco y NO se usaron todavía

Declarado explícitamente para no repetir el error de construir desde cero:

- `~/orca/workspaces/Contabilidad/Electivo-IA-Conta/datos_estructurados.json` —
  capa determinista con IVA 19%, IDPC, retención honorarios 15,25% (2026), PPM, plazos F29,
  UF/UTM, cada valor con URL de fuente y fecha de vigencia. **Se reutiliza, no se rehace.**
- `~/orca/workspaces/Contabilidad/Electivo-IA-Conta/glosario_siglas.json` — glosario de siglas contables.
- `~/orca/workspaces/Contabilidad/Electivo-IA-Conta/fuentes-contables-chile.json` — mapa de fuentes por dominio.
- `~/orca/workspaces/Contabilidad/Electivo-IA-Conta/usach-cursos/clase-harness-contabilidad.html` —
  material de la Clase 0 (7 subsistemas, bucle, 2 ausencias). **Marco de la arquitectura.**
- `~/.hermes/skills/harness-engineering-agentes/SKILL.md` — 7 subsistemas, 29 patrones, 18 recomendaciones.
- `~/orca/workspaces/InnovaciónAGentica/RagInnovación/build_rag.py` — patrón de construcción de RAG reutilizable.
- `~/orca/workspaces/OdooTransmat/Odoo/defensa-nextai-vs-kame.html` — comparativa previa sobre Kame.

**No hay nada de MAST en disco** (buscado en `~/orca`, `~/Documents`, `~/CONECTA_Proyecto`):
el material de José es nuevo. Sobre Kame solo existe la ficha comercial en caché web
(`~/.hermes/cache/web/www.kame.cl-*.md`) y el protocolo lo menciona en la pág. 13.

## Hallazgos de negocio ya detectados (medidos, no supuestos)

1. **Doble origen de datos.** El protocolo (pág. 13) dice: *"La información de la compañía se
   sustenta en KAME ERP"*. El Excel es una copia paralela del ERP. Si ambos discrepan, ¿cuál gana?
   Es el problema central de diseño, no un detalle.
2. **Anticipo de arranque: protocolo dice 30%**, pero el Excel mezcla hitos de 40%, 30% y 10%.
3. **26 filas "Eliminado"** y **8 "Anulada"** dentro de las 331 de ventas → si el agente suma
   sin filtrar, infla el ingreso.
4. **`Factura Exenta` con ratio bruto/neto 1,19 en 3 de 6 filas** → se aplicó IVA a una exenta.
5. **Ratios imposibles**: 900x y 48.514x entre bruto y neto → celdas corruptas. El agente debe
   rechazarlas y decir qué regla las frenó, no sumarlas.

## Pendiente (no ejecutado)

- Instalar `fastembed` + `faiss` en un venv del proyecto (no están en el venv de Hermes).
- OCR de las páginas 11-18 del protocolo si se quieren en el corpus.
- Definir con José/Víctor qué tarea concreta se automatiza primero (el plan propone el pipeline
  de facturación, pero falta la confirmación).

## Comandos para retomar

```bash
cd ~/orca/workspaces/Clases\ Usach/EjercicioClases
# texto del protocolo ya extraído:
#   ~/.hermes/cache/scratch/jose/protocolo.txt
# dump completo del Excel:
#   ~/.hermes/cache/scratch/jose/xlsx_full.txt
```
