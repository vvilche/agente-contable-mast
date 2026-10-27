# RAG del Agente Contable MAST — documentos y estado

Corpus para que el agente de José responda **según las normas de MAST y las leyes chilenas vigentes**,
y no según lo que el modelo crea recordar.

## Los archivos

| Archivo | Qué es | Estado |
|---|---|---|
| `01-manifiesto-corpus.json` | Inventario de los 4 dominios: qué hay, qué falta, cómo se chunkea cada fuente | Completo |
| `02-reglas-mast.json` | 14 reglas del protocolo + 10 criterios del Excel, cada uno con cita literal y página | **14/14 citas verificadas** |
| `03-spec-agente.md` | Especificación previa: qué hace, reglas duras, el bucle, herramientas, formato de respuesta y 12 preguntas de evaluación | Completo |
| `verificar_citas.py` | Comprueba que cada cita del JSON exista en la página que declara | Corre, 14/14, exit 0 |
| `verificar_paquete.py` | Verifica el paquete entero: JSON, HTML y que los conteos declarados cuadren con el Excel real | Corre, 34/34, exit 0 |

## Por qué 4 dominios y no uno

Si se mete todo en un solo índice vectorial, la recuperación se sesga por tipo de documento.
Medido antes en otro proyecto: un catálogo de 580 documentos se perdió dentro de un dominio de
63.000 fragmentos y el agente respondió con el competidor. Acá se separa desde el principio:

| Dominio | Contenido | Índice |
|---|---|---|
| **A** Norma MAST | BTPR 009-011 + reglas de `Parámetros` | Vectorial |
| **B** Kame / ERP | Módulos, campos, semántica OT/OC/OB | Vectorial — **incompleto** |
| **C** Ley chilena | Las 4 leyes que MAST declara + 21.595, DL 825, DL 824 | Vectorial con cita |
| **D** Capa determinista | IVA, retenciones, PPM, plazos, UF/UTM | **Lookup, nunca embedding** |

### El dominio D va aparte por una razón de plata

Una tasa recuperada por similitud semántica puede traer el valor equivocado de otro año.
El IVA no se pregunta al índice: se consulta por clave, y el JSON trae `vigencia_desde`,
`vigencia_hasta`, `fuente_url` y `fecha_consulta` en cada valor.

Esa base **ya existe** y se reutiliza: `~/orca/workspaces/Contabilidad/Electivo-IA-Conta/datos_estructurados.json`
(IVA 19%, IDPC general 27%, Pro Pyme 12,5% hasta 2027 y 15% desde 2028, retención honorarios
15,25% en 2026, PPM, plazos F29, UF y UTM). No se rehízo.

## El hueco principal: el dominio B

El protocolo dice en su página 13 que *"la información de la compañía se sustenta en KAME ERP"*.
Pero en el Mac **no hay documentación técnica de Kame**: solo la ficha comercial en caché web.
Sin acceso a Kame (API o exportación), el agente no puede consultar el ERP y queda dependiente
del Excel — que es justamente la copia que se quiere dejar de usar como fuente.

**Acción pendiente con José: pedir acceso a Kame.**

## Los formularios del protocolo no son normativa

Las páginas 11 a 18 del PDF son imágenes escaneadas: los anexos 9.1 a 9.5 (formularios de
entrega de herramientas, OT/OC, orden de boleta, factura tipo, rendiciones). Son **plantillas**,
no reglas. Se pueden OCR-ear si se quiere que el agente sepa qué campos llenar, pero no son
críticas para responder sobre el proceso.

## Cómo se verifica que las citas son reales

```bash
cd "~/orca/workspaces/Clases Usach/EjercicioClases/rag"
~/.hermes/hermes-agent/venv/bin/python verificar_citas.py
```

Recorre cada regla de `02-reglas-mast.json`, busca las primeras 6 palabras de su `cita_literal`
en la página que declara, y falla si no aparece. Correrlo después de cualquier edición del JSON.

Resultado actual: **14/14**, exit 0.

Una cita falló en la primera corrida y por eso el script existe: `MAST-DEPTO-01` decía
"DEPARTAMENTO DE ADMINISTRACIÓN Y FINANZAS" cuando el PDF escribe "DEPARTAMNENTO" con una N
de más y corta el encabezado en tres líneas. Se corrigió la cita al texto real del PDF.

## Qué falta para tener el RAG funcionando

1. **Instalar** `fastembed` + `faiss` en un venv del proyecto (no están en el venv de Hermes).
2. **Descargar** las leyes del dominio C (`bcn.cl/leychile`) — determinista, sin LLM.
3. **Conseguir** documentación de Kame del dominio B.
4. **Extraer** las reglas de la hoja `Parámetros` a JSON (el bloque de criterios ya está
   transcrito en `02-reglas-mast.json`; falta el resto de la hoja).
5. **Correr** las 12 preguntas de evaluación de `03-spec-agente.md` y registrar recall y faithfulness.
