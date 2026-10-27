# Agente Contable MAST

Arquitectura y plan de un agente de IA para automatizar tareas contables y de control de obra
en **MAST Arquitectura y Construcción** (ERP Kame + Excel maestro), construido respetando el
marco de *harness engineering*: los 7 subsistemas y el bucle.

**Página publicada:** https://vvilche.github.io/agente-contable-mast/

## Contexto

Ejercicio de la clase de **contabilidad agéntica** (Electivo IV IA Agéntica, FAE USACH).

La empresa controla ventas, subcontratos, gastos y flujo de caja en un Excel maestro que
**se llena a mano**, mientras su protocolo interno (BTPR 009-011) declara que *"la información
de la compañía se sustenta en KAME ERP"*.

Hay dos orígenes para el mismo número. El agente no viene a calcular mejor que el Excel:
viene a decidir con qué fuente responde y a **rechazar lo que no puede ser cierto**, devolviendo
la fila exacta para que el humano la corrija.

## Qué hay acá

| Archivo | Contenido |
|---|---|
| `site/index.html` | La arquitectura completa: los 7 subsistemas con posición concreta, el bucle, el diseño del RAG de 4 dominios, el plan en 5 fases y las preguntas abiertas |
| `PUNTO_DE_PARTIDA.md` | Estado del trabajo, fuentes usadas, qué existe en disco y qué no |

Los JSON de reglas y la spec del agente **no están en este repo público**: contienen montos
y detalles internos de MAST. Viven en el entorno de trabajo local.

## Los 7 subsistemas, en una línea cada uno

| # | Subsistema | Posición para MAST |
|---|---|---|
| D1 | Agent loop | Un `while` lineal. Sin framework. Parada dura si un control de cuadratura no cuadra |
| D2 | Integración LLM | El LLM nunca produce un número: solo redacta y elige herramienta |
| D3 | Herramientas | 6, todas deterministas. Sin herramienta de escritura al ERP |
| D4 | Memoria | Contexto del protocolo + historial auditable con fila y artículo de ley |
| D5 | Seguridad | Modo PLAN por defecto. Prohibido emitir, firmar, pagar, declarar |
| D6 | Orquestación | Single-agent por diseño (~15× menos tokens que un swarm) |
| D7 | Extensibilidad | Capacidad nueva = skill, no código nuevo en el loop |

## RAG: 4 dominios, ningún índice colapsado

- **A** Norma MAST — protocolo BTPR 009-011 y reglas internas
- **B** Kame / ERP — módulos, campos, semántica OT/OC/OB
- **C** Ley chilena — Código del Trabajo, Ley 20.123, DS 44, Ley 21.595, DL 825, DL 824
- **D** Capa determinista — IVA, retenciones, PPM, plazos, UF/UTM: **lookup por clave, nunca embedding**

Una tasa recuperada por similitud semántica puede traer el valor de otro año. El IVA no se
pregunta al índice vectorial: se consulta por clave, con vigencia y fuente.

## Publicar localmente

```bash
cd site && python3 -m http.server 8000
# luego abrir http://localhost:8000
```

## Estado

Fase de plan y diseño. Nada de código de agente construido todavía.
Pendiente de decisión: qué tarea se automatiza primero, si hay acceso a Kame por API,
y si el anticipo de arranque es 30% (protocolo) o el 40% que aparece en el Excel.
