# PLAN.md

Etapas: 1) validar cruce de datos → 2) backend Flask → 3) interfaz. No avanzar sin validar la etapa anterior.

## Etapa 1 — Validación del cruce (sin servidor)

Qué se hace:
- Leer ambos Excel con `openpyxl`, ID siempre como string.
- Agrupar evaluaciones por `Identificacion + Programa`.
- Calcular `Promedio` y `Asistencia` por grupo y aplicar veredicto (Aprobación / Participación / Sin certificado).
- Detectar inconsistencias: maestro sin notas (`304560321`), huérfano sin maestro (`999880777`), abandono (`111230984` con ceros), asimetría IA (4 módulos) vs Excel (3).

Validación (antes de seguir):
- Conteos esperados con datos actuales: 24 grupos en evaluaciones, 23 evaluables, **19 Aprobación + 3 Participación = 22 certificados**, 1 sin certificado, 1 sin datos, 1 huérfano reportado.
- Confirmar con el usuario: huérfano se reporta y no se certifica; el `0` cuenta como módulo cursado.

## Etapa 2 — Backend Flask (`app.py`)

Qué se hace:
- Funciones puras en `app.py`: `cargar_maestro`, `cargar_evaluaciones`, `agrupar`, `calcular_promedio`, `calcular_asistencia`, `decidir_tipo`.
- Rutas: `GET /` (resumen), `GET /estudiante/<id>` (ficha), `GET /certificado/<id>/<programa>` (solo si aplica).
- `requirements.txt` con Flask, openpyxl (fpdf2 queda para después del MVP).

Validación:
- Levantar con `python app.py` y pedir: resumen 19/3/1, ficha de un caso de cada tipo (ej. `101230456` Aprobación, `116890234` Participación, `111230984` Sin certificado, `304560321` sin datos).
- Toda decisión viene del backend; el HTML no calcula.

## Etapa 3 — Interfaz (azul marino y dorado)

Qué se hace:
- `templates/base.html`, `index.html` (buscador + tabla), `certificado.html` (vista previa), `static/style.css`.
- `<meta charset="utf-8">` para tildes (`Técnico`, `Módulo`).

Validación:
- Revisión visual: buscador funciona, tildes correctas, veredicto coincide con Etapa 1, sin certificado muestra motivo en vez de diploma.
- Cierre: `Insumos/` intacto, nada generado dentro de git (`certificados_generados/`, `output/`, `dist/` ignorados).
