# AGENTS.md

Academia Horizonte: automatizar emisión de certificados por cohorte (hoy manual en Excel/Word) como aplicación web. Explica en español sencillo lo que vayas haciendo.

## Arquitectura (3 piezas)

- Bodega: `Insumos/` (Excel). Solo lectura, NO modificar.
- Cocina (backend): Python + Flask en `app.py`. Toda la lógica va aquí, NUNCA en la interfaz.
- Salón (frontend): HTML servido por Flask. Solo muestra; no calcula ni decide.

## Entradas (verificado del XML)

- `Insumos/Maestro_Estudiantes.xlsx` — hoja `Estudiantes`, `A1:E25` (header + 24 rows).
  `Identificacion | Nombre_Completo | Correo | Programa | Cohorte`. Programas: `Técnico en IA Aplicada` (16) + `Excel Avanzado para Negocios` (8). Cohorte única: `2026-A`. Emails `@ejemplo.cr` (ficticios).
- `Insumos/Registro_Evaluaciones.xlsx` — hoja `Evaluaciones`, `A1:F89` (header + 88 rows).
  `Identificacion | Programa | Modulo | Nota | Asistencia_Pct | Fecha_Cierre`. `Módulo 1`–`Módulo 4` (M4 con solo 16 filas). `Nota`/`Asistencia_Pct` numéricos; `Fecha_Cierre` string fecha ISO. Celdas con `inlineStr`, sin `sharedStrings.xml`.
- Clave de unión: `Identificacion` como string, sin convertir a int (preserva ceros iniciales).

## Reglas de negocio

- `Promedio = suma de Notas / cantidad de módulos cursados`. `Asistencia = promedio de Asistencia_Pct`.
- Agrupar por `Identificacion + Programa`. Un certificado por estudiante y programa.
- Aprobación si `Promedio >= 70` y `Asistencia >= 80`. Participación si `Promedio < 70` y `Asistencia >= 80`. Sin certificado si `Asistencia < 80`. Límites inclusivos.

## Estilo

- Código simple y comentado en español, señalando los ladrillos: variables, tipos, condicionales, bucles y funciones.
- Diseño web en azul marino y dorado.

## Decisiones técnicas (verificado con `pip list`)

- Excel: `openpyxl` (ya instalado 3.1.5). ID siempre como string. `pandas` descartado para la v1 (esconde la lógica); stdlib `zipfile` solo como plan B.
- Web: `Flask` (ya instalado 3.1.3) con `render_template`. PDF futuro con `fpdf2` (ya instalado 2.8.9).
- Estructura: `app.py`, `requirements.txt`, `templates/base.html`, `templates/index.html`, `templates/certificado.html`, `static/style.css`. Lógica solo en `app.py`.
- Plan por etapas en `PLAN.md`. Resultado esperado actual: 19 Aprobación + 3 Participación = 22 certificados.

## Entorno

- Windows + PowerShell. La ruta contiene espacios (`APP Certificados`): siempre entrecomilla rutas.
- `python` es Python 3.12.10 (verificado). `py`, `node`, `npm` no existen — verifica con `Get-Command`.
- Sin dependencias de Excel asumidas: los `.xlsx` son ZIPs legibles con stdlib (`zipfile` + `xl/worksheets/sheet1.xml` + `xl/workbook.xml`).
- `.gitignore` ya excluye `.env`, `.venv/`, `__pycache__/` y salidas `certificados_generados/`, `output/`, `dist/`. Los certificados generados van fuera de `Insumos/` y fuera de git.
