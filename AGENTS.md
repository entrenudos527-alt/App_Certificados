# AGENTS.md

Greenfield repo — no app code, build, test, or lint config yet. Only data inputs exist.

## Sources of truth (read-only)

- `Insumos/Maestro_Estudiantes.xlsx` — sheet `Estudiantes`, range `A1:E25` (header + 24 rows).
  Columns: `Identificacion | Nombre_Completo | Correo | Programa | Cohorte`.
  Single program/cohort observed: `Técnico en IA Aplicada`, `2026-A`. Sample emails `@ejemplo.cr` (fictional).
- `Insumos/Registro_Evaluaciones.xlsx` — sheet `Evaluaciones`, range `A1:F89` (header + 88 rows).
  Columns: `Identificacion | Programa | Modulo | Nota | Asistencia_Pct | Fecha_Cierre`.
  Modules observed: `Módulo 1`–`Módulo 4`. `Nota`/`Asistencia_Pct` numeric; `Fecha_Cierre` ISO date string. Cells use `inlineStr`, no `sharedStrings.xml`.
- Join key across both files is `Identificacion` (string; preserve leading zeros, do not cast to int).
- No eligibility / passing-grade / attendance rules are defined in the repo — do not invent them. Ask or leave undecided.

Do not edit files under `Insumos/`; treat as immutable inputs.

## Environment quirks

- OS is Windows, default shell is PowerShell. The workspace path contains spaces (`APP Certificados`): always quote paths.
- No Python, Node, or package manager was available at last check (`python`, `py`, `node`, `npm` all missing). Verify toolchain with `Get-Command` before assuming one; do not add speculative run/test commands here.
- `.xlsx` parsing without deps: files are ZIPs readable via `System.IO.Compression.ZipFile` + `xl/worksheets/sheet*.xml` + `xl/workbook.xml`.

## When scaffolding the app

- Keep this file updated: record the chosen stack, exact dev/test/lint commands, and entrypoints once they exist.
- Keep generated certificates / outputs out of `Insumos/` and out of git.
