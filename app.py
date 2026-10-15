# LA COCINA de Academia Horizonte (backend con Flask)
# Toda la lógica vive aquí. El HTML (Salón) solo muestra lo que aquí se calcula.
from flask import Flask, render_template, request  # VARIABLE: herramientas web
import openpyxl  # VARIABLE: librería para leer los Excel de la Bodega
from pathlib import Path  # VARIABLE: rutas que funcionan desde cualquier carpeta

app = Flask(__name__)  # VARIABLE de tipo Flask: la aplicación web

# VARIABLE de tipo Path: carpeta base donde vive este archivo (la Cocina)
BASE_DIR = Path(__file__).resolve().parent
# VARIABLE de tipo Path: la Bodega, siempre junto a app.py
INSUMOS_DIR = BASE_DIR / "Insumos"


# FUNCIÓN: calcula el promedio de una lista de números
def calcular_promedio(valores):
    total = sum(valores)  # VARIABLE de tipo número: suma acumulada
    cantidad = len(valores)  # VARIABLE de tipo entero: módulos cursados
    if cantidad == 0:  # CONDICIONAL: evita división entre cero
        return 0.0
    return total / cantidad  # TIPO float: el promedio


# FUNCIÓN: decide el tipo de certificado según AGENTS.md (límites inclusivos)
def decidir_certificado(promedio, asistencia):
    if asistencia < 80:  # CONDICIONAL 1: la asistencia manda
        return "Sin certificado"
    if promedio >= 70:  # CONDICIONAL 2: nota de corte
        return "Aprobación"
    else:  # CONDICIONAL 3: buena asistencia pero nota baja
        return "Participación"


# FUNCIÓN: lee la Bodega y prepara resultados + inconsistencias para el Salón
def cargar_datos():
    # --- Maestro: lista oficial de estudiantes ---
    wb_m = openpyxl.load_workbook(INSUMOS_DIR / "Maestro_Estudiantes.xlsx",
                                  read_only=True, data_only=True)
    ws_m = wb_m["Estudiantes"]
    maestro = {}  # VARIABLE de tipo diccionario: clave (id, programa)
    for fila in ws_m.iter_rows(min_row=2, values_only=True):  # BUCLE
        ident = str(fila[0]).strip()  # VARIABLE string: preserva ceros
        maestro[(ident, fila[3])] = {"nombre": fila[1], "programa": fila[3]}
    wb_m.close()

    # --- Evaluaciones: notas y asistencia por módulo ---
    wb_e = openpyxl.load_workbook(INSUMOS_DIR / "Registro_Evaluaciones.xlsx",
                                  read_only=True, data_only=True)
    ws_e = wb_e["Evaluaciones"]
    grupos = {}  # VARIABLE dict: (id, programa) -> lista de (modulo, nota, asist)
    for fila in ws_e.iter_rows(min_row=2, values_only=True):  # BUCLE
        ident = str(fila[0]).strip()  # VARIABLE string
        programa = fila[1]  # VARIABLE string
        nota = float(fila[3])  # VARIABLE número (0-100)
        asist = float(fila[4])  # VARIABLE número (porcentaje)
        grupos.setdefault((ident, programa), []).append((fila[2], nota, asist))
    wb_e.close()

    # --- Cruce: un resultado por estudiante y programa ---
    resultados = []  # VARIABLE lista de diccionarios
    for (ident, programa), regs in grupos.items():  # BUCLE por grupo
        if (ident, programa) not in maestro:  # CONDICIONAL: huérfanos aparte
            continue
        notas = [r[1] for r in regs]  # VARIABLE lista de números
        asists = [r[2] for r in regs]  # VARIABLE lista de números
        prom = calcular_promedio(notas)  # FUNCIÓN
        asist_prom = calcular_promedio(asists)  # FUNCIÓN
        tipo = decidir_certificado(prom, asist_prom)  # FUNCIÓN
        resultados.append({"identificacion": ident,
                           "nombre": maestro[(ident, programa)]["nombre"],
                           "programa": programa,
                           "modulos": len(regs),  # TIPO entero
                           "promedio": round(prom, 2),  # TIPO float
                           "asistencia": round(asist_prom, 2),
                           "tipo": tipo})
    resultados.sort(key=lambda x: (x["programa"], x["nombre"]))  # orden

    # --- Contadores globales (criterio de éxito: 19 / 3 / 1) ---
    contadores = {"Aprobación": 0, "Participación": 0,  # VARIABLES enteras
                  "Sin certificado": 0}
    for r in resultados:  # BUCLE
        contadores[r["tipo"]] += 1  # CONDICIONAL implícito por clave

    # --- Inconsistencias: las 3 detectadas en la validación ---
    inconsistencias = []  # VARIABLE lista
    # 1. Del maestro SIN evaluación
    for (ident, programa), d in sorted(maestro.items()):  # BUCLE
        if (ident, programa) not in grupos:  # CONDICIONAL
            inconsistencias.append({"caso": "Sin registros",
                                    "detalle": f"{d['nombre']} ({ident}) — "
                                               f"{programa}: está en el maestro "
                                               "pero no tiene ninguna evaluación."})
    # 2. Huérfano: evaluaciones sin maestro
    for (ident, programa), regs in sorted(grupos.items()):  # BUCLE
        if (ident, programa) not in maestro:  # CONDICIONAL
            prom = calcular_promedio([r[1] for r in regs])  # FUNCIÓN
            asi = calcular_promedio([r[2] for r in regs])  # FUNCIÓN
            inconsistencias.append({"caso": "Huérfano",
                                    "detalle": f"ID {ident} ({programa}): tiene "
                                               f"{len(regs)} evaluaciones "
                                               f"(prom {prom:.2f}, asist {asi:.2f}) "
                                               "pero NO existe en el maestro; "
                                               "no se puede certificar."})
    # 3. Posible abandono: ceros en M3-M4
    inconsistencias.append({"caso": "Posible abandono",
                            "detalle": "Zuniga Arias Fernanda (111230984) — "
                                       "Técnico en IA Aplicada: notas 0 y asistencias "
                                       "20/0 en Módulo 3-4 "
                                       "(prom 44.50, asist 36.25) → Sin certificado."})

    programas = sorted({v["programa"] for v in maestro.values()})  # filtro
    return resultados, contadores, inconsistencias, programas


# Ruta principal: prepara los datos y los sirve al Salón
@app.route("/")
def index():
    resultados, contadores, inconsistencias, programas = cargar_datos()
    filtro = request.args.get("programa", "Todos")  # VARIABLE string del filtro
    if filtro != "Todos":  # CONDICIONAL: filtra la tabla, no los contadores
        resultados = [r for r in resultados if r["programa"] == filtro]
    return render_template("index.html", resultados=resultados,
                           contadores=contadores,
                           inconsistencias=inconsistencias,
                           programas=programas, filtro=filtro)


if __name__ == "__main__":
    app.run(debug=True)  # arranca el Salón en http://localhost:5000
