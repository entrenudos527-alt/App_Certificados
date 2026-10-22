# LA COCINA de Academia Horizonte (backend con Flask)
# Toda la lógica vive aquí. El HTML (Salón) solo muestra lo que aquí se calcula.
from datetime import date  # VARIABLE: fecha de emisión del certificado
from io import BytesIO  # VARIABLE: memoria para armar el PDF y el ZIP sin tocar disco
from pathlib import Path  # VARIABLE: rutas que funcionan desde cualquier carpeta
import unicodedata  # VARIABLE: quita tildes para nombres de archivo seguros
import zipfile  # VARIABLE: empaqueta varios PDF en un solo ZIP

from flask import (Flask, abort, render_template, request,  # herramientas web
                   send_file, url_for)
from fpdf import FPDF  # VARIABLE: librería que dibuja los certificados en PDF
import openpyxl  # VARIABLE: librería para leer los Excel de la Bodega

app = Flask(__name__)  # VARIABLE de tipo Flask: la aplicación web

# Colores de la marca (azul marino y dorado), en formato RGB
AZUL_MARINO = (15, 42, 68)
AZUL_CAJA = (23, 53, 84)
DORADO = (212, 175, 55)
CREMA = (245, 241, 230)
GRIS_CLARO = (210, 214, 220)

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "setiembre", "octubre", "noviembre", "diciembre"]

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


# FUNCIÓN: quita tildes y caracteres raros para que el nombre del archivo sea seguro
def texto_para_archivo(texto):
    sin_tildes = unicodedata.normalize("NFKD", str(texto))
    limpio = sin_tildes.encode("ascii", "ignore").decode("ascii")
    return "".join(c for c in limpio if c.isalnum() or c in " _-").strip()


# FUNCIÓN: arma el nombre del PDF de un estudiante
def nombre_archivo(resultado):
    tipo = texto_para_archivo(resultado["tipo"]).replace(" ", "_")
    nombre = texto_para_archivo(resultado["nombre"]).replace(" ", "_")
    return f"Certificado_{tipo}_{nombre}.pdf"


# FUNCIÓN: busca el resultado de un estudiante y programa concretos
def buscar_resultado(identificacion, programa):
    resultados, _, _, _ = cargar_datos()  # reutiliza el mismo cálculo del tablero
    for r in resultados:  # BUCLE: recorre los 23 resultados
        if r["identificacion"] == identificacion and r["programa"] == programa:
            return r  # lo encontró
    return None  # no existe esa pareja (id + programa)


# FUNCIÓN: convierte una fecha en texto español, por ejemplo "21 de octubre de 2026"
def fecha_en_espanol(fecha):
    return f"{fecha.day} de {MESES[fecha.month - 1]} de {fecha.year}"


# FUNCIÓN: dibuja un certificado completo y lo devuelve como objeto PDF
def crear_certificado(resultado):
    pdf = FPDF(orientation="L", unit="mm", format="A4")  # hoja horizontal
    pdf.set_auto_page_break(False)
    pdf.add_page()
    ancho, alto = 297, 210  # medidas de A4 horizontal en milímetros

    # --- Fondo azul marino y doble marco dorado ---
    pdf.set_fill_color(*AZUL_MARINO)
    pdf.rect(0, 0, ancho, alto, "F")
    pdf.set_draw_color(*DORADO)
    pdf.set_line_width(1.2)
    pdf.rect(8, 8, ancho - 16, alto - 16)
    pdf.set_line_width(0.4)
    pdf.rect(11, 11, ancho - 22, alto - 22)

    # --- Encabezado de la academia ---
    pdf.set_text_color(*DORADO)
    pdf.set_font("Times", "B", 16)
    pdf.set_xy(0, 21)
    pdf.cell(ancho, 8, "ACADEMIA HORIZONTE", align="C")

    pdf.set_draw_color(*DORADO)
    pdf.set_line_width(0.6)
    pdf.line(108, 31, 189, 31)  # línea fina bajo el nombre

    # --- Título y tipo de certificado ---
    pdf.set_font("Times", "B", 38)
    pdf.set_xy(0, 38)
    pdf.cell(ancho, 16, "CERTIFICADO", align="C")

    pdf.set_text_color(*CREMA)
    pdf.set_font("Times", "B", 17)
    pdf.set_xy(0, 55)
    pdf.cell(ancho, 8, f"DE {resultado['tipo'].upper()}", align="C")

    # --- Nombre del estudiante ---
    pdf.set_font("Times", "I", 12)
    pdf.set_xy(0, 69)
    pdf.cell(ancho, 6, "Se otorga a", align="C")

    tamano_nombre = 30 if len(resultado["nombre"]) <= 30 else 24  # CONDICIONAL
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Times", "B", tamano_nombre)
    pdf.set_xy(10, 77)
    pdf.cell(ancho - 20, 14, resultado["nombre"], align="C")

    # --- Programa y cohorte ---
    pdf.set_text_color(*GRIS_CLARO)
    pdf.set_font("Times", "", 12)
    pdf.set_xy(0, 96)
    if resultado["tipo"] == "Aprobación":  # CONDICIONAL: texto según el veredicto
        texto = "por haber cursado y aprobado satisfactoriamente el programa"
    else:
        texto = "por haber participado y completado el programa"
    pdf.cell(ancho, 6, texto, align="C")

    pdf.set_text_color(*DORADO)
    pdf.set_font("Times", "B", 19)
    pdf.set_xy(12, 104)
    pdf.cell(ancho - 24, 10, resultado["programa"], align="C")

    pdf.set_text_color(*GRIS_CLARO)
    pdf.set_font("Times", "", 11)
    pdf.set_xy(0, 115)
    pdf.cell(ancho, 6, "Cohorte 2026-A", align="C")

    # --- Caja con los datos académicos ---
    pdf.set_fill_color(*AZUL_CAJA)
    pdf.set_draw_color(*DORADO)
    pdf.set_line_width(0.3)
    pdf.rect(63, 124, 171, 18, "DF")
    pdf.set_text_color(*CREMA)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_xy(63, 129)
    pdf.cell(171, 8,
             f"Promedio: {resultado['promedio']:.2f}     |     "
             f"Asistencia: {resultado['asistencia']:.2f} %     |     "
             f"Módulos cursados: {resultado['modulos']}",
             align="C")

    # --- Fecha de emisión ---
    pdf.set_text_color(*GRIS_CLARO)
    pdf.set_font("Times", "", 11)
    pdf.set_xy(0, 152)
    pdf.cell(ancho, 6, f"Emitido el {fecha_en_espanol(date.today())}",
             align="C")

    # --- Firma y sello ---
    pdf.set_draw_color(*DORADO)
    pdf.set_line_width(0.5)
    pdf.line(108, 176, 189, 176)
    pdf.set_text_color(*CREMA)
    pdf.set_font("Times", "", 11)
    pdf.set_xy(108, 178)
    pdf.cell(81, 6, "Dirección Académica", align="C")

    pdf.set_draw_color(*DORADO)
    pdf.set_line_width(0.8)
    pdf.ellipse(236, 157, 26, 26)
    pdf.set_line_width(0.4)
    pdf.ellipse(239.5, 160.5, 19, 19)
    pdf.set_text_color(*DORADO)
    pdf.set_font("Times", "B", 14)
    pdf.set_xy(236, 166)
    pdf.cell(26, 8, "AH", align="C")

    return pdf  # TIPO FPDF: el certificado listo para guardar


# Ruta principal: prepara los datos y los sirve al Salón
@app.route("/")
def index():
    resultados, contadores, inconsistencias, programas = cargar_datos()
    filtro = request.args.get("programa", "Todos")  # VARIABLE string del filtro
    if filtro != "Todos":  # CONDICIONAL: filtra la tabla, no los contadores
        resultados = [r for r in resultados if r["programa"] == filtro]
    # VARIABLE entero: cuántos PDF se pueden descargar con el filtro actual
    total_emitibles = len([r for r in resultados if r["tipo"] != "Sin certificado"])
    return render_template("index.html", resultados=resultados,
                           contadores=contadores,
                           inconsistencias=inconsistencias,
                           programas=programas, filtro=filtro,
                           total_emitibles=total_emitibles)


# Ruta: descarga el PDF de un estudiante (solo si tiene certificado)
@app.route("/certificado/<identificacion>")
def certificado(identificacion):
    programa = request.args.get("programa", "")  # VARIABLE string
    resultado = buscar_resultado(identificacion, programa)  # FUNCIÓN
    if resultado is None:  # CONDICIONAL: no existe esa pareja
        abort(404)
    if resultado["tipo"] == "Sin certificado":  # CONDICIONAL: no le corresponde
        abort(403)
    pdf = crear_certificado(resultado)  # FUNCIÓN: dibuja el diploma
    return send_file(BytesIO(bytes(pdf.output())), mimetype="application/pdf",
                     as_attachment=True,
                     download_name=nombre_archivo(resultado))


# Ruta: descarga todos los certificados del filtro actual en un solo ZIP
@app.route("/certificados.zip")
def certificados_zip():
    resultados, _, _, _ = cargar_datos()
    filtro = request.args.get("programa", "Todos")  # VARIABLE string
    if filtro != "Todos":  # CONDICIONAL: respeta el filtro de la pantalla
        resultados = [r for r in resultados if r["programa"] == filtro]
    emitibles = [r for r in resultados if r["tipo"] != "Sin certificado"]

    memoria = BytesIO()  # VARIABLE: el ZIP se arma en memoria, no en disco
    with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as paquete:
        for r in emitibles:  # BUCLE: un PDF por estudiante
            pdf = crear_certificado(r)  # FUNCIÓN
            paquete.writestr(nombre_archivo(r), bytes(pdf.output()))
    memoria.seek(0)  # rebobina para que Flask lo lea desde el inicio

    if filtro == "Todos":
        nombre_zip = "Certificados_Academia_Horizonte.zip"
    else:
        nombre_zip = f"Certificados_{texto_para_archivo(filtro).replace(' ', '_')}.zip"
    return send_file(memoria, mimetype="application/zip", as_attachment=True,
                     download_name=nombre_zip)


if __name__ == "__main__":
    app.run(debug=True)  # arranca el Salón en http://localhost:5000
