# LA COCINA de Academia Horizonte (backend con Flask)
# Toda la lógica vive aquí. El HTML (Salón) solo muestra lo que aquí se calcula.
from datetime import date  # VARIABLE: fecha de emisión del certificado
from functools import wraps  # VARIABLE: herramienta para crear decoradores
from io import BytesIO  # VARIABLE: memoria para armar el PDF y el ZIP sin tocar disco
import os  # VARIABLE: lee variables de entorno (secretos de Vercel)
from pathlib import Path  # VARIABLE: rutas que funcionan desde cualquier carpeta
import unicodedata  # VARIABLE: quita tildes para nombres de archivo seguros
import zipfile  # VARIABLE: empaqueta varios PDF en un solo ZIP

from flask import (Flask, abort, redirect, render_template, request,  # web
                   send_file, session, url_for)
from fpdf import FPDF  # VARIABLE: librería que dibuja los certificados en PDF
import openpyxl  # VARIABLE: librería para leer los Excel de la Bodega
from openpyxl.styles import Alignment, Font, PatternFill  # estilos del Excel
from openpyxl.utils.exceptions import InvalidFileException  # Excel inválido
from werkzeug.security import check_password_hash  # compara claves con hash

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

# ---------------------------------------------------------------------------
# SESIÓN Y PERMISOS (quién entró y qué puede ver)
# ---------------------------------------------------------------------------

# VARIABLE: clave con la que Flask firma la cookie de sesión.
# En Vercel debe venir de la variable de entorno SECRET_KEY.
# Si no existe (por ejemplo en tu computadora) se genera una al azar:
# las sesiones se pierden al reiniciar, pero no queda ningún secreto en el código.
app.secret_key = os.environ.get("SECRET_KEY") or os.urandom(24)

# LÍMITE: 4 MB por petición (Vercel corta en 4.5 MB; nos quedamos debajo).
# Se usa también al revisar cada archivo subido.
LIMITE_BYTES = 4 * 1024 * 1024
app.config["MAX_CONTENT_LENGTH"] = LIMITE_BYTES

# TIPO MIME: el nombre técnico del formato .xlsx para send_file
MIME_XLSX = ("application/vnd.openxmlformats-officedocument"
             ".spreadsheetml.sheet")

# TEXTO: aviso que se muestra cuando los resultados vienen de archivos subidos
AVISO_TEMPORAL = ("Resultados calculados con los archivos que acabas de subir. "
                  "Viven solo en esta pantalla: no se guardaron. Si recargas o "
                  "vuelves al tablero, la app regresa a los datos de fábrica "
                  "(Insumos/), porque todavía no tiene dónde guardarlos.")



# FUNCIÓN: lee del entorno las cuentas permitidas (sin escribir claves en el código).
# Se leen en cada petición para que un cambio en Vercel no exija tocar el archivo.
# Formato:
#   ADMIN_CORREO / ADMIN_CLAVE   (o ADMIN_CLAVE_HASH)   -> rol "admin"
#   USUARIO_CORREO / USUARIO_CLAVE (o USUARIO_CLAVE_HASH) -> rol "usuario"
def usuarios_permitidos():
    usuarios = {}  # VARIABLE dict: correo -> datos de la cuenta
    for prefijo, rol in (("ADMIN", "admin"), ("USUARIO", "usuario")):  # BUCLE
        correo = os.environ.get(f"{prefijo}_CORREO", "").strip().lower()
        clave = os.environ.get(f"{prefijo}_CLAVE", "")
        clave_hash = os.environ.get(f"{prefijo}_CLAVE_HASH", "")
        if correo and (clave or clave_hash):  # CONDICIONAL: cuenta configurada
            usuarios[correo] = {"clave": clave, "clave_hash": clave_hash,
                                "rol": rol}
    return usuarios


# FUNCIÓN: compara la clave escrita con la guardada.
# Si hay hash (recomendado en producción) lo verifica; si no, compara el texto.
def comprobar_clave(usuario, escrita):
    if usuario["clave_hash"]:  # CONDICIONAL: producción con hash
        try:
            return check_password_hash(usuario["clave_hash"], escrita)
        except ValueError:  # hash mal formado en la variable de entorno
            return False
    return usuario["clave"] == escrita  # texto plano (útil en pruebas)


# FUNCIÓN decoradora: exige haber iniciado sesión.
# Si no hay sesión, manda a la pantalla de entrada.
def requiere_sesion(vista):
    @wraps(vista)
    def envoltura(*args, **kwargs):
        if not session.get("rol"):  # CONDICIONAL: nadie ha entrado
            return redirect(url_for("login"))
        return vista(*args, **kwargs)
    return envoltura


# FUNCIÓN decoradora: deja pasar solo si la sesión tiene rol "admin".
# Se pone encima de una ruta así:  @app.route(...)  @solo_admin  def ...
def solo_admin(vista):
    @wraps(vista)  # conserva el nombre de la ruta original
    def envoltura(*args, **kwargs):
        rol = session.get("rol")  # VARIABLE string: el rol guardado al entrar
        if not rol:  # CONDICIONAL: no hay sesión -> a la pantalla de entrada
            return redirect(url_for("login"))
        if rol != "admin":  # CONDICIONAL: hay sesión pero no es administrador
            # Responde 403 con una página amable, no un error crudo
            return render_template("sin_permiso.html",
                                   correo=session.get("correo")
                                   or "Sin sesión iniciada",
                                   rol=rol or "invitado"), 403
        return vista(*args, **kwargs)  # es admin: sigue a la ruta de verdad
    return envoltura



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


# FUNCIÓN: lee la Bodega y prepara resultados + inconsistencias para el Salón.
# Acepta dos archivos opcionales EN MEMORIA (los subidos por un administrador).
# Si no llega ninguno, lee los Excel de la carpeta Insumos (datos de fábrica).
def cargar_datos(archivo_maestro=None, archivo_evaluaciones=None):
    # VARIABLES: la fuente de cada Excel (ruta del disco o caja en memoria)
    fuente_maestro = (archivo_maestro if archivo_maestro is not None
                      else INSUMOS_DIR / "Maestro_Estudiantes.xlsx")
    fuente_evaluaciones = (archivo_evaluaciones if archivo_evaluaciones is not None
                           else INSUMOS_DIR / "Registro_Evaluaciones.xlsx")

    # --- Maestro: lista oficial de estudiantes ---
    wb_m = openpyxl.load_workbook(fuente_maestro,
                                  read_only=True, data_only=True)
    ws_m = wb_m["Estudiantes"]
    maestro = {}  # VARIABLE de tipo diccionario: clave (id, programa)
    for fila in ws_m.iter_rows(min_row=2, values_only=True):  # BUCLE
        ident = str(fila[0]).strip()  # VARIABLE string: preserva ceros
        maestro[(ident, fila[3])] = {"nombre": fila[1], "programa": fila[3]}
    wb_m.close()

    # --- Evaluaciones: notas y asistencia por módulo ---
    wb_e = openpyxl.load_workbook(fuente_evaluaciones,
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


# FUNCIÓN: pasa un archivo subido a una "caja" en memoria, sin tocar el disco.
# ANALOGÍA de BytesIO: es como una caja de cartón dentro de la memoria RAM.
# En vez de fotocopiar el Excel al archivador del disco (que en Vercel está
# cerrado con llave), metemos la fotocopia en la caja. La usamos durante esta
# petición y, al terminar, la caja se tira: no queda nada guardado.
def abrir_en_memoria(archivo):
    contenido = archivo.read()  # VARIABLE bytes: todo el archivo en RAM
    if len(contenido) > LIMITE_BYTES:  # CONDICIONAL: límite de 4 MB
        raise ValueError("El archivo pesa más de 4 MB. Reduce el Excel y "
                         "vuelve a intentarlo (Vercel no acepta envíos "
                         "mayores a 4.5 MB).")
    return BytesIO(contenido)  # TIPO BytesIO: la caja lista para openpyxl


# FUNCIÓN: arma un Excel con los resultados y lo devuelve en memoria
def crear_excel_resultados(resultados):
    libro = openpyxl.Workbook()  # VARIABLE: libro nuevo, en blanco
    hoja = libro.active
    hoja.title = "Resultados"

    # Fila de títulos (el "header" de la tabla)
    hoja.append(["Identificación", "Nombre", "Programa", "Módulos",
                 "Promedio", "Asistencia", "Certificado"])

    for r in resultados:  # BUCLE: una fila por estudiante y programa
        hoja.append([r["identificacion"], r["nombre"], r["programa"],
                     r["modulos"], r["promedio"], r["asistencia"],
                     r["tipo"]])

    # Estilo del encabezado: azul marino con letras doradas, como la app
    relleno = PatternFill("solid", start_color="0F2A44")
    for celda in hoja[1]:  # BUCLE: cada título de columna
        celda.font = Font(bold=True, color="D4AF37")
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center")

    # Anchos de columna para que se lea cómodo
    anchos = [16, 30, 32, 10, 10, 12, 18]
    for letra, ancho in zip("ABCDEFG", anchos):  # BUCLE
        hoja.column_dimensions[letra].width = ancho
    hoja.freeze_panes = "A2"  # deja fija la fila de títulos al bajar

    memoria = BytesIO()  # VARIABLE: otra caja en RAM para el Excel terminado
    libro.save(memoria)  # se guarda en la caja, NO en el disco
    memoria.seek(0)  # rebobina para que Flask lo lea desde el inicio
    return memoria



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


# ORIGEN por defecto: lo que vive en la Bodega (carpeta Insumos)
ORIGEN_FABRICA = "Datos de fábrica (carpeta Insumos)"


# FUNCIÓN: entrega la plantilla del tablero con todo lo que necesita el Salón
def render_tablero(resultados, contadores, inconsistencias, programas,
                   filtro="Todos", origen=ORIGEN_FABRICA,
                   aviso=None, error=None):
    # CONDICIONAL: el filtro de programa solo afecta la tabla, no los contadores
    if filtro != "Todos":
        resultados = [r for r in resultados if r["programa"] == filtro]
    # VARIABLE entero: cuántos PDF se pueden descargar con el filtro actual
    total_emitibles = len([r for r in resultados if r["tipo"] != "Sin certificado"])
    # VARIABLES de sesión: quién entró y qué rol tiene (None si no hay sesión)
    correo = session.get("correo") or "Sin sesión iniciada"
    rol = session.get("rol")
    es_admin = rol == "admin"  # CONDICIONAL: booleano para la plantilla
    return render_template("index.html", resultados=resultados,
                           contadores=contadores,
                           inconsistencias=inconsistencias,
                           programas=programas, filtro=filtro,
                           total_emitibles=total_emitibles,
                           correo=correo, rol=rol, es_admin=es_admin,
                           origen=origen, aviso=aviso, error=error)


# FUNCIÓN: tablero con los datos de fábrica (también se usa al mostrar errores)
def tablero_de_fabrica(error=None):
    resultados, contadores, inconsistencias, programas = cargar_datos()
    return render_tablero(resultados, contadores, inconsistencias, programas,
                          error=error)


# Ruta: pantalla de entrada (correo y contraseña)
@app.route("/login", methods=["GET", "POST"])
def login():
    # Si ya hay sesión, no tiene sentido volver a pedir la clave
    if session.get("rol"):  # CONDICIONAL
        return redirect(url_for("index"))

    error = None  # VARIABLE string: mensaje rojo si algo falla
    correo = ""  # VARIABLE string: lo que la persona escribió

    if request.method == "POST":  # CONDICIONAL: envió el formulario
        correo = request.form.get("correo", "").strip().lower()
        clave = request.form.get("clave", "")
        usuario = usuarios_permitidos().get(correo)  # FUNCIÓN
        if usuario and comprobar_clave(usuario, clave):  # FUNCIÓN
            session.clear()  # sesión limpia antes de guardar al usuario
            session["correo"] = correo
            session["rol"] = usuario["rol"]
            return redirect(url_for("index"))
        error = ("Correo o contraseña incorrectos. Revisa los datos e "
                 "inténtalo de nuevo.")

    return render_template("login.html", error=error, correo=correo)


# Ruta principal: prepara los datos y los sirve al Salón
@app.route("/")
@requiere_sesion
def index():
    resultados, contadores, inconsistencias, programas = cargar_datos()
    filtro = request.args.get("programa", "Todos")  # VARIABLE string del filtro
    return render_tablero(resultados, contadores, inconsistencias, programas,
                          filtro=filtro)


# Ruta: recibe los dos Excel del panel de administrador.
# Todo pasa EN MEMORIA: los archivos se leen con request.files y se meten en
# BytesIO. Nada se escribe en el disco (en Vercel es de solo lectura).
@app.route("/cargar", methods=["POST"])
@solo_admin
def cargar():
    accion = request.form.get("accion", "pantalla")  # VARIABLE string
    archivo_maestro = request.files.get("maestro")
    archivo_evaluaciones = request.files.get("evaluaciones")

    # CONDICIONAL: deben venir los dos archivos
    if archivo_maestro is None or archivo_maestro.filename == "":
        return tablero_de_fabrica("Falta el Maestro de estudiantes. Elige los "
                                  "dos archivos y vuelve a intentarlo.")
    if archivo_evaluaciones is None or archivo_evaluaciones.filename == "":
        return tablero_de_fabrica("Falta el Registro de evaluaciones. Elige "
                                  "los dos archivos y vuelve a intentarlo.")

    # VARIABLES string: nombres limpios para mostrarlos en el panel
    nombre_maestro = Path(archivo_maestro.filename).name
    nombre_evaluaciones = Path(archivo_evaluaciones.filename).name

    try:
        # Las dos "cajas" en memoria que openpyxl va a leer
        caja_maestro = abrir_en_memoria(archivo_maestro)
        caja_evaluaciones = abrir_en_memoria(archivo_evaluaciones)
        # MISMA función de siempre: no se duplica la lógica de negocio
        resultados, contadores, inconsistencias, programas = cargar_datos(
            caja_maestro, caja_evaluaciones)
    except ValueError as error:  # por ejemplo: archivo de más de 4 MB
        return tablero_de_fabrica(str(error))
    except (zipfile.BadZipFile, InvalidFileException, KeyError, TypeError):
        # No es un .xlsx real, está dañado o no tiene las hojas esperadas
        return tablero_de_fabrica(
            "Uno de los archivos no parece un Excel válido (.xlsx). Revisa "
            "que no esté dañado y que tenga las hojas «Estudiantes» y "
            "«Evaluaciones».")

    # VARIABLE string: de dónde salieron los números que se ven en pantalla
    origen = f"Archivos subidos: {nombre_maestro} y {nombre_evaluaciones}"

    # CONDICIONAL: el segundo botón pide el Excel en vez de la pantalla
    if accion == "excel":
        memoria = crear_excel_resultados(resultados)  # FUNCIÓN
        nombre_salida = (f"Resultados_"
                         f"{texto_para_archivo(Path(nombre_maestro).stem)}.xlsx")
        return send_file(memoria, mimetype=MIME_XLSX, as_attachment=True,
                         download_name=nombre_salida)

    # Botón "Procesar y ver en pantalla"
    return render_tablero(resultados, contadores, inconsistencias, programas,
                          origen=origen, aviso=AVISO_TEMPORAL)


# Ruta: descarga en Excel los resultados de los datos de fábrica
@app.route("/resultados.xlsx")
@solo_admin
def resultados_excel():
    resultados, _, _, _ = cargar_datos()
    memoria = crear_excel_resultados(resultados)  # FUNCIÓN
    return send_file(memoria, mimetype=MIME_XLSX, as_attachment=True,
                     download_name="Resultados_Academia_Horizonte.xlsx")


# Ruta: sección completa de inconsistencias, SOLO para administradores.
# El decorador @solo_admin es la protección real: si alguien escribe la URL
# a mano sin ser admin, recibe 403 y no ve la lista.
@app.route("/inconsistencias")
@solo_admin
def inconsistencias():
    _, _, lista, _ = cargar_datos()
    return render_template("inconsistencias.html", inconsistencias=lista,
                           correo=session.get("correo") or "Sin sesión iniciada",
                           rol=session.get("rol") or "invitado")


# Ruta: cierra la sesión y vuelve al tablero
@app.route("/logout")
def logout():
    session.clear()  # borra correo y rol de la cookie firmada
    return redirect(url_for("index"))


# MANEJADOR de error: Vercel (o Flask) corta los envíos de más de 4 MB.
# En vez de una página en blanco, se muestra el tablero con un aviso rojo.
@app.errorhandler(413)
def envio_demasiado_grande(error):
    return tablero_de_fabrica(
        "El envío pesa más de 4 MB. Vercel rechaza las peticiones mayores a "
        "4.5 MB, así que reduce o divide el Excel y vuelve a intentarlo."), 413


# Ruta: descarga el PDF de un estudiante (solo si tiene certificado)
@app.route("/certificado/<identificacion>")
@requiere_sesion
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
@requiere_sesion
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
