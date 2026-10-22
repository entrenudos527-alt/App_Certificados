# PRUEBAS AUTOMÁTICAS de la app de certificados (pytest).
# Se ejecutan con:  python -m pytest
# No abren navegador ni servidor: usan el cliente de pruebas de Flask, que
# simula las peticiones HTTP por dentro.

import io
import os
import re
from pathlib import Path

# --- Credenciales SOLO para las pruebas ---
# Se definen ANTES de importar app, para que la app las lea al arrancar.
# Así las pruebas no dependen del archivo .env de la persona.
os.environ["SECRET_KEY"] = "clave-secreta-solo-para-pruebas"
os.environ["ADMIN_CORREO"] = "admin@ejemplo.cr"
os.environ["ADMIN_CLAVE"] = "clave-admin-123"
os.environ["USUARIO_CORREO"] = "usuario@ejemplo.cr"
os.environ["USUARIO_CLAVE"] = "clave-usuario-123"

import openpyxl  # noqa: E402  (va después de las variables a propósito)
import pytest  # noqa: E402

import app as aplicacion  # noqa: E402  (la app que vamos a probar)

# VARIABLES Path: la carpeta del proyecto y la Bodega de Excel originales
RAIZ = Path(__file__).resolve().parent.parent
INSUMOS = RAIZ / "Insumos"


# FIXTURE: entrega un "navegador falso" nuevo para cada prueba
@pytest.fixture
def cliente():
    aplicacion.app.config.update(TESTING=True)
    with aplicacion.app.test_client() as navegador:
        yield navegador


# FUNCIÓN auxiliar: lee un contador del HTML del tablero
def leer_contador(texto, nombre):
    patron = rf"<strong>(\d+)</strong><span>{nombre}</span>"
    encontrado = re.search(patron, texto)
    assert encontrado, f"No se encontró el contador «{nombre}» en la página"
    return int(encontrado.group(1))


# PRUEBA 1: sin sesión, la app no muestra nada y manda al login
def test_sin_sesion_redirige_a_login(cliente):
    # Simula abrir la página principal sin haber entrado nunca
    respuesta = cliente.get("/", follow_redirects=False)

    assert respuesta.status_code == 302  # redirección
    assert "/login" in respuesta.headers["Location"]  # va a la entrada


# PRUEBA 2: una clave equivocada no deja entrar y muestra el error
def test_clave_incorrecta_no_entra(cliente):
    # Simula escribir el correo del admin con una contraseña que no es
    respuesta = cliente.post("/login", data={
        "correo": "admin@ejemplo.cr",
        "clave": "esta-clave-es-incorrecta",
    })

    texto = respuesta.get_data(as_text=True)
    assert respuesta.status_code == 200  # se queda en la pantalla de entrada
    assert "Correo o contraseña incorrectos" in texto  # mensaje para la persona
    # Y sigue sin sesión: el tablero vuelve a mandar al login
    assert cliente.get("/").status_code == 302


# PRUEBA 3: un usuario normal no puede usar la carga de Excel
def test_usuario_no_puede_cargar(cliente):
    # Simula una sesión ya iniciada con rol "usuario"
    with cliente.session_transaction() as sesion:
        sesion["correo"] = "usuario@ejemplo.cr"
        sesion["rol"] = "usuario"

    respuesta = cliente.post("/cargar")

    assert respuesta.status_code == 403  # prohibido


# PRUEBA 4: cambiar notas en un Excel subido cambia los contadores
def test_excel_modificado_cambia_contadores(cliente):
    # Simula una sesión de administrador
    with cliente.session_transaction() as sesion:
        sesion["correo"] = "admin@ejemplo.cr"
        sesion["rol"] = "admin"

    # --- Construye el Registro de evaluaciones modificado EN MEMORIA ---
    # Se lee el Excel original y se cambia solo la copia que vive en RAM:
    # el archivo de Insumos NO se toca.
    original = INSUMOS / "Registro_Evaluaciones.xlsx"
    bytes_originales = original.read_bytes()  # para comprobar que no cambió

    libro = openpyxl.load_workbook(original)
    hoja = libro["Evaluaciones"]
    for fila in hoja.iter_rows(min_row=2):  # BUCLE: todas las evaluaciones
        if str(fila[0].value).strip() == "101230456":  # Camila Rojas Méndez
            fila[3].value = 40  # columna Nota: baja sus notas a 40

    memoria = io.BytesIO()  # caja en RAM, como la que usa la app
    libro.save(memoria)
    memoria.seek(0)

    # --- Sube los dos Excel al panel de administrador ---
    respuesta = cliente.post("/cargar", data={
        "accion": "pantalla",
        "maestro": (io.BytesIO(
            (INSUMOS / "Maestro_Estudiantes.xlsx").read_bytes()),
            "Maestro_Estudiantes.xlsx"),
        "evaluaciones": (memoria, "Registro_Evaluaciones.xlsx"),
    }, content_type="multipart/form-data")

    texto = respuesta.get_data(as_text=True)
    assert respuesta.status_code == 200
    # Camila pasa de Aprobación a Participación: 19/3/1 -> 18/4/1
    assert leer_contador(texto, "Aprobación") == 18
    assert leer_contador(texto, "Participación") == 4
    assert leer_contador(texto, "Sin certificado") == 1
    # 18 + 4 = 22 certificados descargables
    assert "Descargar 22 certificados" in texto

    # El archivo de la Bodega quedó intacto
    assert original.read_bytes() == bytes_originales
