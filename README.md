# App_Certificados — Academia Horizonte

Aplicación web para emitir los certificados de una cohorte a partir de los Excel
de `Insumos/`.

## Qué hace

- Lee `Maestro_Estudiantes.xlsx` y `Registro_Evaluaciones.xlsx` (solo lectura).
- Calcula el promedio y la asistencia por estudiante y programa.
- Decide el veredicto: **Aprobación** (promedio >= 70 y asistencia >= 80),
  **Participación** (promedio < 70 y asistencia >= 80) o **Sin certificado**
  (asistencia < 80).
- Muestra el tablero en `/` con contadores, filtro por programa e inconsistencias.
- Descarga el PDF de un estudiante en `/certificado/<identificacion>?programa=<programa>`.
- Descarga todos los certificados del filtro en un ZIP desde `/certificados.zip?programa=<programa>`.

Con los datos actuales se generan **22 certificados**: 19 de Aprobación y
3 de Participación. El estudiante sin certificado no se incluye.

## Cómo usarla en tu computadora

1. Abre esta carpeta en el Explorador de Windows.
2. Haz doble clic en `iniciar_app.bat`. Se abre una ventana negra con el servidor.
3. En el navegador entra a <http://127.0.0.1:5000>.
4. Elige el programa en **Filtrar por programa** (o deja **Todos**).
5. Pulsa **Descargar N certificados (ZIP)** para bajar todos, o
   **Descargar PDF** en la fila de un estudiante para bajar solo ese.
6. El archivo queda en tu carpeta de **Descargas**. El ZIP contiene un PDF por
   estudiante.
7. Para apagar la app, cierra la ventana negra o presiona `Ctrl + C`.

## Publicar los cambios en el enlace de Vercel

El sitio <https://app-certificados-rouge.vercel.app/> se actualiza solo cuando
los cambios llegan a GitHub (`main`). Desde esta carpeta:

```powershell
git add app.py requirements.txt templates/index.html static/style.css README.md
git commit -m "Genera certificados PDF y descarga en ZIP"
git push
```

Vercel detecta el push y vuelve a publicar en uno o dos minutos. La nueva versión
del enlace ya tendrá el botón de descarga.
