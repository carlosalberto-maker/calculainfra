"""
Aplicación web Flask — Calculadora de Infraestructura EDS (IMSS-Bienestar).

Sube un archivo Excel con las hojas Servicios, Personal y Red, y recibe:
  - Reporte HTML visual con el análisis completo (como tecamac_reportehtml.html)
  - Excel actualizado con los faltantes escritos en el bloque derecho de Red

Uso:
  pip install flask pandas openpyxl
  python app.py
  → Abrir http://localhost:5000
"""

import os
import uuid
import io
import tempfile
import shutil
from flask import Flask, render_template, request, send_file, redirect, url_for, flash

import calcular_infraestructura as calc

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "eds-calculadora-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

# Almacén en memoria para descarga única de archivos Excel
_download_store: dict[str, tuple[str, bytes]] = {}


# ---------------------------------------------------------------------------
# Ruta principal: formulario de carga
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


# ---------------------------------------------------------------------------
# Procesar archivo (todo en memoria / tempfiles, sin persistencia)
# ---------------------------------------------------------------------------
@app.route("/procesar", methods=["POST"])
def procesar():
    if "archivo" not in request.files:
        flash("No se envió ningún archivo.", "error")
        return redirect(url_for("index"))

    file = request.files["archivo"]
    if file.filename == "":
        flash("No se seleccionó ningún archivo.", "error")
        return redirect(url_for("index"))

    if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
        flash("El archivo debe ser un Excel (.xlsx o .xls).", "error")
        return redirect(url_for("index"))

    # Guardar archivo subido en un archivo temporal
    suffix = os.path.splitext(file.filename)[1] or ".xlsx"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    file.save(tmp.name)
    upload_path = tmp.name
    tmp.close()

    try:
        actual, requerido, faltantes, excesos, por_servicio, header_rows = \
            calc.procesar_archivo(upload_path, overrides=None)
    except Exception as e:
        os.unlink(upload_path)
        flash(f"Error al procesar el archivo: {e}", "error")
        return redirect(url_for("index"))

    # Generar reporte HTML en memoria
    html_content = calc.generar_html(
        upload_path, actual, requerido, faltantes, excesos, por_servicio
    )

    # Generar Excel actualizado en archivo temporal
    excel_tmp = upload_path + "_actualizado.xlsx"
    shutil.copy2(upload_path, excel_tmp)
    calc.actualizar_excel_red(
        excel_tmp, por_servicio, header_rows.get("Red", 4)
    )

    # Leer Excel a memoria y limpiar archivos temporales
    with open(excel_tmp, "rb") as f:
        excel_bytes = f.read()
    os.unlink(upload_path)
    os.unlink(excel_tmp)

    # Guardar en memoria para descarga única
    download_token = uuid.uuid4().hex[:16]
    base_name = os.path.splitext(file.filename)[0]
    excel_filename = f"{base_name}_actualizado.xlsx"
    _download_store[download_token] = (excel_filename, excel_bytes)

    # Reporte en consola del servidor
    calc.generar_reporte(actual, requerido, faltantes, excesos, por_servicio)

    return render_template("reporte.html",
                           html_content=html_content,
                           download_token=download_token)


# ---------------------------------------------------------------------------
# Descargar Excel actualizado (un solo uso, desde memoria)
# ---------------------------------------------------------------------------
@app.route("/descargar/<token>")
def descargar(token):
    entry = _download_store.pop(token, None)
    if entry is None:
        flash("El archivo ya no está disponible.", "error")
        return redirect(url_for("index"))

    filename, excel_bytes = entry
    return send_file(
        io.BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename
    )


# ---------------------------------------------------------------------------
# Punto de entrada
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    print("=" * 60)
    print("  Calculadora de Infraestructura EDS — Servidor Web")
    print("=" * 60)
    print(f"  Abre http://localhost:{port} en tu navegador")
    print("=" * 60)
    app.run(debug=debug, host="0.0.0.0", port=port)
