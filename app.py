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
import shutil
from flask import Flask, render_template, request, send_file, redirect, url_for, flash

import calcular_infraestructura as calc

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "eds-calculadora-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


# ---------------------------------------------------------------------------
# Ruta principal: formulario de carga
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


# ---------------------------------------------------------------------------
# Procesar archivo
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

    session_id = uuid.uuid4().hex[:12]
    upload_path = os.path.join(UPLOAD_FOLDER, f"{session_id}_{file.filename}")
    file.save(upload_path)

    # Recoger overrides del formulario
    overrides = {}
    for campo in ("comp_req", "imp_req", "ap_req", "sw24_req"):
        val = request.form.get(campo, "").strip()
        if val:
            try:
                overrides[campo] = int(val)
            except ValueError:
                pass

    try:
        actual, requerido, faltantes, excesos, por_servicio, header_rows = \
            calc.procesar_archivo(upload_path, overrides or None)
    except Exception as e:
        flash(f"Error al procesar el archivo: {e}", "error")
        return redirect(url_for("index"))

    # Generar reporte HTML
    html_string = calc.generar_html(
        upload_path, actual, requerido, faltantes, excesos, por_servicio
    )

    # Guardar el HTML en outputs con nombre único
    base_name = os.path.splitext(file.filename)[0]
    html_filename = f"{session_id}_{base_name}_reporte.html"
    html_path = os.path.join(OUTPUT_FOLDER, html_filename)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_string)

    # Generar Excel actualizado (copia + escritura bloque derecho)
    excel_filename = f"{session_id}_{base_name}_actualizado.xlsx"
    excel_path = os.path.join(OUTPUT_FOLDER, excel_filename)
    shutil.copy2(upload_path, excel_path)
    calc.actualizar_excel_red(
        excel_path, por_servicio, header_rows.get("Red", 4)
    )

    # Reporte en consola del servidor
    calc.generar_reporte(actual, requerido, faltantes, excesos, por_servicio)

    return redirect(url_for("reporte", filename=html_filename,
                            excel_filename=excel_filename))


# ---------------------------------------------------------------------------
# Ver reporte HTML
# ---------------------------------------------------------------------------
@app.route("/reporte/<filename>")
def reporte(filename):
    excel_filename = request.args.get("excel_filename", "")
    html_path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(html_path):
        flash("El reporte ya no está disponible.", "error")
        return redirect(url_for("index"))

    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    return render_template("reporte.html",
                           html_content=html_content,
                           excel_filename=excel_filename)


# ---------------------------------------------------------------------------
# Descargar Excel actualizado
# ---------------------------------------------------------------------------
@app.route("/descargar/<filename>")
def descargar(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        flash("El archivo ya no está disponible.", "error")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True)


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
