import os, uuid, hmac
from flask import Flask, render_template, request, send_file, Response, abort
from analisis import calcular, generar_excel

app = Flask(__name__, template_folder=".")
app.config["MAX_CONTENT_LENGTH"] = 60 * 1024 * 1024
TMP = "/tmp/resumenes"
os.makedirs(TMP, exist_ok=True)


@app.before_request
def auth():
    clave = os.environ.get("APP_PASSWORD")
    if not clave:
        return
    a = request.authorization
    if not (a and hmac.compare_digest(a.password or "", clave)):
        return Response("Acceso restringido", 401, {"WWW-Authenticate": 'Basic realm="Stock"'})


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("index.html", res=None)
    f = request.files.get("archivo")
    if not f or not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return render_template("index.html", res=None, error="Subí un archivo Excel (.xlsx)."), 400
    try:
        res = calcular(f.stream)
    except Exception as e:
        return render_template("index.html", res=None, error=str(e)), 400
    token = uuid.uuid4().hex
    generar_excel(res, os.path.join(TMP, f"{token}.xlsx"))
    return render_template("index.html", res=res, token=token, nombre=f.filename)


@app.route("/descargar/<token>")
def descargar(token):
    if not token.isalnum():
        abort(404)
    ruta = os.path.join(TMP, f"{token}.xlsx")
    if not os.path.exists(ruta):
        abort(404)
    return send_file(ruta, as_attachment=True, download_name="Resumen_Pallets_Stock.xlsx")


if __name__ == "__main__":
    app.run(debug=True)
