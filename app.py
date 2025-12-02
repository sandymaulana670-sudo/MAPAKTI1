from flask import Flask, request, jsonify, render_template
import pandas as pd
import os

app = Flask(__name__, template_folder="templates")

# Path ke file excel
EXCEL_FILE = "data/peserta.xlsx"

def load_excel():
    """Load file excel setiap kali ada request."""
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=["nomor", "nama", "status", "alasan"])
    return pd.read_excel(EXCEL_FILE, dtype=str).fillna("")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/cek", methods=["GET"])
def cek():
    nomor = request.args.get("nomor", "").strip()

    if nomor == "":
        return jsonify({"error": "Nomor tidak boleh kosong"}), 400

    df = load_excel()

    # Cocokkan nomor EXACT
    hasil = df[df["nomor"].astype(str).str.strip() == nomor]

    if hasil.empty:
        return jsonify({"found": False, "message": "Nomor tidak ditemukan"}), 404

    row = hasil.iloc[0]

    return jsonify({
        "found": True,
        "nomor": row["nomor"],
        "nama": row["nama"],
        "status": row["status"],
        "alasan": row["alasan"]
    })


if __name__ == "__main__":
    app.run(debug=True)
