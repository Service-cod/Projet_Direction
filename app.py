
from flask import Flask, render_template, request
import sqlite3
from pathlib import Path
from openpyxl import Workbook
from flask import send_file
app = Flask(__name__)

# Emplacement de la base de données
DB_PATH = Path(__file__).resolve().parent / "direction.db"

# Création de la table
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS directeurs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom TEXT NOT NULL,
                etablissement TEXT NOT NULL,
                date_envoi DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

# Page d'accueil
@app.route("/", methods=["GET", "POST"])
def accueil():
    if request.method == "POST":
        nom = request.form.get("nom", "").strip()
        etablissement = request.form.get("etablissement", "").strip()

        if not nom or not etablissement:
            return "Veuillez remplir tous les champs.", 400

        # Enregistrement dans la base de données
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(
                "INSERT INTO directeurs (nom, etablissement) VALUES (?, ?)",
                (nom, etablissement)
            )

        return render_template(
            "success.html",
            nom=nom,
            etablissement=etablissement
        )

    return render_template("index.html")

@app.route("/admin")
def admin():
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        directeurs = conn.execute(
            "SELECT * FROM directeurs ORDER BY id DESC"
        ).fetchall()

    return render_template(
        "admin.html",
        directeurs=directeurs
    )

@app.route("/export")
def export_excel():

    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row

        directeurs = conn.execute("""
            SELECT id, nom, etablissement, date_envoi
            FROM directeurs
            ORDER BY id DESC
        """).fetchall()

    # إنشاء ملف Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Réponses"

    # العناوين
    ws.append([
        "ID",
        "Nom du directeur",
        "Établissement",
        "Date d'envoi"
    ])

    # المعلومات
    for directeur in directeurs:
        ws.append([
            directeur["id"],
            directeur["nom"],
            directeur["etablissement"],
            directeur["date_envoi"]
        ])

    # حفظ الملف
    fichier = Path(__file__).resolve().parent / "reponses.xlsx"
    wb.save(fichier)

    return send_file(
        fichier,
        as_attachment=True,
        download_name="Reponses_Directeurs.xlsx"
    )
if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=5000, debug=True)