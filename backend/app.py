from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pathlib import Path
import os
import uuid
import traceback

from extraire_releve_attijari import (
    extraire_operations,
    construire_classeur,
)

app = Flask(__name__)
from flask_cors import CORS

CORS(
    app,
    resources={r"/*": {"origins": "*"}}
)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


@app.route("/", methods=["GET"])
def home():
    return jsonify({
        "status": "OK",
        "message": "Attijari PDF to Excel API"
    })


@app.route("/convert", methods=["POST"])
def convert():

    try:

        # Vérifier qu'un fichier a été envoyé
        if "file" not in request.files:
            return jsonify({
                "success": False,
                "error": "Aucun fichier reçu."
            }), 400

        file = request.files["file"]

        if file.filename == "":
            return jsonify({
                "success": False,
                "error": "Aucun fichier sélectionné."
            }), 400

        # Vérifier l'extension
        if not file.filename.lower().endswith(".pdf"):
            return jsonify({
                "success": False,
                "error": "Le fichier doit être un PDF."
            }), 400

        # Nom sécurisé
        filename = secure_filename(file.filename)

        unique_id = str(uuid.uuid4())

        pdf_path = os.path.join(
            UPLOAD_FOLDER,
            f"{unique_id}.pdf"
        )

        excel_path = os.path.join(
            OUTPUT_FOLDER,
            f"{unique_id}.xlsx"
        )

        # Sauvegarde du PDF
        file.save(pdf_path)

        # Extraction
        operations, resume = extraire_operations(pdf_path)

        # Génération Excel
        construire_classeur(
            operations,
            resume,
            excel_path
        )

        # Retour du fichier Excel
        return send_file(
            excel_path,
            as_attachment=True,
            download_name=Path(filename).stem + ".xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        print(traceback.format_exc())

        return jsonify({
            "success": False,
            "error": str(e)
        }), 500


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "healthy"
    })


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )
