from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import io
from comparador import load_data_xlsx, preprocess_dataframe, compare_dataframes, create_diff_excel

app = Flask(__name__)
CORS(app) # Enable CORS for all routes

@app.route("/")
def hello_world():
    return "Hello, Flask Backend!"

@app.route("/upload", methods=["POST"])
def upload_files():
    if "file1" not in request.files or "file2" not in request.files:
        return jsonify({"error": "Por favor, envie ambos os arquivos (file1 e file2)."}), 400

    file1 = request.files["file1"]
    file2 = request.files["file2"]
    
    # Assume has_header is always True for now, or pass it from frontend if needed
    has_header = True 

    try:
        # Process File 1
        df1_raw = load_data_xlsx(file1.read(), has_header)
        df1 = preprocess_dataframe(df1_raw, has_header, file1.filename)
        if df1 is None or df1.empty:
            return jsonify({"error": f"Não foi possível processar a Planilha 1 ({file1.filename}). Verifique o formato."}), 400

        # Process File 2
        df2_raw = load_data_xlsx(file2.read(), has_header)
        df2 = preprocess_dataframe(df2_raw, has_header, file2.filename)
        if df2 is None or df2.empty:
            return jsonify({"error": f"Não foi possível processar a Planilha 2 ({file2.filename}). Verifique o formato."}), 400

        # Compare DataFrames
        df_diffs = compare_dataframes(df1, df2)

        # Generate Excel file
        excel_output = create_diff_excel(df1, df2, df_diffs)

        return send_file(
            io.BytesIO(excel_output),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="comparacao_verbas.xlsx"
        )

    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as e:
        return jsonify({"error": f"Ocorreu um erro inesperado: {str(e)}"}), 500

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

