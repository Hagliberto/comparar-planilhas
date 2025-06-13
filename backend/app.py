from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import io
import pandas as pd
from comparador import preprocess_dataframe, compare_dataframes, create_diff_excel

app = Flask(__name__)
CORS(app)

def load_group(files, headers_flags):
    dfs = []
    for f, has_header in zip(files, headers_flags):
        df = pd.read_excel(f, header=0 if has_header=='true' else None)
        df = preprocess_dataframe(df, has_header=='true', f.filename)
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

@app.route("/upload-json", methods=["POST"])
def upload_json():
    nucleo_files   = request.files.getlist('files_nucleo')
    unidades_files = request.files.getlist('files_unidades')
    nucleo_flags   = request.form.getlist('headers_nucleo')
    unidades_flags = request.form.getlist('headers_unidades')

    if not nucleo_files or not unidades_files:
        return jsonify(error="Envie ao menos um arquivo em cada grupo."),400

    df1 = load_group(nucleo_files, nucleo_flags)
    df2 = load_group(unidades_files, unidades_flags)
    diffs = compare_dataframes(df1, df2)
    return jsonify(diffs=diffs.to_dict(orient="records"))

@app.route("/upload", methods=["POST"])
def upload_file():
    nucleo_files   = request.files.getlist('files_nucleo')
    unidades_files = request.files.getlist('files_unidades')
    nucleo_flags   = request.form.getlist('headers_nucleo')
    unidades_flags = request.form.getlist('headers_unidades')

    df1 = load_group(nucleo_files, nucleo_flags)
    df2 = load_group(unidades_files, unidades_flags)
    diffs = compare_dataframes(df1, df2)
    excel_bytes = create_diff_excel(df1, df2, diffs)

    return send_file(
        io.BytesIO(excel_bytes),
        mimetype=(
          "application/vnd.openxmlformats-officedocument."
          "spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name="comparacao_verbas.xlsx"
    )

if __name__ == "__main__":
    app.run(port=5000, debug=True)
