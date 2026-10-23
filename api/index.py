
from flask import Flask, request, send_file, jsonify
import pandas as pd
from io import BytesIO

app = Flask(__name__)


@app.route("/", methods=["POST"])
def remove_duplicates():
    uploaded_file = request.files.get("file")
    column = request.form.get("column")

    if uploaded_file is None:
        return jsonify({"error": "Please upload an Excel file."}), 400

    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.astype(str).str.strip()

        if df.empty:
            return jsonify({"error": "The Excel file has no data."}), 400

        if not column:
            return jsonify({"error": "Please select a column."}), 400

        if column not in df.columns:
            return jsonify({"error": "Invalid column selected."}), 400

        cleaned_df = df.drop_duplicates(
            subset=[column],
            keep="first"
        )

        removed = len(df) - len(cleaned_df)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            cleaned_df.to_excel(writer, index=False)

        output.seek(0)

        response = send_file(
            output,
            as_attachment=True,
            download_name="cleaned_excel.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

        response.headers["X-Rows-Removed"] = str(removed)
        response.headers["Access-Control-Expose-Headers"] = "X-Rows-Removed"

        return response

    except Exception as error:
        return jsonify({"error": str(error)}), 400
