 
from flask import Flask, request, jsonify, send_file
import pandas as pd
import requests
from io import BytesIO

app = Flask(__name__)

@app.route('/process_excel', methods=['POST'])
def process_excel():
    try:
        if request.content_type != 'application/json':
            return jsonify({"error": "Request must be JSON"}), 415

        data = request.get_json()
        file_url = data.get("file_url")

        if not file_url:
            return jsonify({"error": "No file URL provided"}), 400

        response = requests.get(file_url)
        response.raise_for_status()
        excel_file = BytesIO(response.content)

        df = pd.read_excel(excel_file, sheet_name=0, dtype=str)
        df = df.applymap(lambda x: x.replace('\xa0', ' ') if isinstance(x, str) else x)

        csv_output = BytesIO()
        df.to_csv(csv_output, index=False, encoding='utf-8-sig')
        csv_output.seek(0)

        return send_file(csv_output, mimetype='text/csv', as_attachment=True, download_name='import_product.csv')

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
