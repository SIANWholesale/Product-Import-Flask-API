import pandas as pd
import requests
from flask import Flask, request, send_file, jsonify
from io import BytesIO

app = Flask(__name__)

@app.route('/process_excel', methods=['POST'])
def process_excel():
    try:
        # Get JSON data from Zapier
        data = request.get_json()
        file_url = data.get("file_url")

        if not file_url:
            return jsonify({"error": "No file URL provided"}), 400

        # Download file from OneDrive link
        response = requests.get(file_url)
        if response.status_code != 200:
            return jsonify({"error": f"Failed to download file. Status: {response.status_code}"}), 400

        # Read Excel file into Pandas DataFrame
        df = pd.read_excel(BytesIO(response.content), sheet_name=0, dtype=str)

        # Fix formatting issues (replace non-breaking spaces)
        df = df.map(lambda x: x.replace('\xa0', ' ') if isinstance(x, str) else x)

        # Convert DataFrame to CSV and store in memory
        output = BytesIO()
        df.to_csv(output, index=False, encoding='utf-8-sig')
        output.seek(0)

        # Return CSV file with correct headers
        return send_file(
            output,
            mimetype='text/csv',
            as_attachment=True,
            download_name="import_product.csv"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
