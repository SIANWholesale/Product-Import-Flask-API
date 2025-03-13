import pandas as pd
import requests
import os
from dotenv import load_dotenv
from flask import Flask, request, jsonify
from io import BytesIO

# Load environment variables from .env (for local development)
load_dotenv()

# OneDrive API credentials (read from environment variables)
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
UPLOAD_PATH = os.getenv("UPLOAD_PATH", "/SIAN Marketing/Website Lists/import_product.csv")  # Default path

app = Flask(__name__)

def get_access_token():
    """Get an OAuth2 access token for OneDrive API."""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data)
    response_json = response.json()
    return response_json.get("access_token")

def upload_to_onedrive(csv_data):
    """Upload CSV file to OneDrive."""
    access_token = get_access_token()
    if not access_token:
        return {"error": "Failed to get OneDrive access token"}

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "text/csv"
    }
    
    # Upload file to OneDrive
    upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{UPLOAD_PATH}:/content"
    response = requests.put(upload_url, headers=headers, data=csv_data)

    if response.status_code in [200, 201]:
        return {"success": "File uploaded successfully"}
    else:
        return {"error": f"Failed to upload file. Status: {response.status_code}, Response: {response.text}"}

@app.route('/process_excel', methods=['POST'])
def process_excel():
    try:
        # Get file URL from Zapier request
        data = request.json
        file_url = data.get("file_url")

        if not file_url:
            return jsonify({"error": "No file URL provided"}), 400

        # Download file from OneDrive
        response = requests.get(file_url, stream=True)
        if response.status_code != 200:
            return jsonify({"error": f"Failed to download file. Status: {response.status_code}"}), 400

        # Read Excel into Pandas DataFrame
        df = pd.read_excel(BytesIO(response.content), sheet_name=0, dtype=str)

        # Fix formatting issues (replace non-breaking spaces)
        df = df.map(lambda x: x.replace('\xa0', ' ') if isinstance(x, str) else x)

        # Convert DataFrame to CSV
        output = BytesIO()
        df.to_csv(output, index=False, encoding='utf-8-sig')
        output.seek(0)

        # Upload to OneDrive
        upload_response = upload_to_onedrive(output.getvalue())

        return jsonify(upload_response)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
