from flask import Flask, request, jsonify
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import numbers
from openpyxl.drawing.image import Image

app = Flask(__name__)

# Your main endpoint for processing the Excel file
@app.route('/process_excel', methods=['POST'])
def process_excel():
    try:
        # Expecting the file name from the POST request
        data = request.json
        file_name = data.get("file_name", "Book.xlsx")  # Default is "Book.xlsx"

        # Your existing script starts here
        df = pd.read_excel(file_name)

        # Example processing (can replace with full logic)
        names = list(filter(lambda x: len(str(x).split()) > 1, df.groupby("Source").count().reset_index()["Source"]))
        name_indices = []
        for name in names:
            name_indices.append(np.nonzero(np.array(df["Source"]) == name)[0][0])
        name_indices = np.array(sorted(name_indices) + [len(df)])

        # Simulated response for testing
        return jsonify({
            "message": "Script executed successfully!",
            "names": names,
            "indices": name_indices.tolist()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(port=5000)  # The app will run on http://localhost:5000
