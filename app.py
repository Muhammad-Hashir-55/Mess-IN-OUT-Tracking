from flask import Flask, render_template, request, redirect
import pandas as pd
from datetime import datetime
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['sheet']
        if not file:
            return "No file uploaded"

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        start_date = request.form.get('start')
        end_date = request.form.get('end')

        try:
            df = pd.read_excel(filepath)

            # Ensure date is datetime
            df['Date:'] = pd.to_datetime(df['Date:'], errors='coerce')

            # Filter IN only (case insensitive)
            ins = df[df['Type IN or OUT only.'].str.upper() == 'IN']

            # Drop duplicate INs per reg per day
            ins['DateStr'] = ins['Date:'].dt.date.astype(str)
            unique_ins = ins.drop_duplicates(subset=['Reg Number:', 'DateStr'])

            # Filter by selected date range
            start = datetime.strptime(start_date, "%Y-%m-%d").date()
            end = datetime.strptime(end_date, "%Y-%m-%d").date()
            filtered = unique_ins[(unique_ins['Date:'].dt.date >= start) & (unique_ins['Date:'].dt.date <= end)]

            # Group by Reg, Name, Hostel, Room and Side
            grouped = filtered.groupby(['Reg Number:', 'Name:', 'Hostel:', 'Room and Side:']).size().reset_index(name='Days Present')

            return render_template('results.html', tables=grouped.to_dict('records'))
        except Exception as e:
            return f"Error processing file: {e}"

    return render_template('index.html')

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
