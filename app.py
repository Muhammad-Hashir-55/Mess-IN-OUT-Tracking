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
            df['Date:'] = pd.to_datetime(df['Date:'], errors='coerce')

            start = datetime.strptime(start_date, "%Y-%m-%d")
            end = datetime.strptime(end_date, "%Y-%m-%d")

            ins = df[df['Type IN or OUT only.'] == 'IN']
            unique_ins = ins.drop_duplicates(subset=['Reg Number:', 'Date:'])
            filtered = unique_ins[(unique_ins['Date:'] >= start) & (unique_ins['Date:'] <= end)]
            grouped = filtered.groupby(['Reg Number:', 'Name:','Hostel:' , 'Room and Side:']).size().reset_index(name='Days Present')

            return render_template('results.html', tables=grouped.to_dict('records'))
        except Exception as e:
            return f"Error processing file: {e}"

    return render_template('index.html')


import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

