from flask import Flask, render_template, request, redirect
import pandas as pd
from datetime import datetime
import os
from werkzeug.utils import secure_filename
from flask import send_file, session
from io import BytesIO


app = Flask(__name__)
app.secret_key = 'simple'  # Needed for session

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

            start = datetime.strptime(start_date, "%Y-%m-%d").date()
            end = datetime.strptime(end_date, "%Y-%m-%d").date()

            ins = df[df['Type IN or OUT only.'].str.upper() == 'IN'].copy()
            ins['DateStr'] = ins['Date:'].dt.date
            unique_ins = ins.drop_duplicates(subset=['Reg Number:', 'DateStr'])

            filtered = unique_ins[(unique_ins['DateStr'] >= start) & (unique_ins['DateStr'] <= end)]

            # Normalize Name casing (optional but recommended)
            filtered['Name:'] = filtered['Name:'].str.strip().str.title()

            # Now group by Reg Number only, and aggregate the rest
            grouped = filtered.groupby('Reg Number:').agg({
                'Name:': 'first',
                'Hostel:': 'first',
                'Room and Side:': 'first',
                'DateStr': 'count'
            }).reset_index().rename(columns={'DateStr': 'Days Present'})

            # Store in session
            session['summary_data'] = grouped.to_json()

            return render_template('results.html', tables=grouped.to_dict('records'))

        except Exception as e:
            return f"Error processing file: {e}"

    return render_template('index.html')

@app.route('/download')
def download_summary():
    if 'summary_data' not in session:
        return "No summary data found. Please generate summary first."

    # Load JSON back into DataFrame
    df = pd.read_json(session['summary_data'])

    # Create Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Summary')
    output.seek(0)

    return send_file(output,
                     download_name='summary.xlsx',
                     as_attachment=True,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')



import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

