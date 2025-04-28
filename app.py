
from flask import Flask, render_template, request, send_file
import sqlite3
import pandas as pd
import os

app = Flask(__name__)

DB_NAME = 'database.db'

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS msel_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_number TEXT,
            unit TEXT,
            location TEXT,
            dtg TEXT
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    event_number = request.form['eventNumber']
    unit = request.form['unit']
    location = request.form['location']
    dtg = request.form['dtg']

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('INSERT INTO msel_entries (event_number, unit, location, dtg) VALUES (?, ?, ?, ?)',
              (event_number, unit, location, dtg))
    conn.commit()
    conn.close()

    return "Submission successful! <br><a href='/'>Go back</a>"

@app.route('/export')
def export():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query('SELECT * FROM msel_entries', conn)
    conn.close()

    file_path = 'msel_data.xlsx'
    df.to_excel(file_path, index=False)

    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    init_db()
    app.run(debug=True)
