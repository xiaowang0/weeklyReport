from flask import Flask, render_template, jsonify, request, send_file
from datetime import datetime
import pymysql
from config import DB_CONFIG
from openpyxl import Workbook

app = Flask(__name__)

def get_db_connection():
    return pymysql.connect(**DB_CONFIG)

@app.route('/')
def index():
    return render_template('tasks.html')

@app.route('/tasks')
def tasks_page():
    return render_template('tasks.html')

@app.route('/daily')
def daily_page():
    return render_template('daily.html')

@app.route('/report')
def report_page():
    return render_template('report.html')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
