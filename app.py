from flask import Flask, render_template, jsonify, request, send_file
from datetime import datetime
import pymysql
from pymysql.cursors import DictCursor
from config import DB_CONFIG
from openpyxl import Workbook

app = Flask(__name__)

def get_db_connection():
    conn = pymysql.connect(**DB_CONFIG)
    conn.cursorclass = DictCursor
    return conn

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

# Task CRUD API
@app.route('/api/tasks', methods=['GET'])
def get_tasks():
    """查询所有任务，按 category, sort_order, id 排序"""
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM tasks ORDER BY category, sort_order, id")
            tasks = cursor.fetchall()
            # 转换日期为字符串
            for task in tasks:
                for key in ['plan_start', 'plan_end', 'created_at', 'updated_at']:
                    if task.get(key) and isinstance(task[key], datetime):
                        task[key] = task[key].strftime('%Y-%m-%d')
            return jsonify({'success': True, 'data': tasks})
    finally:
        conn.close()

@app.route('/api/tasks', methods=['POST'])
def create_task():
    """创建新任务"""
    data = request.get_json()
    required_fields = ['name', 'category']
    for field in required_fields:
        if not data.get(field):
            return jsonify({'success': False, 'error': f'缺少必填字段: {field}'}), 400

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            sql = """INSERT INTO tasks
                     (name, category, priority, status, plan_start, plan_end, sort_order, created_at, updated_at)
                     VALUES (%s, %s, %s, %s, %s, %s, %s, NOW(), NOW())"""
            cursor.execute(sql, (
                data['name'],
                data['category'],
                data.get('priority', '中'),
                data.get('status', '未开始'),
                data.get('plan_start'),
                data.get('plan_end'),
                data.get('sort_order', 0)
            ))
            conn.commit()
            task_id = cursor.lastrowid
            cursor.execute("SELECT * FROM tasks WHERE id = %s", (task_id,))
            task = cursor.fetchone()
            for key in ['plan_start', 'plan_end', 'created_at', 'updated_at']:
                if task.get(key) and isinstance(task[key], datetime):
                    task[key] = task[key].strftime('%Y-%m-%d')
            return jsonify({'success': True, 'data': task}), 201
    finally:
        conn.close()

@app.route('/api/tasks/<int:task_id>', methods=['PUT'])
def update_task(task_id):
    """更新任务所有字段"""
    data = request.get_json()
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            sql = """UPDATE tasks SET
                     name = %s, category = %s, priority = %s, status = %s,
                     plan_start = %s, plan_end = %s, progress_note = %s,
                     pause_reason = %s, sort_order = %s, updated_at = NOW()
                     WHERE id = %s"""
            cursor.execute(sql, (
                data.get('name'),
                data.get('category'),
                data.get('priority'),
                data.get('status'),
                data.get('plan_start'),
                data.get('plan_end'),
                data.get('progress_note'),
                data.get('pause_reason'),
                data.get('sort_order', 0),
                task_id
            ))
            conn.commit()
            if cursor.rowcount == 0:
                return jsonify({'success': False, 'error': '任务不存在'}), 404
            cursor.execute("SELECT * FROM tasks WHERE id = %s", (task_id,))
            task = cursor.fetchone()
            for key in ['plan_start', 'plan_end', 'created_at', 'updated_at']:
                if task.get(key) and isinstance(task[key], datetime):
                    task[key] = task[key].strftime('%Y-%m-%d')
            return jsonify({'success': True, 'data': task})
    finally:
        conn.close()

@app.route('/api/tasks/<int:task_id>', methods=['DELETE'])
def delete_task(task_id):
    """删除任务"""
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("DELETE FROM tasks WHERE id = %s", (task_id,))
            conn.commit()
            if cursor.rowcount == 0:
                return jsonify({'success': False, 'error': '任务不存在'}), 404
            return jsonify({'success': True, 'message': '删除成功'})
    finally:
        conn.close()

@app.route('/api/tasks/<int:task_id>/status', methods=['PUT'])
def update_task_status(task_id):
    """更新任务状态，支持 progress_note 和 pause_reason"""
    data = request.get_json()
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            sql = """UPDATE tasks SET
                     status = %s, progress_note = %s, pause_reason = %s, updated_at = NOW()
                     WHERE id = %s"""
            cursor.execute(sql, (
                data.get('status'),
                data.get('progress_note'),
                data.get('pause_reason'),
                task_id
            ))
            conn.commit()
            if cursor.rowcount == 0:
                return jsonify({'success': False, 'error': '任务不存在'}), 404
            cursor.execute("SELECT * FROM tasks WHERE id = %s", (task_id,))
            task = cursor.fetchone()
            for key in ['plan_start', 'plan_end', 'created_at', 'updated_at']:
                if task.get(key) and isinstance(task[key], datetime):
                    task[key] = task[key].strftime('%Y-%m-%d')
            return jsonify({'success': True, 'data': task})
    finally:
        conn.close()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
