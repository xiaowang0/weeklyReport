from flask import Flask, render_template, jsonify, request, send_file
from datetime import datetime
import pymysql
from pymysql.cursors import DictCursor
from config import DB_CONFIG
from openpyxl import Workbook
import tempfile
import os

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

@app.route('/task/<int:task_id>')
def task_detail_page(task_id):
    return render_template('task_detail.html', task_id=task_id)

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
    required_fields = ['name', 'category', 'plan_start', 'plan_end']
    for field in required_fields:
        if not data.get(field):
            return jsonify({'success': False, 'error': '缺少必填字段'}), 400

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

@app.route('/api/tasks/<int:task_id>', methods=['GET'])
def get_task(task_id):
    """获取单个任务"""
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM tasks WHERE id = %s", (task_id,))
            task = cursor.fetchone()
            if not task:
                return jsonify({'success': False, 'error': '任务不存在'}), 404
            for key in ['plan_start', 'plan_end', 'created_at', 'updated_at']:
                if task.get(key) and isinstance(task[key], datetime):
                    task[key] = task[key].strftime('%Y-%m-%d')
            return jsonify({'success': True, 'data': task})
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

@app.route('/api/tasks/<int:task_id>/sort', methods=['PUT'])
def update_task_sort(task_id):
    """更新任务的 sort_order 字段，用于同类内调整顺序"""
    data = request.get_json()
    new_order = data.get('sort_order', 0)

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            cursor.execute("UPDATE tasks SET sort_order = %s, updated_at = NOW() WHERE id = %s", (new_order, task_id))
            conn.commit()
            if cursor.rowcount == 0:
                return jsonify({'success': False, 'error': '任务不存在'}), 404
            return jsonify({'success': True, 'message': '排序更新成功'})
    finally:
        conn.close()

@app.route('/api/daily', methods=['GET'])
def get_daily_tasks():
    """按日期查询任务，不传date则默认当天"""
    date = request.args.get('date')

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            if date:
                sql = """SELECT * FROM tasks
                         WHERE DATE(plan_start) = %s
                         ORDER BY category, sort_order, id"""
                cursor.execute(sql, (date,))
            else:
                today = datetime.now().strftime('%Y-%m-%d')
                sql = """SELECT * FROM tasks
                         WHERE DATE(plan_start) = %s
                         ORDER BY category, sort_order, id"""
                cursor.execute(sql, (today,))

            tasks = cursor.fetchall()
            for task in tasks:
                for key in ['plan_start', 'plan_end', 'created_at', 'updated_at']:
                    if task.get(key) and isinstance(task[key], datetime):
                        task[key] = task[key].strftime('%Y-%m-%d')
            return jsonify({'success': True, 'data': tasks})
    finally:
        conn.close()

@app.route('/api/report/generate', methods=['POST'])
def generate_report():
    """生成周报Excel"""
    data = request.get_json()
    week_start = data.get('week_start')
    week_end = data.get('week_end')

    if not week_start or not week_end:
        return jsonify({'success': False, 'error': '缺少 week_start 或 week_end'}), 400

    # 查询周期内已完成任务
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            sql = """SELECT * FROM tasks
                     WHERE status = '已完成'
                     AND plan_end BETWEEN %s AND %s
                     ORDER BY category, sort_order, id"""
            cursor.execute(sql, (week_start, week_end))
            tasks = cursor.fetchall()
    finally:
        conn.close()

    # 按六大分类归类任务（存储字典列表）
    categories = ['教学', '竞赛', '就业', '科研', '项目', '职能组', '院校支撑']
    categorized = {cat: [] for cat in categories}
    for task in tasks:
        cat = task.get('category')
        if cat in categorized:
            categorized[cat].append({
                'name': task.get('name', ''),
                'progress_note': task.get('progress_note', '')
            })

    # 生成Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "工作周报"

    # 设置列宽
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 25
    ws.column_dimensions['E'].width = 25

    # A1: 工作周报
    ws['A1'] = '工作周报'
    ws.merge_cells('A1:E1')

    # A2: 汇报人    B2: 王弯弯    C2: 汇报时间段    D2: {week_start} 至 {week_end}
    ws['A2'] = '汇报人'
    ws['B2'] = '王弯弯'
    ws['C2'] = '汇报时间段'
    ws['D2'] = f'{week_start} 至 {week_end}'

    # A3: 所属项目  B3: 河工大    C3: 提交时间      D3: {提交时间}
    ws['A3'] = '所属项目'
    ws['B3'] = '河工大'
    ws['C3'] = '提交时间'
    ws['D3'] = datetime.now().strftime('%Y-%m-%d')

    # A4: 工作分类  B4: 本周计划  C4: 本周完成情况  D4: 下周工作计划    E4: 待协调、支撑事项
    ws['A4'] = '工作分类'
    ws['B4'] = '本周计划'
    ws['C4'] = '本周完成情况'
    ws['D4'] = '下周工作计划'
    ws['E4'] = '待协调、支撑事项'

    # A5-A11: 六大分类及本周完成情况
    for i, cat in enumerate(categories, start=5):
        ws[f'A{i}'] = cat
        # 填入本周完成情况（C列）
        tasks_in_cat = categorized.get(cat, [])
        if tasks_in_cat:
            task_texts = []
            for t in tasks_in_cat:
                name = t['name']
                note = t.get('progress_note', '')
                if note:
                    task_texts.append(f"{name}：{note}")
                else:
                    task_texts.append(name)
            ws[f'C{i}'] = '\n'.join(task_texts)
        # B/D/E列保留为空（手动填写区域）

    # 保存到临时文件
    filename = f"{week_start}-{week_end}-周报.xlsx"
    temp_dir = tempfile.gettempdir()
    filepath = os.path.join(temp_dir, filename)
    wb.save(filepath)

    return send_file(filepath, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
