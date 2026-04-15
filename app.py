from flask import Flask, render_template, jsonify, request, send_file
from datetime import datetime
import pymysql
from pymysql.cursors import DictCursor
from config import DB_CONFIG
from openpyxl import Workbook, load_workbook
import tempfile
import os
import shutil

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

@app.route('/api/report/preview', methods=['GET'])
def get_report_preview():
    """获取周报预览数据"""
    week_start = request.args.get('start_date')
    week_end = request.args.get('end_date')

    if not week_start or not week_end:
        return jsonify({'success': False, 'error': '缺少日期参数'}), 400

    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            # 本周计划（plan_start 在本周内的任务，不限状态）
            cursor.execute("""SELECT * FROM tasks
                           WHERE DATE(plan_start) <= %s
                           AND DATE(plan_end) >= %s
                           ORDER BY category, sort_order, id""",
                           (week_end, week_start))
            week_plan_tasks = cursor.fetchall()

            # 本周完成（plan_end 在本周内且状态为已完成）
            cursor.execute("""SELECT * FROM tasks
                           WHERE status = '已完成'
                           AND DATE(plan_end) BETWEEN %s AND %s
                           ORDER BY category, sort_order, id""",
                           (week_start, week_end))
            completed_tasks = cursor.fetchall()

        categories = ['教学', '竞赛', '就业', '科研', '项目', '职能组', '院校支撑']

        def categorize(task_list):
            result = {cat: [] for cat in categories}
            for task in task_list:
                cat = task.get('category')
                if cat in result:
                    result[cat].append({
                        'id': task.get('id'),
                        'name': task.get('name', ''),
                        'progress_note': task.get('progress_note', ''),
                        'status': task.get('status', '')
                    })
            return result

        # 转换日期
        def convert_dates(tasks):
            for task in tasks:
                for key in ['plan_start', 'plan_end', 'created_at', 'updated_at']:
                    if task.get(key) and isinstance(task[key], datetime):
                        task[key] = task[key].strftime('%Y-%m-%d')
            return tasks

        return jsonify({
            'success': True,
            'data': {
                'week_plan': categorize(convert_dates(week_plan_tasks)),
                'completed': categorize(convert_dates(completed_tasks))
            }
        })
    finally:
        conn.close()

@app.route('/api/report/generate', methods=['POST'])
def generate_report():
    """使用模板生成周报Excel"""
    data = request.get_json()
    week_start = data.get('start_date')
    week_end = data.get('end_date')

    # 编辑内容
    week_plan_content = data.get('week_plan_content', {})  # {category: content}
    next_week_plan = data.get('next_week_plan', '')  # 下周工作计划
    coordinated_items = data.get('coordinated_items', '')  # 待协调事项

    if not week_start or not week_end:
        return jsonify({'success': False, 'error': '缺少日期参数'}), 400

    # 查询周期内已完成任务
    conn = get_db_connection()
    try:
        with conn.cursor() as cursor:
            sql = """SELECT * FROM tasks
                     WHERE status = '已完成'
                     AND DATE(plan_end) BETWEEN %s AND %s
                     ORDER BY category, sort_order, id"""
            cursor.execute(sql, (week_start, week_end))
            tasks = cursor.fetchall()
    finally:
        conn.close()

    # 按分类归类任务
    categories = ['教学', '竞赛', '就业', '科研', '项目', '职能组', '院校支撑']
    categorized = {cat: [] for cat in categories}
    for task in tasks:
        cat = task.get('category')
        if cat in categorized:
            categorized[cat].append({
                'name': task.get('name', ''),
                'progress_note': task.get('progress_note', '')
            })

    # 加载模板文件
    template_path = os.path.join(os.path.dirname(__file__), '周报模版-新.xlsx')
    wb = load_workbook(template_path)
    ws = wb.active

    # 填充基本信息
    ws['B2'] = '王弯弯'
    ws['D2'] = f'{week_start} 至 {week_end}'
    ws['B3'] = '河工大'
    ws['D3'] = datetime.now().strftime('%Y-%m-%d')

    # 填充分类数据 (A5-A11 对应 教学-院校支撑)
    for i, cat in enumerate(categories, start=5):
        # 本周计划（B列）- 从编辑内容获取
        if week_plan_content and week_plan_content.get(cat):
            ws[f'B{i}'] = week_plan_content[cat]
        # 本周完成情况（C列）
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
        # 下周工作计划（D列）- 通用内容
        if next_week_plan:
            if i == 5:  # 只在第一行填入
                ws[f'D{i}'] = next_week_plan

    # 待协调事项（E5）
    if coordinated_items:
        ws['E5'] = coordinated_items

    # 保存到临时文件
    filename = f"{week_start}-{week_end}-周报.xlsx"
    temp_dir = tempfile.gettempdir()
    filepath = os.path.join(temp_dir, filename)
    wb.save(filepath)

    return send_file(filepath, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
