# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Flask + MySQL web application for weekly report management. Supports task management, daily task viewing, and weekly report export to Excel.

## Commands

```bash
# Initialize database
python init_db.py

# Run Flask application
python app.py
```

## Architecture

### Tech Stack
- **Backend:** Flask, PyMySQL, openpyxl
- **Database:** MySQL (database: `weekly_report`)
- **Frontend:** Pure HTML/JS with Jinja2 templates

### File Structure
```
├── app.py           # Flask application with all API routes
├── config.py        # Database configuration
├── init_db.py       # Database initialization (creates tasks table)
├── templates/
│   ├── base.html        # Base template
│   ├── tasks.html       # Task management page (/)
│   ├── daily.html       # Daily tasks page (/daily)
│   ├── report.html      # Weekly report export page (/report)
│   └── task_detail.html # Task detail page
```

### Database Schema
Tasks table (`tasks`):
- `id`, `name`, `category`, `priority`, `status`
- `plan_start`, `plan_end`, `progress_note`, `pause_reason`
- `sort_order`, `created_at`, `updated_at`

### API Endpoints
| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/tasks` | Query all tasks |
| POST | `/api/tasks` | Create task |
| GET | `/api/tasks/<id>` | Get single task |
| PUT | `/api/tasks/<id>` | Update task |
| DELETE | `/api/tasks/<id>` | Delete task |
| PUT | `/api/tasks/<id>/status` | Update status |
| PUT | `/api/tasks/<id>/sort` | Update sort order |
| GET | `/api/daily?date=YYYY-MM-DD` | Query tasks by date |
| POST | `/api/report/generate` | Export weekly Excel |

### Task Categories
`教学`, `竞赛`, `就业`, `科研`, `项目`, `职能组`, `院校支撑`

### Task Statuses
`未开始`, `进行中`, `已完成`, `暂停`

### Task Priorities
`高`, `中`, `低`

### Report Export Logic
The `/api/report/generate` endpoint expects JSON with `start_date` and `end_date` fields (not `week_start`/`week_end`). It queries tasks with `status='已完成'` and `plan_end` within the date range, then groups by category for Excel export.

### Key Implementation Notes
- Date fields in API responses are converted to `YYYY-MM-DD` string format
- `tasks.html` uses a category navigation bar; "all" view only shows high priority tasks
- Adding a task on tasks.html refreshes the full task list via `loadTasks()` to ensure correct filtering
