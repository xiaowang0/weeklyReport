# 工作周报系统

Flask + JSON 文件存储的任务管理和周报导出系统。

## 功能特点

- **任务管理**：创建、编辑、删除任务，支持分类和优先级
- **任务预览**：按教学周查看计划任务
- **每日记录**：按完成日期查看已完成任务
- **今日提醒**：本地存储的每日提醒事项
- **周报导出**：生成 Excel 格式周报

## 技术栈

- **后端**：Flask, PyMySQL, openpyxl
- **前端**：Vue 3, Pure HTML/CSS
- **存储**：JSON 文件（tasks.json）

## 快速开始

### 1. 安装依赖

```bash
pip install flask pymysql openpyxl
```

### 2. 运行应用

```bash
python app.py
```

访问 http://localhost:5000

## 目录结构

```
├── app.py              # Flask 应用主文件
├── config.py           # 数据库配置
├── tasks.json          # 任务数据存储
├── templates/
│   ├── index.html      # SPA 主页面（任务管理、每日记录、今日提醒、周报导出）
│   ├── record.html     # 每日记录独立页面
│   ├── daily.html      # 每日任务独立页面
│   ├── report.html     # 周报导出独立页面
│   └── base.html       # 基础模板
└── 周报模版-新.xlsx    # 周报导出模板
```

## 任务分类

- 教学
- 竞赛
- 就业
- 科研
- 项目
- 职能组
- 院校支撑

## 任务状态

- 未开始
- 进行中
- 已完成
- 暂停

## 任务优先级

- 高
- 中
- 低

## 教学周计算

系统采用教学周计算方式：
- **第1周起始日期**：2026-03-02（可按需修改）
- 周选择器选择 ISO 周后自动转换为教学周显示

## API 接口

### 任务管理

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/tasks` | 获取所有任务 |
| POST | `/api/tasks` | 创建任务 |
| GET | `/api/tasks/<id>` | 获取单个任务 |
| PUT | `/api/tasks/<id>` | 更新任务 |
| DELETE | `/api/tasks/<id>` | 删除任务 |
| PUT | `/api/tasks/<id>/status` | 更新任务状态 |
| PUT | `/api/tasks/<id>/sort` | 更新排序 |

### 任务预览

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/daily?date=YYYY-MM-DD` | 按日期查询任务 |

### 周报

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/api/report/preview?start_date=&end_date=` | 获取周报预览 |
| POST | `/api/report/generate` | 导出周报 Excel |

## 页面说明

### 1. 任务管理（首页）

- 左侧导航栏切换分类
- 顶部卡片显示紧急任务（高优先级）
- 点击复选框标记任务完成
- 完成时自动记录完成日期

### 2. 任务预览

- 按教学周选择，查看该周计划内的所有任务
- 支持按状态筛选

### 3. 每日记录

- 按教学周选择
- 显示该周每天的已完成任务（按完成日期）
- 星期一~星期日 分栏展示

### 4. 今日提醒

- 本地存储，不同步到服务器
- 按日期自动分组
- 刷新页面数据保留

### 5. 周报导出

- 选择周日期范围
- 预览本周计划和已完成任务
- 可编辑内容后导出 Excel

## 数据字段

### 任务字段

```json
{
  "id": 1,
  "name": "任务名称",
  "category": "教学",
  "priority": "高",
  "status": "已完成",
  "plan_start": "2026-04-13",
  "plan_end": "2026-04-19",
  "progress_note": "进度说明",
  "pause_reason": "暂停原因",
  "completed_at": "2026-04-16",
  "sort_order": 0,
  "created_at": "2026-04-01",
  "updated_at": "2026-04-16"
}
```

## 注意事项

1. **数据文件**：`tasks.json` 是任务数据存储文件，请定期备份
2. **时区问题**：系统使用本地时区计算日期
3. **教学周**：第1周起始日期在代码中硬编码（TERM_START），如需修改请在 index.html、daily.html、record.html 中查找并修改
4. **今日提醒**：数据存储在浏览器 localStorage 中，不同浏览器/设备数据不共享
