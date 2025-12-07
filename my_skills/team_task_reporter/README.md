# Team Task Reporter Skill

## 概述

**Team Task Reporter** 是一个自动化的团队任务管理和报表生成 Skill，用于：

- 📋 **48小时 SOP 任务追踪** - 快速添加、分配和追踪短期任务
- 📊 **7天有效期统计** - 汇总本周的任务完成情况
- 📧 **每日报表生成** - 自动生成并格式化邮件报表
- 👥 **团队完成度分析** - 实时展示每个成员的任务进度和完成率

## 核心功能

### 1. 任务管理 (TaskManager)
```python
# 添加任务
task = reporter.add_task(
    task_name="完成API文档",
    assigned_to="张三",
    deadline_hours=48
)

# 更新任务状态
reporter.update_task_status(task_id="abc123", status="completed")

# 获取活跃任务
active_tasks = task_manager.get_active_tasks()
```

### 2. 报表生成 (ReportGenerator)
```python
# 生成每日报表（48小时SOP）
daily_report = reporter.generate_report("daily")

# 生成周报
weekly_report = reporter.generate_report("weekly")

# 获取邮件格式的报表
email_content = reporter.get_report_email("daily")
```

### 3. 显示当前任务
```python
# 列出所有活跃任务
print(reporter.display_current_tasks())
```

## 安装与使用

### 前置条件
- Python 3.8+

### 基本用法

```python
from team_task_reporter import TeamTaskReporter

# 初始化
reporter = TeamTaskReporter()

# 添加任务
task1 = reporter.add_task("完成文档", "张三", 48)
task2 = reporter.add_task("代码审查", "李四", 24)

# 更新状态
reporter.update_task_status(task1['id'], "completed")

# 生成报表
print(reporter.get_report_email("daily"))
```

### 输出示例

#### 每日报表
```
📋 每日任务报表 - 2025-12-07 10:30:00
==================================================

活跃任务总数: 3
团队完成率: 33.3%

📌 各成员任务统计:

  张三:
    ✅ 已完成: 1/2
    ⏳ 待完成: 1
    📊 完成率: 50.0%

  李四:
    ✅ 已完成: 0/1
    ⏳ 待完成: 1
    📊 完成率: 0.0%

💡 详细任务列表:

【张三】
  ✅ 完成API文档 (截止: 12-09 10:30)
  ⏳ 修复bug#123 (截止: 12-09 10:30)

【李四】
  ⏳ 代码审查 (截止: 12-08 10:30)
```

#### 周报
```
📊 周报 - 2025-12-07
==================================================

周期: 2025-12-01 至 2025-12-08

✅ 本周完成: 5 个任务
⏳ 本周待完成: 3 个任务
📈 总任务数: 8 个
🎯 完成率: 62.5%
```

## 数据存储

所有任务数据保存在 `tasks.json` 文件中，采用 JSON 格式存储：

```json
{
  "id": "a1b2c3d4",
  "name": "完成API文档",
  "assigned_to": "张三",
  "created_at": "2025-12-07T10:30:00",
  "deadline_at": "2025-12-09T10:30:00",
  "status": "pending",
  "completed_at": null,
  "deadline_hours": 48
}
```

## API 参考

### TeamTaskReporter

#### `add_task(task_name, assigned_to, deadline_hours=48)`
添加新任务到 SOP 清单

**参数：**
- `task_name` (str): 任务名称
- `assigned_to` (str): 分配给的团队成员
- `deadline_hours` (int): 截止时间（小时，默认48）

**返回：** 任务字典

#### `update_task_status(task_id, status)`
更新任务状态

**参数：**
- `task_id` (str): 任务ID
- `status` (str): 'completed' 或 'pending'

**返回：** 更新后的任务字典

#### `generate_report(report_type='daily')`
生成报表

**参数：**
- `report_type` (str): 'daily' 或 'weekly'

**返回：** 报表字典

#### `get_report_email(report_type='daily')`
获取邮件格式的报表

**参数：**
- `report_type` (str): 'daily' 或 'weekly'

**返回：** 格式化的报表字符串

#### `display_current_tasks()`
显示所有活跃任务

**返回：** 任务列表字符串

## 扩展功能建议

- [ ] 集成 Slack/邮件 API 自动发送报表
- [ ] 添加任务优先级和标签功能
- [ ] 支持任务中断和延期处理
- [ ] 集成日历视图和甘特图
- [ ] 添加成员工作负载均衡建议
- [ ] 支持 CSV/Excel 导出

## 许可证

MIT

## 作者

Generated for Team Task Management

---

**更新日期**: 2025-12-07
