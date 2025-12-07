"""
Team Task Reporter - 团队任务管理与报表生成 Skill
自动化处理48小时SOP任务、7天有效期统计和每日报表生成
"""

import json
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import uuid

class TaskManager:
    """任务管理器 - 处理任务的添加、更新和查询"""
    
    def __init__(self, storage_file: str = "tasks.json"):
        self.storage_file = storage_file
        self.tasks = self._load_tasks()
    
    def _load_tasks(self) -> List[Dict]:
        """从存储文件加载任务"""
        try:
            with open(self.storage_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return []
    
    def _save_tasks(self):
        """保存任务到存储文件"""
        with open(self.storage_file, 'w', encoding='utf-8') as f:
            json.dump(self.tasks, f, ensure_ascii=False, indent=2)
    
    def add_task(self, task_name: str, assigned_to: str, deadline_hours: int = 48) -> Dict:
        """添加新任务到SOP清单"""
        task = {
            "id": str(uuid.uuid4())[:8],
            "name": task_name,
            "assigned_to": assigned_to,
            "created_at": datetime.now().isoformat(),
            "deadline_at": (datetime.now() + timedelta(hours=deadline_hours)).isoformat(),
            "status": "pending",
            "completed_at": None,
            "deadline_hours": deadline_hours
        }
        self.tasks.append(task)
        self._save_tasks()
        return task
    
    def update_status(self, task_id: str, status: str) -> Optional[Dict]:
        """更新任务状态"""
        for task in self.tasks:
            if task["id"] == task_id:
                task["status"] = status
                if status == "completed":
                    task["completed_at"] = datetime.now().isoformat()
                self._save_tasks()
                return task
        return None
    
    def get_active_tasks(self, deadline_hours: int = 48) -> List[Dict]:
        """获取活跃任务（在截止时间内）"""
        now = datetime.now()
        active = []
        
        for task in self.tasks:
            deadline = datetime.fromisoformat(task["deadline_at"])
            if deadline > now and task["status"] == "pending":
                active.append(task)
        
        return active
    
    def get_weekly_stats(self) -> Dict:
        """获取本周统计数据"""
        now = datetime.now()
        week_start = now - timedelta(days=now.weekday())
        week_end = week_start + timedelta(days=7)
        
        weekly_tasks = [
            t for t in self.tasks
            if datetime.fromisoformat(t["created_at"]) >= week_start
            and datetime.fromisoformat(t["created_at"]) <= week_end
        ]
        
        completed = sum(1 for t in weekly_tasks if t["status"] == "completed")
        total = len(weekly_tasks)
        
        return {
            "week_start": week_start.strftime("%Y-%m-%d"),
            "week_end": week_end.strftime("%Y-%m-%d"),
            "total_tasks": total,
            "completed_tasks": completed,
            "pending_tasks": total - completed,
            "completion_rate": f"{(completed/total*100):.1f}%" if total > 0 else "0%"
        }


class ReportGenerator:
    """报表生成器 - 生成任务报表和统计分析"""
    
    def __init__(self, task_manager: TaskManager):
        self.task_manager = task_manager
    
    def generate_daily_report(self) -> Dict:
        """生成每日报表"""
        now = datetime.now()
        active_tasks = self.task_manager.get_active_tasks()
        
        # 按分配成员分组
        tasks_by_member = {}
        for task in active_tasks:
            member = task["assigned_to"]
            if member not in tasks_by_member:
                tasks_by_member[member] = []
            tasks_by_member[member].append(task)
        
        # 计算每个成员的统计
        member_stats = {}
        for member, tasks in tasks_by_member.items():
            completed = sum(1 for t in tasks if t["status"] == "completed")
            member_stats[member] = {
                "assigned": len(tasks),
                "completed": completed,
                "pending": len(tasks) - completed,
                "completion_rate": f"{(completed/len(tasks)*100):.1f}%" if len(tasks) > 0 else "0%"
            }
        
        return {
            "report_date": now.strftime("%Y-%m-%d %H:%M:%S"),
            "report_type": "Daily SOP Report (48H)",
            "total_active_tasks": len(active_tasks),
            "tasks_by_member": tasks_by_member,
            "member_statistics": member_stats,
            "team_completion_rate": self._calculate_team_rate(active_tasks)
        }
    
    def generate_weekly_report(self) -> Dict:
        """生成周报"""
        weekly_stats = self.task_manager.get_weekly_stats()
        
        return {
            "report_date": datetime.now().strftime("%Y-%m-%d"),
            "report_type": "Weekly Report (7 Days)",
            "weekly_statistics": weekly_stats,
            "all_tasks": self.task_manager.tasks[-20:]  # 最近20条任务
        }
    
    def _calculate_team_rate(self, tasks: List[Dict]) -> str:
        """计算团队整体完成率"""
        if not tasks:
            return "0%"
        completed = sum(1 for t in tasks if t["status"] == "completed")
        return f"{(completed/len(tasks)*100):.1f}%"
    
    def format_report_for_email(self, report: Dict) -> str:
        """格式化报表为邮件文本"""
        if report["report_type"].startswith("Daily"):
            return self._format_daily_email(report)
        else:
            return self._format_weekly_email(report)
    
    def _format_daily_email(self, report: Dict) -> str:
        """格式化每日邮件"""
        lines = [
            f"📋 每日任务报表 - {report['report_date']}",
            "=" * 50,
            f"\n活跃任务总数: {report['total_active_tasks']}",
            f"团队完成率: {report['team_completion_rate']}\n",
            "📌 各成员任务统计:\n"
        ]
        
        for member, stats in report["member_statistics"].items():
            lines.append(f"  {member}:")
            lines.append(f"    ✅ 已完成: {stats['completed']}/{stats['assigned']}")
            lines.append(f"    ⏳ 待完成: {stats['pending']}")
            lines.append(f"    📊 完成率: {stats['completion_rate']}\n")
        
        lines.append("\n💡 详细任务列表:\n")
        for member, tasks in report["tasks_by_member"].items():
            lines.append(f"\n【{member}】")
            for task in tasks:
                status_emoji = "✅" if task["status"] == "completed" else "⏳"
                deadline = datetime.fromisoformat(task["deadline_at"]).strftime("%m-%d %H:%M")
                lines.append(f"  {status_emoji} {task['name']} (截止: {deadline})")
        
        return "\n".join(lines)
    
    def _format_weekly_email(self, report: Dict) -> str:
        """格式化周报邮件"""
        stats = report["weekly_statistics"]
        lines = [
            f"📊 周报 - {report['report_date']}",
            "=" * 50,
            f"\n周期: {stats['week_start']} 至 {stats['week_end']}",
            f"\n✅ 本周完成: {stats['completed_tasks']} 个任务",
            f"⏳ 本周待完成: {stats['pending_tasks']} 个任务",
            f"📈 总任务数: {stats['total_tasks']} 个",
            f"🎯 完成率: {stats['completion_rate']}\n"
        ]
        
        return "\n".join(lines)


class TeamTaskReporter:
    """主体 - 团队任务报告器，集成任务管理和报表生成"""
    
    def __init__(self):
        self.task_manager = TaskManager()
        self.report_generator = ReportGenerator(self.task_manager)
    
    def add_task(self, task_name: str, assigned_to: str, deadline_hours: int = 48) -> Dict:
        """添加任务"""
        return self.task_manager.add_task(task_name, assigned_to, deadline_hours)
    
    def update_task_status(self, task_id: str, status: str) -> Optional[Dict]:
        """更新任务状态"""
        return self.task_manager.update_status(task_id, status)
    
    def generate_report(self, report_type: str = "daily") -> Dict:
        """生成报表"""
        if report_type == "daily":
            return self.report_generator.generate_daily_report()
        else:
            return self.report_generator.generate_weekly_report()
    
    def get_report_email(self, report_type: str = "daily") -> str:
        """获取邮件格式的报表"""
        report = self.generate_report(report_type)
        return self.report_generator.format_report_for_email(report)
    
    def display_current_tasks(self) -> str:
        """显示当前所有活跃任务"""
        active_tasks = self.task_manager.get_active_tasks()
        if not active_tasks:
            return "暂无活跃任务"
        
        lines = ["📋 当前活跃任务列表:\n"]
        for task in active_tasks:
            deadline = datetime.fromisoformat(task["deadline_at"])
            hours_left = (deadline - datetime.now()).total_seconds() / 3600
            status_emoji = "✅" if task["status"] == "completed" else "⏳"
            lines.append(
                f"{status_emoji} [{task['id']}] {task['name']} "
                f"({task['assigned_to']}) - 剩余 {hours_left:.1f}h"
            )
        
        return "\n".join(lines)


# ============ 使用示例 ============
if __name__ == "__main__":
    reporter = TeamTaskReporter()
    
    # 添加示例任务
    print("添加示例任务...\n")
    task1 = reporter.add_task("完成API文档", "张三", 48)
    task2 = reporter.add_task("代码审查", "李四", 24)
    task3 = reporter.add_task("修复bug#123", "王五", 48)
    
    print(f"✅ 已添加任务: {task1['id']}, {task2['id']}, {task3['id']}\n")
    
    # 显示当前任务
    print(reporter.display_current_tasks())
    print("\n")
    
    # 更新任务状态
    print("更新任务状态...\n")
    reporter.update_task_status(task1['id'], "completed")
    
    # 生成每日报表
    print("生成每日报表...\n")
    daily_report = reporter.generate_report("daily")
    print(reporter.get_report_email("daily"))
    
    print("\n" + "="*50 + "\n")
    
    # 生成周报
    print("生成周报...\n")
    weekly_report = reporter.generate_report("weekly")
    print(reporter.get_report_email("weekly"))
