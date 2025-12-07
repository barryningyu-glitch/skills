"""
二单元数据处理器 - 基于 Excel 数据
处理 120s 通话数，按学科和小组分类，标注 7 天内未完成的记录
输出为 Excel 格式，支持颜色标记
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from pathlib import Path
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Optional, Tuple, List
import csv


class ErDanProcessor:
    """二单元处理器 - 处理 Excel 数据，按 SS 和小组分组"""

    def __init__(self, file_path: str, report_date: Optional[str] = None, output_base: Optional[str] = None):
        self.file_path = Path(file_path)
        
        # 报告日期
        if report_date:
            self.report_date = datetime.strptime(report_date, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
        else:
            self.report_date = datetime.now().replace(hour=23, minute=59, second=59)
        self.report_day = self.report_date.date()
        
        # 输出基础路径
        if output_base:
            self.output_base = Path(output_base)
        else:
            self.output_base = Path('/Users/mac/Desktop/AL/skills/anthropics-skills/my_skills/exports/organized')
        
        # 检查文件
        if not self.file_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.file_path}")

    def _try_parse_date(self, s) -> Optional[datetime]:
        """解析日期"""
        if s is None or s == '':
            return None
        
        # 如果已经是 datetime
        if isinstance(s, datetime):
            return s
        
        s_str = str(s).strip()
        if not s_str:
            return None
        
        try:
            from dateutil import parser
            return parser.parse(s_str)
        except Exception:
            pass
        
        fmts = [
            '%Y-%m-%d %H:%M:%S.%f',
            '%Y-%m-%d %H:%M:%S',
            '%Y/%m/%d %H:%M:%S',
            '%Y-%m-%d %H:%M',
            '%Y/%m/%d %H:%M',
            '%Y-%m-%d',
            '%Y/%m/%d',
        ]
        for f in fmts:
            try:
                return datetime.strptime(s_str, f)
            except Exception:
                continue
        
        return None

    def process(self) -> Tuple[int, int, Path]:
        """
        处理 Excel 数据
        返回: (total_rows, matched_rows, output_base_path)
        """
        wb = openpyxl.load_workbook(self.file_path)
        ws = wb.active
        
        # 读取表头（第一行）
        header = [cell.value for cell in ws[1]]
        header = [str(h).strip() if h else '' for h in header]
        
        # 查找关键列
        idx_group = None
        idx_ss = None
        idx_student_id = None
        idx_complete_time = None
        idx_120s = None
        idx_20s = None
        
        for i, h in enumerate(header):
            if 'SS小组' in h or '小组' in h:
                idx_group = i
            elif 'SS' in h and idx_ss is None:  # 第一个 SS 列
                idx_ss = i
            elif '学员' in h or '学员ID' in h:
                idx_student_id = i
            elif '完课时间' in h or '完成时间' in h:
                idx_complete_time = i
            elif '120s' in h or '120秒' in h:
                idx_120s = i
            elif '20s' in h or '20秒' in h:
                idx_20s = i
        
        if None in (idx_group, idx_ss, idx_complete_time, idx_120s):
            raise ValueError(
                f"表头缺少必需列。找到的列:\n"
                f"  group={idx_group}, ss={idx_ss}, time={idx_complete_time}, 120s={idx_120s}\n"
                f"  表头: {header}"
            )
        
        # 计算7天时间范围
        report_day = self.report_date.date()
        seven_days_ago = report_day - timedelta(days=6)  # 包含今天算7天
        
        # 处理数据
        groups = defaultdict(list)
        total_rows = 0
        matched_rows = 0
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            total_rows += 1
            
            if len(row) <= max(idx_group, idx_ss, idx_complete_time, idx_120s):
                continue
            
            # 提取数据
            group_name = str(row[idx_group]).strip() if row[idx_group] else 'UNKNOWN'
            ss_name = str(row[idx_ss]).strip() if row[idx_ss] else 'UNKNOWN'
            complete_time_raw = row[idx_complete_time]
            call_120s = row[idx_120s]
            
            # 解析完成时间
            complete_dt = self._try_parse_date(complete_time_raw)
            if complete_dt is None:
                continue
            
            complete_date = complete_dt.date()
            
            # 检查是否在7天范围内
            if complete_date < seven_days_ago or complete_date > report_day:
                continue
            
            # 检查 120s 通话数是否为空/"-"/0（未完成）
            call_120s_str = str(call_120s).strip() if call_120s is not None else ''
            if call_120s_str == '' or call_120s_str == '-' or call_120s_str == '0':
                # 这是未完成的记录
                pass
            else:
                # 120s 通话数有值，说明已完成，跳过
                continue
            
            # 计算年龄（距报告日期的天数）
            age_days = (report_day - complete_date).days
            alert = 'RED' if age_days >= 3 else 'YELLOW' if age_days >= 1 else ''
            
            # 组装输出行（保留原始行 + 扩展列）
            out_row = list(row) + [age_days, alert, complete_date.isoformat()]
            
            # 分组（按小组和SS）
            groups[(group_name, ss_name)].append(out_row)
            matched_rows += 1
        
        # 写入文件到 output_base/二单元/YYYY-MM-DD/SS小组/SS.xlsx
        base_out = self.output_base / '二单元'
        date_dir = base_out / report_day.isoformat()
        
        created = 0
        for (group, ss), rows in groups.items():
            group_dir = date_dir / group
            group_dir.mkdir(parents=True, exist_ok=True)
            
            fn = ss.strip()[:120].replace('/', '_').replace('\\', '_') or 'UNKNOWN'
            out_file = group_dir / f"{fn}.xlsx"
            
            # 用 openpyxl 写 Excel
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = fn[:31]
            
            # 写表头
            ws.append(header + ['age_days', 'alert', 'complete_date'])
            
            # 定义填充颜色
            red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            white_font = Font(color='FFFFFF')
            black_font = Font(color='000000')
            
            # 边框样式
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # 写数据行，根据 alert 标记颜色
            for row in rows:
                ws.append(row)
                row_num = ws.max_row
                alert = row[-2]  # alert 列是倒数第二列
                
                # 根据警戒级别应用颜色
                fill = None
                font = black_font
                if alert == 'RED':
                    fill = red_fill
                    font = white_font
                elif alert == 'YELLOW':
                    fill = yellow_fill
                    font = black_font
                
                # 应用样式到整行
                for col_num, cell in enumerate(ws[row_num], 1):
                    if fill:
                        cell.fill = fill
                    cell.font = font
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            # 应用表头样式
            header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            header_font = Font(color='FFFFFF', bold=True)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            # 调整列宽
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_len:
                            max_len = len(str(cell.value))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
            
            wb.save(out_file)
            created += 1
        
        return total_rows, matched_rows, base_out

    def report(self) -> str:
        """生成处理报告"""
        total, matched, out_path = self.process()
        
        seven_days_ago = (self.report_date - timedelta(days=6)).date()
        
        lines = [
            f"📊 二单元数据处理报告",
            f"{'='*50}",
            f"报告日期: {self.report_day.isoformat()}",
            f"时间范围: {seven_days_ago.isoformat()} ~ {self.report_day.isoformat()}（7天）",
            f"筛选条件: 120s通话数 为空/'-'/0（未完成）",
            f"",
            f"处理结果:",
            f"  • 总行数: {total}",
            f"  • 匹配行数（7天未完成）: {matched}",
            f"  • 生成 CSV 文件数: {len(list(out_path.rglob('*.csv')))}",
            f"  • 输出目录: {out_path}",
        ]
        
        # 列出分组统计
        date_dir = out_path / self.report_day.isoformat()
        if date_dir.exists():
            groups = sorted([d.name for d in date_dir.iterdir() if d.is_dir()])
            lines.append(f"\n📁 按SS小组分类（{self.report_day.isoformat()}）:")
            for g in groups[:20]:
                group_dir = date_dir / g
                files = list(group_dir.glob('*.csv'))
                lines.append(f"  • 【{g}】- {len(files)} 个销售")
            if len(groups) > 20:
                lines.append(f"  ... 及其他 {len(groups) - 20} 个小组")
        
        return "\n".join(lines)


# ============ 命令处理接口 ============

def handle_erdan_command(file_path: str, report_date: Optional[str] = None, output_base: Optional[str] = None) -> str:
    """处理 /二单元 命令"""
    try:
        processor = ErDanProcessor(file_path, report_date, output_base)
        return processor.report()
    except Exception as e:
        import traceback
        return f"❌ 处理二单元数据时出错:\n{traceback.format_exc()}"


# ============ 使用示例 ============
if __name__ == "__main__":
    file_path = '/Users/mac/Downloads/国内SS业务数据仪表盘-SH_二单元覆盖率数据副本_20251207_1611.xlsx'
    output_base = '/Users/mac/Desktop/AL/skills/anthropics-skills/my_skills/exports/organized'
    
    print("处理二单元数据...")
    result = handle_erdan_command(file_path, output_base=output_base)
    print(result)
