"""
SOP & 二单元 Report Generator Skill
处理 SOP（SS首通/首单元）和二单元数据，按日期和销售名称组织，标注 24h 警戒
"""

import csv
from pathlib import Path
from collections import defaultdict
from datetime import datetime, timedelta
from typing import Tuple, Optional


class DataProcessor:
    """数据处理器 - 处理 CSV 并按规则组织"""

    def __init__(self, file_path: str, report_date: Optional[str] = None, output_base: Optional[str] = None):
        self.file_path = Path(file_path)
        if report_date:
            self.report_date = datetime.strptime(report_date, '%Y-%m-%d').replace(hour=23, minute=59, second=59)
        else:
            self.report_date = datetime.now().replace(hour=23, minute=59, second=59)
        self.report_day = self.report_date.date()
        
        # 输出基础路径（默认为同目录下的 exports/organized/SOP）
        if output_base:
            self.output_base = Path(output_base)
        else:
            self.output_base = Path('/Users/mac/Desktop/AL/skills/anthropics-skills/my_skills/exports/organized')
        
        # 检查文件
        if not self.file_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.file_path}")

    def _try_parse_date(self, s: str) -> Optional[datetime]:
        """解析日期字符串"""
        s = s.strip()
        if not s:
            return None
        
        # 尝试 dateutil
        try:
            from dateutil import parser
            return parser.parse(s)
        except Exception:
            pass
        
        # 常见格式
        fmts = [
            '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S',
            '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M',
            '%Y-%m-%d', '%Y/%m/%d',
            '%Y.%m.%d %H:%M:%S', '%Y.%m.%d'
        ]
        for f in fmts:
            try:
                return datetime.strptime(s, f)
            except Exception:
                continue
        
        # 时间戳
        try:
            ts = float(s)
            if ts > 1e12:
                ts = ts / 1000.0
            return datetime.fromtimestamp(ts)
        except Exception:
            return None

    def _detect_encoding(self) -> str:
        """检测文件编码"""
        encs = ['utf-8-sig', 'utf-8', 'gbk', 'gb18030', 'latin-1', 'cp1252']
        for e in encs:
            try:
                with open(self.file_path, newline='', encoding=e) as f:
                    _ = f.readline()
                return e
            except Exception:
                continue
        return 'utf-8'

    def _find_column_index(self, header: list, names: list) -> Optional[int]:
        """查找列索引"""
        for name in names:
            if name in header:
                return header.index(name)
        return None

    def process(self, mode: str) -> Tuple[int, int, Path]:
        """
        处理数据
        mode: 'sop' 或 'erdan'（二单元）
        返回: (total_rows, matched_rows, output_base_path)
        """
        encoding = self._detect_encoding()
        
        # 读取表头
        with open(self.file_path, newline='', encoding=encoding) as f:
            reader = csv.reader(f)
            header = next(reader)
            header = [h.strip() for h in header]

        # 查找列
        idx_sales = self._find_column_index(header, ['销售名称', '销售', 'sales'])
        idx_status = self._find_column_index(header, ['sop状态', '状态', 'status'])
        idx_type = self._find_column_index(header, ['sop类型', '类型', 'type'])
        idx_gen = self._find_column_index(header, ['sop生成日期', '生成日期'])

        if None in (idx_sales, idx_status, idx_type, idx_gen):
            raise ValueError(f"表头缺少必需列。找到的列: {header}")

        # 尝试找小组名称列（可选）
        idx_group = self._find_column_index(header, ['小组名称', '小组', 'group'])

        # 设置输出目录和过滤规则
        if mode == 'sop':
            base_out = self.output_base / 'SOP'
        else:
            base_out = self.output_base / '二单元'
        base_out.mkdir(parents=True, exist_ok=True)

        if mode == 'sop':
            include_types = {'SS首通', 'SS首单元'}
            exclude_keywords = {'CC'}
            category = 'SOP'
        else:  # erdan
            include_types = {'二单元'}
            exclude_keywords = set()
            category = '二单元'

        # 计算日期范围
        dates_to_include = {
            (self.report_day - timedelta(days=i)).isoformat() 
            for i in range(0, 3)
        }

        # 处理数据
        groups = defaultdict(list)
        total_rows = 0
        matched_rows = 0

        with open(self.file_path, newline='', encoding=encoding) as f:
            reader = csv.reader(f)
            _ = next(reader)  # skip header
            
            for row in reader:
                total_rows += 1
                if len(row) < len(header):
                    row = row + [''] * (len(header) - len(row))
                row = [c.strip() for c in row]

                # 类型过滤
                typ_raw = row[idx_type]
                typ = typ_raw.strip()
                
                # 排除 CC（SOP 模式）
                if exclude_keywords and any(kw in typ for kw in exclude_keywords):
                    continue
                
                # 检查是否匹配要求的类型
                if mode == 'sop':
                    # 严格匹配 SS首通 或 SS首单元
                    if typ not in include_types:
                        continue
                else:  # erdan
                    # 检查是否包含二单元
                    if '二单元' not in typ:
                        continue

                # 生成日期
                gen_raw = row[idx_gen]
                gen_dt = self._try_parse_date(gen_raw)
                if gen_dt is None:
                    continue
                gen_date_str = gen_dt.date().isoformat()

                # 检查日期是否在范围内
                if gen_date_str not in dates_to_include:
                    continue

                # 状态检查
                status = row[idx_status].replace(' ', '')
                is_unfinished = (
                    (status == '未完成') or 
                    ('未' in status and '完成' not in status) or 
                    ('未完成' in status)
                )
                if not is_unfinished:
                    continue

                # 计算年龄和警戒
                age_hours = (self.report_date - gen_dt).total_seconds() / 3600.0
                age_hours_str = f"{age_hours:.1f}"
                alert = 'YELLOW' if age_hours > 24 else ''

                # 分组（按日期、小组、销售）
                sales = row[idx_sales] or 'UNKNOWN'
                group_name = row[idx_group] or 'UNKNOWN' if idx_group is not None else 'UNKNOWN'
                out_row = row + [age_hours_str, alert, gen_date_str]
                groups[(category, self.report_day.isoformat(), group_name, sales)].append(out_row)
                matched_rows += 1

        # 写入文件
        extra_cols = ['age_hours', 'alert', 'gen_date']
        created = 0
        
        for (cat, rpt, group, sales), rows in groups.items():
            dir_path = base_out / rpt / group
            dir_path.mkdir(parents=True, exist_ok=True)
            fn = sales.strip()[:120].replace('/', '_').replace('\\', '_') or 'UNKNOWN'
            out_file = dir_path / f"{fn}.csv"
            
            with open(out_file, 'w', newline='', encoding='utf-8-sig') as wf:
                w = csv.writer(wf)
                w.writerow(header + extra_cols)
                w.writerows(rows)
            created += 1

        return total_rows, matched_rows, base_out

    def report(self, mode: str) -> str:
        """生成处理报告"""
        total, matched, out_path = self.process(mode)
        
        mode_name = "SOP(SS首通/首单元)" if mode == 'sop' else "二单元"
        lines = [
            f"📊 {mode_name} 数据处理报告",
            f"{'='*50}",
            f"报告日期: {self.report_day.isoformat()}",
            f"包含生成日期: {self.report_day.isoformat()} 及前两日",
            f"",
            f"处理结果:",
            f"  • 总行数: {total}",
            f"  • 匹配行数: {matched}",
            f"  • 生成 CSV 文件数: {len(list(out_path.rglob('*.csv')))}",
            f"  • 输出目录: {out_path}",
        ]
        
        # 列出子目录
        sop_dir = out_path / self.report_day.isoformat()
        if sop_dir.exists():
            groups = sorted([d.name for d in sop_dir.iterdir() if d.is_dir()])
            lines.append(f"\n📁 按小组分类（{self.report_day.isoformat()}）:")
            for g in groups[:15]:
                group_dir = sop_dir / g
                files = list(group_dir.glob('*.csv'))
                lines.append(f"  • 【{g}】- {len(files)} 个销售")
            if len(groups) > 15:
                lines.append(f"  ... 及其他 {len(groups) - 15} 个小组")
        
        return "\n".join(lines)


# ============ 命令处理接口 ============

def handle_sop_command(file_path: str, report_date: Optional[str] = None, output_base: Optional[str] = None) -> str:
    """处理 /sop 命令"""
    try:
        processor = DataProcessor(file_path, report_date, output_base)
        return processor.report('sop')
    except Exception as e:
        return f"❌ 处理 SOP 数据时出错: {e}"


def handle_erdan_command(file_path: str, report_date: Optional[str] = None, output_base: Optional[str] = None) -> str:
    """处理 /二单元 命令"""
    try:
        processor = DataProcessor(file_path, report_date, output_base)
        return processor.report('erdan')
    except Exception as e:
        return f"❌ 处理二单元数据时出错: {e}"


# ============ 使用示例 ============
if __name__ == "__main__":
    # 示例用法
    file_path = '/Users/mac/Downloads/1001035468_20250621_0000.csv'
    report_date = '2025-06-21'  # 使用实际数据的日期
    
    print("处理 SOP 数据...")
    print(handle_sop_command(file_path, report_date))
    print("\n" + "="*50 + "\n")
    print("处理二单元数据...")
    print(handle_erdan_command(file_path, report_date))
