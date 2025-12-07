"""
Service Report Generator - 生成当月服务绩效汇总报表
按小组/销售人员统计首通、首单元、二单元的完成率
"""

import csv
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from pathlib import Path
from datetime import datetime, timedelta
from collections import defaultdict
import os


class ServiceReportGenerator:
    """服务绩效汇总报表生成器"""
    
    def __init__(self, report_date=None):
        """
        初始化报表生成器
        
        Args:
            report_date: 报表日期 (datetime.date 对象)，不指定则默认当天
        """
        if report_date is None:
            report_date = datetime.now().date()
        self.report_date = report_date
        self.report_month = report_date.replace(day=1)  # 当月 1 号
        
    def read_sop_data(self, sop_file):
        """读取 SOP CSV 数据 - 只统计上海 SS 团队，分别统计首通和首单元"""
        sop_data = defaultdict(lambda: {
            "首通": {"总数": 0, "完成": 0, "销售": {}},
            "首单元": {"总数": 0, "完成": 0, "销售": {}}
        })
        
        try:
            with open(sop_file, 'r', encoding='gb18030') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # 清理列名空格
                    row = {k.strip(): v for k, v in row.items()}
                    
                    # 过滤：只统计上海 SS 团队（小组名称包含 "SH-SS"）
                    小组 = row.get('小组名称', '').strip()
                    if not 小组.startswith('SH-SS'):
                        continue
                    
                    # 解析生成日期
                    try:
                        sop_date = datetime.strptime(row.get('sop生成日期', '')[:10], '%Y-%m-%d').date()
                    except (ValueError, TypeError):
                        continue
                    
                    # 只统计当月数据（分母）
                    if sop_date.month != self.report_month.month or sop_date.year != self.report_month.year:
                        continue
                    
                    销售 = row.get('销售名称', '未知销售').strip()
                    sop_类型 = row.get('sop类型', '').strip()
                    状态 = row.get('sop状态', '').strip()
                    
                    # 判断是首通还是首单元
                    if '首课' in sop_类型 or '首通' in sop_类型:
                        type_key = '首通'
                    elif '首单元' in sop_类型:
                        type_key = '首单元'
                    else:
                        continue
                    
                    # 初始化销售数据
                    if 销售 not in sop_data[小组][type_key]['销售']:
                        sop_data[小组][type_key]['销售'][销售] = {"总数": 0, "完成": 0}
                    
                    # 统计当月总数（分母）
                    sop_data[小组][type_key]['总数'] += 1
                    sop_data[小组][type_key]['销售'][销售]['总数'] += 1
                    
                    # 统计完成数（状态为"已完成"，不受日期限制，包括历史完成）
                    if 状态 == '已完成':
                        sop_data[小组][type_key]['完成'] += 1
                        sop_data[小组][type_key]['销售'][销售]['完成'] += 1
            
            return dict(sop_data)
        except Exception as e:
            print(f"错误：无法读取 SOP 文件 {sop_file}: {e}")
            return {}
    
    def read_erdan_data(self, erdan_file):
        """读取二单元 Excel 数据 - 只统计上海 SS 团队"""
        erdan_data = defaultdict(lambda: {"总数": 0, "完成": 0, "销售": {}})
        
        try:
            wb = openpyxl.load_workbook(erdan_file)
            ws = wb.active
            
            for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                # 列：大区, SS小组, SS, 学员ID, 二单元第八节课完课时间, 20s通话数, 120s通话数
                if not row[0]:  # 跳过空行
                    continue
                
                小组 = row[1] if len(row) > 1 else '未知小组'
                
                # 过滤：只统计上海 SS 团队（小组名称包含 "SH-SS"）
                if not (小组 and str(小组).startswith('SH-SS')):
                    continue
                
                销售 = row[2] if len(row) > 2 else '未知销售'
                完课时间_str = row[4] if len(row) > 4 else None
                call_120s = row[6] if len(row) > 6 else None
                
                # 解析完课时间
                try:
                    if isinstance(完课时间_str, str):
                        完课时间 = datetime.strptime(完课时间_str[:10], '%Y-%m-%d').date()
                    elif isinstance(完课时间_str, datetime):
                        完课时间 = 完课时间_str.date()
                    else:
                        完课时间 = None
                except (ValueError, TypeError, AttributeError):
                    完课时间 = None
                
                # 只统计当月完课的数据（分母）
                if 完课时间 is None or 完课时间.month != self.report_month.month or 完课时间.year != self.report_month.year:
                    continue
                
                # 初始化销售数据
                if 销售 not in erdan_data[小组]['销售']:
                    erdan_data[小组]['销售'][销售] = {"总数": 0, "完成": 0}
                
                # 统计当月总数（分母）
                erdan_data[小组]['总数'] += 1
                erdan_data[小组]['销售'][销售]['总数'] += 1
                
                # 统计完成数（120s通话数有数字）
                try:
                    if call_120s and str(call_120s).strip() not in ['', '−', '0']:
                        erdan_data[小组]['完成'] += 1
                        erdan_data[小组]['销售'][销售]['完成'] += 1
                except (ValueError, TypeError):
                    pass
            
            return dict(erdan_data)
        except Exception as e:
            print(f"错误：无法读取二单元文件 {erdan_file}: {e}")
            return {}
    
    def _apply_rate_color(self, cell, rate, green_fill, yellow_fill, red_fill, green_font, yellow_font, red_font):
        """根据完成率给单元格着色"""
        if rate >= 80:
            cell.fill = green_fill
            cell.font = green_font
        elif rate >= 60:
            cell.fill = yellow_fill
            cell.font = yellow_font
        else:
            cell.fill = red_fill
            cell.font = red_font
    
    def generate_report(self, sop_file, erdan_file, output_base=None):
        """生成汇总报表"""
        if output_base is None:
            output_base = str(Path.home() / 'Desktop' / 'AL' / 'skills' / 'exports' / 'organized')
        
        # 创建输出目录
        output_dir = Path(output_base) / '服务报表' / self.report_date.isoformat()
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # 读取数据
        sop_data = self.read_sop_data(sop_file)
        erdan_data = self.read_erdan_data(erdan_file)
        
        # 合并所有小组
        all_groups = set(sop_data.keys()) | set(erdan_data.keys())
        all_groups = sorted(all_groups)
        
        # 创建 Excel 工作簿
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"服务报表_{self.report_date.strftime('%Y%m%d')}"
        
        # 定义样式
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        group_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        group_font = Font(bold=True, size=10)
        green_fill = PatternFill(start_color='70AD47', end_color='70AD47', fill_type='solid')
        green_font = Font(color='FFFFFF', bold=True)
        red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        red_font = Font(color='FFFFFF', bold=True)
        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        yellow_font = Font(color='000000', bold=True)
        
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 写入标题
        ws.merge_cells('A1:L1')
        title = ws['A1']
        title.value = f"上海SS团队服务绩效汇总报表 - {self.report_date.strftime('%Y年%m月%d日')}"
        title.font = Font(bold=True, size=14)
        title.alignment = center_align
        
        # 写入日期范围说明
        ws.merge_cells('A2:L2')
        date_range = ws['A2']
        month_start = self.report_month.strftime('%Y-%m-01')
        month_end = self.report_date.strftime('%Y-%m-%d')
        date_range.value = f"统计周期：{month_start} 至 {month_end}"
        date_range.font = Font(italic=True, size=10)
        date_range.alignment = center_align
        
        # 写入列标题
        headers = ['小组', 'SS销售', '首通总数', '首通完成', '首通率', '首单元总数', '首单元完成', '首单元率', '二单元总数', '二单元完成', '二单元率', '综合评分']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
        
        # 设置列宽
        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 16
        for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws.column_dimensions[col].width = 12
        
        # 写入数据
        current_row = 5
        
        total_stt_count = 0  # 首通总数
        total_stt_completed = 0  # 首通完成数
        total_sdy_count = 0  # 首单元总数
        total_sdy_completed = 0  # 首单元完成数
        total_erdan_count = 0  # 二单元总数
        total_erdan_completed = 0  # 二单元完成数
        
        for group in all_groups:
            group_sop = sop_data.get(group, {})
            group_erdan = erdan_data.get(group, {})
            
            # 首通数据
            stt_total = group_sop.get('首通', {}).get('总数', 0)
            stt_completed = group_sop.get('首通', {}).get('完成', 0)
            stt_rate = (stt_completed / stt_total * 100) if stt_total > 0 else 0
            
            # 首单元数据
            sdy_total = group_sop.get('首单元', {}).get('总数', 0)
            sdy_completed = group_sop.get('首单元', {}).get('完成', 0)
            sdy_rate = (sdy_completed / sdy_total * 100) if sdy_total > 0 else 0
            
            # 二单元数据
            erdan_total = group_erdan.get('总数', 0)
            erdan_completed = group_erdan.get('完成', 0)
            erdan_rate = (erdan_completed / erdan_total * 100) if erdan_total > 0 else 0
            
            # 汇总统计
            total_stt_count += stt_total
            total_stt_completed += stt_completed
            total_sdy_count += sdy_total
            total_sdy_completed += sdy_completed
            total_erdan_count += erdan_total
            total_erdan_completed += erdan_completed
            
            # 综合评分（三个维度的平均）
            combined_rate = (stt_rate + sdy_rate + erdan_rate) / 3
            
            # 获取所有销售
            all_sales = set(group_sop.get('首通', {}).get('销售', {}).keys()) | \
                       set(group_sop.get('首单元', {}).get('销售', {}).keys()) | \
                       set(group_erdan.get('销售', {}).keys())
            all_sales = sorted(all_sales)
            
            # 小组汇总行
            if len(all_sales) > 0:
                ws.merge_cells(f'A{current_row}:A{current_row + len(all_sales) - 1}')
                group_cell = ws[f'A{current_row}']
                group_cell.value = group
                group_cell.fill = group_fill
                group_cell.font = group_font
                group_cell.alignment = center_align
                group_cell.border = border
            
            # 写入各销售的数据
            for sale_idx, sale in enumerate(all_sales):
                row = current_row + sale_idx
                
                # 首通
                stt_sale = group_sop.get('首通', {}).get('销售', {}).get(sale, {})
                stt_sale_total = stt_sale.get('总数', 0)
                stt_sale_completed = stt_sale.get('完成', 0)
                stt_sale_rate = (stt_sale_completed / stt_sale_total * 100) if stt_sale_total > 0 else 0
                
                # 首单元
                sdy_sale = group_sop.get('首单元', {}).get('销售', {}).get(sale, {})
                sdy_sale_total = sdy_sale.get('总数', 0)
                sdy_sale_completed = sdy_sale.get('完成', 0)
                sdy_sale_rate = (sdy_sale_completed / sdy_sale_total * 100) if sdy_sale_total > 0 else 0
                
                # 二单元
                erdan_sale = group_erdan.get('销售', {}).get(sale, {})
                erdan_sale_total = erdan_sale.get('总数', 0)
                erdan_sale_completed = erdan_sale.get('完成', 0)
                erdan_sale_rate = (erdan_sale_completed / erdan_sale_total * 100) if erdan_sale_total > 0 else 0
                
                sale_combined_rate = (stt_sale_rate + sdy_sale_rate + erdan_sale_rate) / 3
                
                # 销售名称
                b_cell = ws.cell(row=row, column=2)
                b_cell.value = sale
                b_cell.border = border
                b_cell.alignment = center_align
                
                # 首通总数
                c_cell = ws.cell(row=row, column=3)
                c_cell.value = stt_sale_total
                c_cell.border = border
                c_cell.alignment = center_align
                
                # 首通完成
                d_cell = ws.cell(row=row, column=4)
                d_cell.value = stt_sale_completed
                d_cell.border = border
                d_cell.alignment = center_align
                
                # 首通完成率
                e_cell = ws.cell(row=row, column=5)
                e_cell.value = stt_sale_rate
                e_cell.number_format = '0.0"%"'
                e_cell.border = border
                e_cell.alignment = center_align
                self._apply_rate_color(e_cell, stt_sale_rate, green_fill, yellow_fill, red_fill, green_font, yellow_font, red_font)
                
                # 首单元总数
                f_cell = ws.cell(row=row, column=6)
                f_cell.value = sdy_sale_total
                f_cell.border = border
                f_cell.alignment = center_align
                
                # 首单元完成
                g_cell = ws.cell(row=row, column=7)
                g_cell.value = sdy_sale_completed
                g_cell.border = border
                g_cell.alignment = center_align
                
                # 首单元完成率
                h_cell = ws.cell(row=row, column=8)
                h_cell.value = sdy_sale_rate
                h_cell.number_format = '0.0"%"'
                h_cell.border = border
                h_cell.alignment = center_align
                self._apply_rate_color(h_cell, sdy_sale_rate, green_fill, yellow_fill, red_fill, green_font, yellow_font, red_font)
                
                # 二单元总数
                i_cell = ws.cell(row=row, column=9)
                i_cell.value = erdan_sale_total
                i_cell.border = border
                i_cell.alignment = center_align
                
                # 二单元完成
                j_cell = ws.cell(row=row, column=10)
                j_cell.value = erdan_sale_completed
                j_cell.border = border
                j_cell.alignment = center_align
                
                # 二单元完成率
                k_cell = ws.cell(row=row, column=11)
                k_cell.value = erdan_sale_rate
                k_cell.number_format = '0.0"%"'
                k_cell.border = border
                k_cell.alignment = center_align
                self._apply_rate_color(k_cell, erdan_sale_rate, green_fill, yellow_fill, red_fill, green_font, yellow_font, red_font)
                
                # 综合评分
                l_cell = ws.cell(row=row, column=12)
                l_cell.value = sale_combined_rate
                l_cell.number_format = '0.0"%"'
                l_cell.border = border
                l_cell.alignment = center_align
                self._apply_rate_color(l_cell, sale_combined_rate, green_fill, yellow_fill, red_fill, green_font, yellow_font, red_font)
            
            current_row += len(all_sales)
        
        # 添加总汇总行
        summary_row = current_row + 1
        summary_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
        summary_font = Font(bold=True, size=11, color='000000')
        
        # 总汇总标题
        a_summary = ws.cell(row=summary_row, column=1)
        a_summary.value = '【上海SS团队总汇总】'
        a_summary.fill = summary_fill
        a_summary.font = summary_font
        a_summary.border = border
        a_summary.alignment = center_align
        
        # 计算总的完成率
        total_stt_rate = (total_stt_completed / total_stt_count * 100) if total_stt_count > 0 else 0
        total_sdy_rate = (total_sdy_completed / total_sdy_count * 100) if total_sdy_count > 0 else 0
        total_erdan_rate = (total_erdan_completed / total_erdan_count * 100) if total_erdan_count > 0 else 0
        total_combined_rate = (total_stt_rate + total_sdy_rate + total_erdan_rate) / 3
        
        # 总汇总数据
        summary_data = [
            (2, ''),
            (3, total_stt_count),
            (4, total_stt_completed),
            (5, total_stt_rate),
            (6, total_sdy_count),
            (7, total_sdy_completed),
            (8, total_sdy_rate),
            (9, total_erdan_count),
            (10, total_erdan_completed),
            (11, total_erdan_rate),
            (12, total_combined_rate),
        ]
        
        for col, val in summary_data:
            cell = ws.cell(row=summary_row, column=col)
            cell.fill = summary_fill
            cell.font = summary_font
            cell.border = border
            cell.alignment = center_align
            
            if col == 2:
                cell.value = ''
            elif col in [5, 8, 11, 12]:
                cell.value = val
                cell.number_format = '0.0"%"'
                if col in [5, 8, 11, 12]:
                    self._apply_rate_color(cell, val, green_fill, yellow_fill, red_fill, green_font, yellow_font, red_font)
            else:
                cell.value = val
        
        # 保存文件
        report_file = output_dir / f'服务绩效汇总报表_{self.report_date.strftime("%Y%m%d")}.xlsx'
        wb.save(str(report_file))
        
        return {
            'success': True,
            'groups': len(all_groups),
            'report_file': str(report_file),
            'output_dir': str(output_dir),
            'total_stt_rate': total_stt_rate,
            'total_sdy_rate': total_sdy_rate,
            'total_erdan_rate': total_erdan_rate,
            'total_combined_rate': total_combined_rate
        }


def handle_sop_report_command(sop_file, erdan_file, report_date=None, output_base=None):
    """
    处理 /SOP报表 命令
    
    Args:
        sop_file: SOP CSV 文件路径
        erdan_file: 二单元 Excel 文件路径
        report_date: 报表日期 (YYYY-MM-DD)，不指定则为当天
        output_base: 输出基础目录
    
    Returns:
        报告信息
    """
    if report_date:
        try:
            report_date = datetime.strptime(report_date, '%Y-%m-%d').date()
        except ValueError:
            return f"错误：日期格式不正确，应为 YYYY-MM-DD"
    
    generator = ServiceReportGenerator(report_date)
    result = generator.generate_report(sop_file, erdan_file, output_base)
    
    if result['success']:
        report_msg = f"""
📊 上海SS团队服务绩效汇总报表
==================================================
报告日期: {generator.report_date.strftime('%Y-%m-%d')}
统计周期: {generator.report_month.strftime('%Y-%m-01')} ~ {generator.report_date.strftime('%Y-%m-%d')}
涵盖小组数: {result['groups']}

【整体完成率】
  • 首通完成率: {result['total_stt_rate']:.1f}%
  • 首单元完成率: {result['total_sdy_rate']:.1f}%
  • 二单元完成率: {result['total_erdan_rate']:.1f}%
  • 综合评分: {result['total_combined_rate']:.1f}%

输出文件: {Path(result['report_file']).name}
输出目录: {result['output_dir']}
"""
        return report_msg.strip()
    else:
        return f"错误：生成报表失败"
