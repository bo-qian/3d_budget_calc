import re
import pandas as pd
from datetime import datetime

def calculate_multipart_cost(parts, total_print_duration, pricing_standard):
    # 总材料计算
    total_volume = sum(p['volume'] for p in parts)
    material_weight_g = (total_volume * 1e-3 * pricing_standard["钛粉密度"]
                         * pricing_standard["用量比例"] * pricing_standard["致密系数"])
    material_cost = material_weight_g * pricing_standard["材料单价"] * 1e-3

    # 机时费用
    machine_hours = convert_duration_to_hours(total_print_duration)
    machine_cost = machine_hours * pricing_standard["机时费率"]

    # 其他费用
    argon_cost = pricing_standard["氩气单价"] * pricing_standard["氩气用量"] * pricing_standard["氩气数量"]
    post_processing = pricing_standard["后处理费"]

    # 费用汇总
    total_cost = material_cost + machine_cost + argon_cost + post_processing
    actual_cost = total_cost * pricing_standard["折扣优惠"]

    return {
        "输入参数": {
            "零件清单": [f"{p['name']} ({p['volume']:.3f}mm³)" for p in parts],
            "总打印时长": total_print_duration,
            "零件数量": len(parts)
        },
        "定价标准": pricing_standard,
        "计算明细": {
            "材料费用": round(material_cost, 2),
            "机时费用": round(machine_cost, 2),
            "氩气费用": round(argon_cost, 2),
            "后处理费": round(post_processing, 2),
            "总费用": round(total_cost, 2),
            "实际费用": round(actual_cost, 2)
        }
    }

def convert_duration_to_hours(duration_str):
    # 初始化时间单位
    days = hours = minutes = seconds = 0
    
    # 使用正则表达式提取各时间单位
    patterns = {
        'days': r'(\d+)天',
        'hours': r'(\d+)小时',
        'minutes': r'(\d+)分',
        'seconds': r'(\d+)秒'
    }
    
    for unit, pattern in patterns.items():
        match = re.search(pattern, duration_str)
        if match:
            value = int(match.group(1))
            if unit == 'days':
                days = value
            elif unit == 'hours':
                hours = value
            elif unit == 'minutes':
                minutes = value
            elif unit == 'seconds':
                seconds = value
    
    # 转换为总小时数（保留3位小数）
    total_hours = (
        days * 24 + 
        hours + 
        minutes / 60 + 
        seconds / 3600
    )

    # 四舍五入到2位小数
    return total_hours

def format_terminal_output(result):
    """增强型终端报表"""
    border = "=" * 61
    parts_info = "\n".join([f"  零件{i+1}: {name}" 
                          for i, name in enumerate(result['输入参数']['零件清单'])])
    
    output = [
        f"\n{border}",
        " 多零件3D打印成本预算报告 ".center(50,'◈'),
        border,
        f"\n[打印参数] 零件数量：{result['输入参数']['零件数量']}件",
        f"         总打印时长：{result['输入参数']['总打印时长']}",
        f"\n[零件清单]\n{parts_info}",
        
        "\n[费用明细]".ljust(25)+"金额".rjust(30),
        f"材料成本：{result['计算明细']['材料费用']:>15,.2f}¥".rjust(55),
        f"机时费用：{result['计算明细']['机时费用']:>15,.2f}¥".rjust(55),
        f"氩气消耗：{result['计算明细']['氩气费用']:>15,.2f}¥".rjust(55),
        f"后处理费：{result['计算明细']['后处理费']:>15,.2f}¥".rjust(55),
        "-"*60,
        f"合计金额：{result['计算明细']['总费用']:>15,.2f}¥".rjust(55),
        f"折扣优惠：{result['定价标准']['折扣优惠']:>14}折".rjust(54),
        f"实付金额：{result['计算明细']['实际费用']:>15,.2f}¥".rjust(55),
        border
    ]
    return "\n".join(output)


def export_to_txt(result, filename="预算报告.txt"):
    """将终端报表写入txt文件"""
    # 生成终端输出内容（网页1基础方法）
    content = format_terminal_output(result)
    
    # 使用with语句安全写入（网页2推荐方法）
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f"终端报表已保存至：{filename}")


def export_to_excel(result, filename="多零件预算报告.xlsx"):
    """专业级多零件报表"""
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('预算总览')
        
        # 高级格式配置
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        currency_format = workbook.add_format({
            'num_format': '¥##0.00', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        number_format = workbook.add_format({
            'num_format': '0.00', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        normal_format = workbook.add_format({
            'align': 'center', 'valign': 'vcenter',
            'border': 1
        })
        
        # 标题区块
        worksheet.merge_range('A1:B1', '多零件合并打印预算报告',
                              workbook.add_format({
                                  'bold': True, 'font_size': 14,
                                  'align': 'center', 'border': 1
                              }))
        worksheet.merge_range('A2:B2', f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
                              normal_format)
        
        # 输入参数动态生成
        params = [
            ['总打印时长', result['输入参数']['总打印时长']],
            ['零件数量', f"{result['输入参数']['零件数量']}件"]
        ]
        
        # 修改零件体积提取逻辑
        for i, part in enumerate(result['输入参数']['零件清单'], 1):
            match = re.search(r'\(([\d.]+)mm³\)', part)
            if match:
                volume = float(match.group(1))  # 提取括号中的体积值
                params.extend([
                    [f'零件{i}名称', part.split('(')[0].strip()],
                    [f'零件{i}体积', f"{volume:,.3f}mm³"]
                ])
            else:
                raise ValueError(f"无法解析零件体积：{part}")
        
        # 定义定价标准的单位
        pricing_units = {
            "钛粉密度": "g/cm³",
            "致密系数": "",  # 无单位
            "用量比例": "",  # 无单位
            "材料单价": "元/公斤",
            "机时费率": "元/小时",
            "氩气数量": "瓶",
            "氩气单价": "元",
            "氩气用量": "瓶",  # 无单位
            "后处理费": "元",
            "折扣优惠": ""  # 无单位
        }
        
        # 为定价标准添加单位
        pricing_standard_with_units = [
            [param, f"{value} {pricing_units.get(param, '')}".strip()]
            for param, value in result['定价标准'].items()
        ]
        
        # 数据写入逻辑
        def write_section(data, start_row, title):
            worksheet.merge_range(start_row, 0, start_row, 1, title, header_format)
            for row_idx, (label, value) in enumerate(data, start_row + 1):
                if title == "费用明细":
                    if "费用" in label or "金额" in label or "后处理费" in label:  # 判断是否为货币
                        cell_format = currency_format
                    else:
                        cell_format = normal_format
                else:
                    cell_format = number_format if isinstance(value, (int, float)) else normal_format
                
                worksheet.write(row_idx, 0, label, normal_format)
                worksheet.write(row_idx, 1, value, cell_format)
            return start_row + len(data) + 2
        
        current_row = 3
        current_row = write_section(params, current_row, "输入参数")
        current_row = write_section(pricing_standard_with_units, current_row, "定价标准")
        current_row = write_section(
            [[k, v] for k, v in result['计算明细'].items()],
            current_row, "费用明细"
        )
        
        # 智能列宽设置
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:B', 25)
        
        print(f"\n专业级报表已生成：{filename}")