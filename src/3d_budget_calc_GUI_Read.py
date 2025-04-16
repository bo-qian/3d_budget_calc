import sys
import os
import unicodedata
import re
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QPlainTextEdit, QFormLayout, QFileDialog, QCheckBox
from PyQt5.QtGui import QFont, QFontDatabase, QFontMetrics, QIcon
from PyQt5.QtCore import Qt
from openpyxl import load_workbook  # 添加用于读取 Excel 文件的库

def get_display_width(text):
    """计算字符串的显示宽度"""
    width = 0
    for char in text:
        if unicodedata.east_asian_width(char) in ('F', 'W'):  # 全角字符
            width += 2
        else:  # 半角字符
            width += 1
    return width

def center_text(text, total_width):
    """根据显示宽度居中字符串"""
    text_width = get_display_width(text)
    padding = (total_width - text_width) // 2
    return " " * padding + text + " " * padding

def format_terminal_output(result):
    """增强型终端报表，支持对齐"""
    border = "=" * 62

    # 确保零件清单中的每个元素是字典
    parts_info = "\n".join([
        f"  零件{i+1}: {part['name']}（总体积：{part['volume'] + part['support_volume']:.3f}mm³）"
        if isinstance(part, dict) else f"  零件{i+1}: {part}"  # 如果不是字典，直接输出字符串
        for i, part in enumerate(result['输入参数']['零件清单'])
    ])

    # 使用宽度感知的居中方法
    title = " 多零件3D打印成本预算报告 "
    centered_title = center_text(title, 61)

    output = [
        f"{border}",
        centered_title,
        border,
        "[打印参数]",
        f"  零件数量：{result['输入参数']['零件数量']}件",
        f"  打印时长：{result['输入参数']['总打印时长']}",
        "\n[零件清单]",
        f"{parts_info}",
        "\n[费用明细]",
        f"{'  项目名称'.ljust(20)}{'金额'.rjust(34)}",
        "  " + "-" * 58 + "  ",
        f"  材料成本：".ljust(20) + f"¥{result['计算明细']['材料费用']:>10,.2f}".rjust(35),
        f"  机时费用：".ljust(20) + f"¥{result['计算明细']['机时费用']:>10,.2f}".rjust(35),
        f"  氩气消耗：".ljust(20) + f"¥{result['计算明细']['氩气费用']:>10,.2f}".rjust(35),
        f"  后处理费：".ljust(20) + f"¥{result['计算明细']['后处理费']:>10,.2f}".rjust(35),
        "  " + "-" * 58 + "  ",
        f"  合计金额：".ljust(20) + f"¥{result['计算明细']['总费用']:>10,.2f}".rjust(35),
        f"  折扣优惠：".ljust(20) + f"{result['定价标准']['折扣优惠']}".rjust(35),
        f"  实付金额：".ljust(20) + f"¥{result['计算明细']['实际费用']:>10,.2f}".rjust(35),
        border
    ]
    return "\n".join(output)

def calculate_multipart_cost(parts, total_print_duration, pricing_standard):
    # 总材料计算，使用零件体积和支撑体积的总和
    total_volume = sum(p['volume'] + p['support_volume'] for p in parts)
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
            "零件清单": [f"{p['name']} (零件体积：{p['volume']:.3f}mm³，支撑体积：{p['support_volume']:.3f}mm³)" for p in parts],
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

def export_to_excel(result, filename="多零件预算报告.xlsx"):
    """专业级多零件报表"""
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('预算总览')
        
        # 高级格式配置
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4F81BD', 'font_color': '#FFFFFF', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        part_name_format = workbook.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        part_detail_format = workbook.add_format({
            'bg_color': '#FCE4D6', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        currency_format = workbook.add_format({
            'num_format': '¥##0.00', 'bg_color': '#E2EFDA', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        number_format = workbook.add_format({
            'num_format': '0.00', 'bg_color': '#FFF2CC', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        normal_format = workbook.add_format({
            'bg_color': '#FFFFFF', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        
        # 标题区块
        worksheet.merge_range('A1:B1', '多零件合并打印预算报告',
                              workbook.add_format({
                                  'bold': True, 'font_size': 14, 'bg_color': '#4F81BD', 'font_color': '#FFFFFF',
                                  'align': 'center', 'border': 1
                              }))
        worksheet.merge_range('A2:B2', f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
                              normal_format)
        
        # 输入参数动态生成
        params = [
            ['总打印时长', result['输入参数']['总打印时长']],
            ['零件数量', f"{result['输入参数']['零件数量']}件"]
        ]
        
        # 修改零件体积提取逻辑，直接从字典中获取数据
        for i, part in enumerate(result['输入参数']['零件清单'], 1):
            params.extend([
                [f'零件{i}名称', part['name']],
                [f'零件{i}体积', f"{part['volume']:.3f}mm³"],
                [f'零件{i}支撑体积', f"{part['support_volume']:.3f}mm³"]
            ])
        
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
                if "零件" in label and "名称" in label:  # 零件名称行加背景颜色
                    cell_format = part_name_format
                elif "体积" in label:  # 零件体积和支撑体积行加背景颜色
                    cell_format = part_detail_format
                elif title == "费用明细":
                    if "费用" in label or "金额" in label or "后处理费" in label:  # 判断是否为货币
                        cell_format = currency_format
                    else:
                        cell_format = normal_format
                else:
                    cell_format = number_format if isinstance(value, (int, float)) else normal_format
                
                worksheet.write(row_idx, 0, label, cell_format)
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

class CostCalculatorApp(QWidget):
    def __init__(self):
        super().__init__()

        # 初始化定价标准
        self.pricing_standard = {
            "钛粉密度": 4.50,         # 单位：g/cm³
            "致密系数": 0.9995,       # 无量纲
            "用量比例": 1.5,          # 无量纲 
            "材料单价": 1800,          # 元/公斤
            "机时费率": 250,          # 元/小时
            "氩气数量": 1,            # 无量纲
            "氩气单价": 1800,         # 元
            "氩气用量": 0.8,          # 无量纲
            "后处理费": 1500,         # 元
            "折扣优惠": 1.0           # 百分比
        }

        self.parts = []  # 用于存储零件信息
        self.init_ui()

    def init_ui(self):
        # 设置窗口大小，确保内容可以完全显示
        self.setWindowTitle("3D Printing Estimator")
        self.setGeometry(100, 100, 800, 500)  # 调整窗口大小

        # 获取图标路径
        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, "3dprint.ico")
        else:
            icon_path = "3dprint.ico"

        self.setWindowIcon(QIcon(icon_path))

        main_layout = QVBoxLayout()  # 主布局，垂直分布

        # 尝试加载系统字体
        font = self.load_chinese_font()

        # 设置样式表，应用圆角框并将背景颜色改为白色
        rounded_style = """
            QLineEdit, QPushButton, QPlainTextEdit {
                border: 2px solid #8f8f91;
                border-radius: 10px;
                padding: 5px;
                background-color: #ffffff;  /* 设置背景颜色为白色 */
            }
            QLineEdit:focus, QPushButton:pressed, QPlainTextEdit:focus {
                border: 2px solid #0078d7;
            }
        """

        # 创建水平布局：左侧零件信息，右侧定价标准参数
        content_layout = QHBoxLayout()

        # 左侧布局：零件信息输入
        left_layout = QVBoxLayout()

        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignRight)  # 设置标签右对齐

        # 替换零件信息输入部分为读取 Excel 文件按钮
        load_button = QPushButton("加载零件信息 (xlsm)", self)
        load_button.setFont(font)
        load_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;  /* 绿色背景 */
                color: white;  /* 白色文字 */
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #45a049;  /* 鼠标悬停时的颜色 */
            }
            QPushButton:pressed {
                background-color: #3e8e41;  /* 按下时的颜色 */
            }
        """)
        load_button.clicked.connect(self.load_parts_from_excel)
        left_layout.addWidget(load_button)  # 将按钮添加到左侧布局

        # 零件信息框
        self.parts_display = QPlainTextEdit(self)
        self.parts_display.setFont(font)
        self.parts_display.setReadOnly(True)
        self.parts_display.setStyleSheet(rounded_style)
        self.parts_display.setLineWrapMode(QPlainTextEdit.NoWrap)  # 禁用自动换行

        # 美化滑动条样式
        self.parts_display.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # 启用垂直滚动条
        self.parts_display.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # 启用水平滚动条
        self.parts_display.verticalScrollBar().setStyleSheet("""
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background: none;
                height: 0px;
            }
        """)
        self.parts_display.horizontalScrollBar().setStyleSheet("""
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:horizontal {
                background: #c0c0c0;
                border-radius: 6px;
                min-width: 20px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #a0a0a0;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background: none;
                width: 0px;
            }
        """)

        # 调整零件信息框的高度
        self.parts_display.setFixedHeight(343)  # 设置固定高度为 250 像素

        # 将零件信息框添加到左侧布局
        left_layout.addWidget(self.parts_display)

        # 打印时长输入框
        duration_label = QLabel("打印时长", self)
        duration_label.setFont(font)
        self.duration_input = QLineEdit(self)
        self.duration_input.setText("11天11小时11分11秒")  # 设置默认值
        self.duration_input.setFont(font)
        self.duration_input.setStyleSheet(rounded_style)

        # 将打印时长输入框添加到布局
        duration_layout = QFormLayout()
        duration_layout.addRow(duration_label, self.duration_input)
        left_layout.addLayout(duration_layout)

        # 启用导出到 Excel 的复选框
        self.export_checkbox = QCheckBox("导出到 Excel 报告", self)
        self.export_checkbox.setFont(font)  # 使用加载的 Microsoft YaHei 字体
        self.export_checkbox.setChecked(False)  # 默认未选中
        self.export_checkbox.setFixedHeight(self.duration_input.sizeHint().height())  # 设置高度与打印时长输入框一致
        self.export_checkbox.setStyleSheet("""
            QCheckBox {
                spacing: 10px;  /* 文字与复选框的间距 */
                vertical-align: middle;  /* 垂直居中 */
            }
            QCheckBox::indicator {
                width: 18px;  /* 复选框宽度 */
                height: 18px;  /* 复选框高度 */
            }
            QCheckBox::indicator:unchecked {
                border: 2px solid #8f8f91;  /* 未选中时的边框颜色 */
                background-color: #ffffff;  /* 未选中时的背景颜色 */
                border-radius: 3px;  /* 圆角 */
            }
            QCheckBox::indicator:checked {
                border: 2px solid #4CAF50;  /* 选中时的边框颜色 */
                background-color: #4CAF50;  /* 选中时的背景颜色 */
                border-radius: 3px;  /* 圆角 */
            }
            QCheckBox::indicator:unchecked:hover {
                border: 2px solid #0078D7;  /* 鼠标悬停时未选中状态的边框颜色 */
            }
            QCheckBox::indicator:checked:hover {
                border: 2px solid #45a049;  /* 鼠标悬停时选中状态的边框颜色 */
                background-color: #45a049;  /* 鼠标悬停时选中状态的背景颜色 */
            }
        """)
        
        # 将复选框添加到布局中，与右侧的折扣优惠上下对齐
        duration_layout.addRow(self.export_checkbox)

        # 一键清零按钮
        clear_button = QPushButton("一键清零", self)
        clear_button.setFont(font)
        clear_button.setStyleSheet("""
            QPushButton {
                background-color: #FF5722;  /* 橙色背景 */
                color: white;  /* 白色文字 */
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #E64A19;  /* 鼠标悬停时的颜色 */
            }
            QPushButton:pressed {
                background-color: #D84315;  /* 按下时的颜色 */
            }
        """)
        clear_button.clicked.connect(self.clear_parts_display)  # 连接到修改后的方法

        # 将一键清零按钮添加到左侧布局的最下面
        left_layout.addWidget(clear_button)

        content_layout.addLayout(left_layout)

        # 右侧布局：定价标准参数输入和打印时长
        right_layout = QVBoxLayout()

        param_layout = QFormLayout()
        param_layout.setLabelAlignment(Qt.AlignRight)  # 设置标签右对齐
        self.param_inputs = {}
        for param, default_value in self.pricing_standard.items():
            label = QLabel(param, self)
            label.setFont(font)

            # 创建输入框和单位标签
            param_input_layout = QHBoxLayout()
            input_field = QLineEdit(self)
            input_field.setFont(font)
            input_field.setStyleSheet(rounded_style)
            input_field.setText(str(default_value))  # 设置默认值
            self.param_inputs[param] = input_field
            param_input_layout.addWidget(input_field)

            # 添加单位标签（如果有）
            if param == "钛粉密度":
                unit_label = QLabel("g/cm³", self)
            elif param == "材料单价":
                unit_label = QLabel("元/公斤", self)
            elif param == "机时费率":
                unit_label = QLabel("元/小时", self)
            elif param == "氩气数量":
                unit_label = QLabel("瓶", self)
            elif param == "氩气单价":
                unit_label = QLabel("元", self)
            elif param == "后处理费":
                unit_label = QLabel("元", self)
            else:
                unit_label = None

            if unit_label:
                unit_label.setFont(font)
                param_input_layout.addWidget(unit_label)

            param_layout.addRow(label, param_input_layout)

        right_layout.addLayout(param_layout)

        # 计算成本按钮
        calc_button = QPushButton("计算成本", self)
        calc_button.setFont(font)
        calc_button.setStyleSheet("""
            QPushButton {
                background-color: #0078D7;  /* 蓝色背景 */
                color: white;  /* 白色文字 */
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #005A9E;  /* 鼠标悬停时的颜色 */
            }
            QPushButton:pressed {
                background-color: #004578;  /* 按下时的颜色 */
            }
        """)
        calc_button.clicked.connect(self.calculate_cost)

        right_layout.addWidget(calc_button)

        content_layout.addLayout(right_layout)

        # 添加内容布局到主布局
        main_layout.addLayout(content_layout)

        # 设置结果显示框容器
        result_container = QWidget(self)  # 创建一个容器
        result_layout = QVBoxLayout(result_container)  # 容器内部使用垂直布局
        result_layout.setContentsMargins(0, 0, 0, 0)  # 去除容器的边距

        # 设置容器的圆角样式
        result_container.setStyleSheet("""
            QWidget {
                border: 2px solid #8f8f91;
                border-radius: 10px;  /* 设置圆角半径 */
                background-color: #ffffff;  /* 设置背景颜色为白色 */
            }
        """)

        self.result_output = QPlainTextEdit(self)
        self.result_output.setFont(QFont("Maple Mono NF CN", 10))  # 设置等宽字体
        self.result_output.setReadOnly(True)
        self.result_output.setStyleSheet(rounded_style)
        self.result_output.setLineWrapMode(QPlainTextEdit.NoWrap)  # 禁用自动换行

        # 美化滑动条样式
        self.result_output.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # 启用垂直滚动条
        self.result_output.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # 启用水平滚动条
        self.result_output.verticalScrollBar().setStyleSheet("""
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background: none;
                height: 0px;
            }
        """)
        self.result_output.horizontalScrollBar().setStyleSheet("""
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:horizontal {
                background: #c0c0c0;
                border-radius: 6px;
                min-width: 20px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #a0a0a0;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background: none;
                width: 0px;
            }
        """)

        # 设置最小高度
        self.result_output.setMinimumHeight(300)  # 设置最小高度为 200 像素

        # 启用滚动条并设置滚动条样式
        self.result_output.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # 启用垂直滚动条
        self.result_output.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # 启用水平滚动条

        # 将结果显示框添加到容器布局
        result_layout.addWidget(self.result_output)

        # 将容器添加到主布局
        main_layout.addWidget(result_container)

        # 初始隐藏结果显示框
        result_container.setVisible(False)

        # 设置主布局
        self.setLayout(main_layout)

    def load_chinese_font(self):
        # 创建字体对象
        font = QFont()
        font.setFamily("Microsoft YaHei")  # 将字体替换为微软雅黑
        font.setBold(True)  # 设置加粗

        # 设置字体大小
        font.setPointSize(12)

        # 设置字体样式策略，优先使用 Helvetica 显示英文和数字
        font.setStyleStrategy(QFont.PreferDefault)

        return font

    def load_parts_from_excel(self):
        """从 Excel 文件加载零件信息"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsm)")
        if not file_path:
            return

        try:
            workbook = load_workbook(file_path, data_only=True)
            sheet = workbook.active

            # 从 Excel 文件中读取零件信息
            part_count = int(sheet["C2"].value)
            self.parts = []
            self.parts_display.clear()

            for row in range(8, 8 + part_count):
                name = sheet[f"B{row}"].value
                volume = float(sheet[f"C{row}"].value)
                support_volume = float(sheet[f"D{row}"].value)
                self.parts.append({'name': name, 'volume': volume, 'support_volume': support_volume})
                # 修改输出格式
                self.parts_display.appendPlainText(
                    f"零件{row - 7}: {name}\n    零件体积：{volume:.3f}mm³\n    支撑体积：{support_volume:.3f}mm³"
                )

        except Exception as e:
            self.result_output.setStyleSheet("color: red; font-size: 12pt;")
            self.result_output.setPlainText(f"加载 Excel 文件失败：{e}")

    def clear_inputs(self):
        """清空所有输入框的内容"""
        self.duration_input.clear()
        for input_field in self.param_inputs.values():
            input_field.clear()
        self.result_output.clear()
        self.parts_display.clear()

    def clear_parts_display(self):
        """清空零件信息框和输出信息框的内容"""
        self.parts_display.clear()  # 清空零件信息框
        self.result_output.clear()  # 清空输出信息框
        self.parts = []  # 清空零件信息列表

    def calculate_cost(self):
        # 获取用户输入的参数值
        for param, input_field in self.param_inputs.items():
            try:
                value = float(input_field.text())
                self.pricing_standard[param] = value
            except ValueError:
                self.result_output.setStyleSheet("color: red; font-size: 12pt;")  # 设置字体为红色和大小
                self.result_output.setPlainText(f"参数 {param} 的值无效，请输入数字！")
                return

        # 调用成本计算函数
        total_print_duration = self.duration_input.text().strip()
        if not total_print_duration or not self.parts:
            self.result_output.setStyleSheet("color: red; font-size: 12pt;")  # 设置字体为红色和大小
            self.result_output.setPlainText("请先加载零件信息和填写打印时长！\n")
            return

        # 确保零件信息格式正确
        formatted_parts = [
            {'name': part['name'], 'volume': part['volume'], 'support_volume': part['support_volume']}
            for part in self.parts
        ]

        # 调用成本计算函数
        result = calculate_multipart_cost(formatted_parts, total_print_duration, self.pricing_standard)

        # 确保支撑体积在报告中正确显示
        result['输入参数']['零件清单'] = formatted_parts

        report = format_terminal_output(result)
        self.result_output.setStyleSheet("color: black; font-size: 12pt;")  # 恢复正常字体颜色
        self.result_output.setPlainText(report)

        # 显示结果显示框
        self.result_output.parentWidget().setVisible(True)

        # 检查是否启用了导出功能
        if self.export_checkbox.isChecked():
            filename, _ = QFileDialog.getSaveFileName(self, "保存为 Excel", "多零件预算报告.xlsx", "Excel 文件 (*.xlsx)")
            if filename:
                export_to_excel(result, filename)
                self.result_output.appendPlainText(f"\n报表已保存至：{filename}")

if __name__ == "__main__":
    import sys
    import os
    from PyQt5.QtGui import QIcon
    from PyQt5.QtWidgets import QApplication

    # 获取图标路径
    if hasattr(sys, '_MEIPASS'):
        icon_path = os.path.join(sys._MEIPASS, "3dprint.ico")
    else:
        icon_path = "3dprint.ico"

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(icon_path))  # 设置全局图标
    window = CostCalculatorApp()
    window.show()
    sys.exit(app.exec_())
