import sys
import unicodedata
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QPlainTextEdit, QFormLayout, QFileDialog, QCheckBox
from PyQt5.QtGui import QFont, QFontDatabase, QFontMetrics, QIcon
from PyQt5.QtCore import Qt
from cost_module import calculate_multipart_cost, export_to_excel

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
    parts_info = "\n".join([f"  零件{i+1}: {name}" 
                            for i, name in enumerate(result['输入参数']['零件清单'])])
    
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

class CostCalculatorApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("3D Printing Estimator")
        self.setGeometry(100, 100, 800, 600)

        # 设置程序图标
        self.setWindowIcon(QIcon("3dprint.ico"))  # 替换为你的图标文件路径

        # 初始化定价标准
        self.pricing_standard = {
            "钛粉密度": 4.50,        # 单位：g/cm³
            "致密系数": 0.9995,      # 无量纲
            "用量比例": 1.5,          # 无量纲 
            "材料单价": 900,          # 元/公斤
            "机时费率": 250,          # 元/小时
            "氩气数量": 1,            # 无量纲
            "氩气单价": 1800,         # 元
            "氩气用量": 0.8,          # 无量纲
            "后处理费": 1500,         # 元
            "折扣优惠": 0.8             # 百分比
        }

        self.parts = []  # 用于存储零件信息
        self.init_ui()

    def init_ui(self):
        # 设置窗口大小，确保内容可以完全显示
        self.setWindowTitle("3D Printing Estimator")
        self.setGeometry(100, 100, 800, 500)  # 调整窗口大小

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

        # 零件名称输入框
        name_label = QLabel("零件名称", self)
        name_label.setFont(font)
        self.name_input = QLineEdit(self)
        self.name_input.setFont(font)
        self.name_input.setStyleSheet(rounded_style)
        form_layout.addRow(name_label, self.name_input)

        # 零件体积输入框和单位
        volume_layout = QHBoxLayout()
        self.volume_input = QLineEdit(self)
        self.volume_input.setFont(font)
        self.volume_input.setStyleSheet(rounded_style)
        volume_label = QLabel("mm³", self)
        volume_label.setFont(font)
        volume_layout.addWidget(self.volume_input)
        volume_layout.addWidget(volume_label)
        volume_label = QLabel("零件体积", self)
        volume_label.setFont(font)  # 确保字体与其他标签一致
        form_layout.addRow(volume_label, volume_layout)

        # 添加零件按钮
        add_button = QPushButton("添加零件", self)
        add_button.setFont(font)
        add_button.setStyleSheet("""
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
        add_button.clicked.connect(self.add_part)
        form_layout.addRow(add_button)

        # 将 form_layout 添加到左侧布局
        left_layout.addLayout(form_layout)

        # 零件信息框
        self.parts_display = QPlainTextEdit(self)
        self.parts_display.setFont(font)
        self.parts_display.setReadOnly(True)
        self.parts_display.setStyleSheet(rounded_style)

        # 调整零件信息框的高度
        self.parts_display.setFixedHeight(250)  # 设置固定高度为 250 像素

        # 将零件信息框添加到左侧布局，紧靠“添加零件”按钮
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
        self.export_checkbox = QCheckBox("启用导出到 Excel", self)
        self.export_checkbox.setFont(font)  # 使用加载的 PingFang SC 字体
        self.export_checkbox.setChecked(False)  # 默认未选中
        self.export_checkbox.setFixedHeight(self.duration_input.sizeHint().height())  # 设置高度与打印时长输入框一致
        self.export_checkbox.setStyleSheet("""
            QCheckBox {
                spacing: 10px;  /* 文字与复选框的间距 */
                font-size: 12pt;  /* 字体大小，与右侧一致 */
                color: #333333;  /* 文字颜色 */
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

        # 设置最小高度
        self.result_output.setMinimumHeight(300)  # 设置最小高度为 200 像素

        # 启用滚动条并设置滚动条样式
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
        font.setFamily("PingFang SC")  # 设置中文字体为苹方
        font.setBold(True)  # 设置加粗

        # 设置字体大小
        font.setPointSize(12)

        # 设置字体样式策略，优先使用 Helvetica 显示英文和数字
        font.setStyleStrategy(QFont.PreferDefault)

        return font

    def add_part(self):
        name = self.name_input.text().strip()
        try:
            volume = float(self.volume_input.text())
        except ValueError:
            self.result_output.setPlainText("体积必须为数字！")
            return

        self.parts.append({'name': name, 'volume': volume})
        self.parts_display.appendPlainText(f"{name} - {volume:.2f}mm³")
        self.name_input.clear()
        self.volume_input.clear()

    def clear_inputs(self):
        """清空所有输入框的内容"""
        self.name_input.clear()
        self.volume_input.clear()
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
            self.result_output.setPlainText("请先填写零件信息和打印时长！\n")
            return

        result = calculate_multipart_cost(self.parts, total_print_duration, self.pricing_standard)
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
    app = QApplication(sys.argv)
    window = CostCalculatorApp()
    window.show()
    sys.exit(app.exec_())
