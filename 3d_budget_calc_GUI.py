import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTextEdit, QFormLayout
from PyQt5.QtGui import QFont, QFontDatabase
from cost_module import calculate_multipart_cost, format_terminal_output, export_to_excel

class CostCalculatorApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("3D Printing Estimator")
        self.setGeometry(100, 100, 600, 400)

        self.parts = []  # 用于存储零件信息
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 尝试加载系统字体
        font = self.load_chinese_font()

        # 创建零件信息输入框
        form_layout = QFormLayout()
        self.name_input = QLineEdit(self)
        self.name_input.setFont(font)  # 设置控件字体
        form_layout.addRow("零件名称", self.name_input)

        self.volume_input = QLineEdit(self)
        self.volume_input.setFont(font)  # 设置控件字体
        form_layout.addRow("零件体积 (mm³)", self.volume_input)

        add_button = QPushButton("添加零件", self)
        add_button.setFont(font)  # 设置按钮字体
        add_button.clicked.connect(self.add_part)
        form_layout.addRow(add_button)

        self.parts_display = QTextEdit(self)
        self.parts_display.setFont(font)  # 设置控件字体
        self.parts_display.setReadOnly(True)
        layout.addLayout(form_layout)
        layout.addWidget(self.parts_display)

        self.duration_input = QLineEdit(self)
        self.duration_input.setPlaceholderText("输入总打印时长 (如：0天4小时11分46秒)")
        self.duration_input.setFont(font)  # 设置控件字体
        layout.addWidget(self.duration_input)

        button_layout = QHBoxLayout()
        calc_button = QPushButton("计算成本", self)
        calc_button.setFont(font)  # 设置按钮字体
        calc_button.clicked.connect(self.calculate_cost)

        export_button = QPushButton("导出 Excel", self)
        export_button.setFont(font)  # 设置按钮字体
        export_button.clicked.connect(self.export_excel)

        button_layout.addWidget(calc_button)
        button_layout.addWidget(export_button)

        layout.addLayout(button_layout)

        self.result_output = QTextEdit(self)
        self.result_output.setFont(font)  # 设置控件字体
        self.result_output.setReadOnly(True)
        layout.addWidget(self.result_output)

        self.setLayout(layout)

    def load_chinese_font(self):
        # 尝试加载 "Microsoft YaHei"（微软雅黑），如果不可用则加载 SimHei（黑体）
        font = QFont("Microsoft YaHei", 12)  # 默认字体为微软雅黑
        if not font.exactMatch():
            font = QFont("SimHei", 12)  # 如果微软雅黑不可用，使用黑体
        return font

    def add_part(self):
        name = self.name_input.text().strip()
        try:
            volume = float(self.volume_input.text())
        except ValueError:
            self.result_output.setPlainText("体积必须为数字！")
            return

        self.parts.append({'name': name, 'volume': volume})
        self.parts_display.append(f"{name} - {volume:.2f}mm³")
        self.name_input.clear()
        self.volume_input.clear()

    def calculate_cost(self):
        total_print_duration = self.duration_input.text().strip()
        if not total_print_duration or not self.parts:
            self.result_output.setPlainText("请先填写零件信息和打印时长！")
            return

        result = calculate_multipart_cost(self.parts, total_print_duration)
        report = format_terminal_output(result)
        self.result_output.setPlainText(report)

    def export_excel(self):
        if not hasattr(self, 'result'):
            self.result_output.setPlainText("请先计算成本！")
            return

        filename, _ = QFileDialog.getSaveFileName(self, "保存为 Excel", "多零件预算报告.xlsx", "Excel 文件 (*.xlsx)")
        if filename:
            export_to_excel(self.result, filename)
            self.result_output.setPlainText(f"报表已保存至：{filename}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CostCalculatorApp()
    window.show()
    sys.exit(app.exec_())
