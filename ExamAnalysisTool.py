# File: ExamAnalysisTool.py

import os
import threading
import queue
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QMenu, QAction, QMessageBox, QFileDialog, QProgressBar, QListWidget, QPushButton, QLabel, QVBoxLayout, QWidget, QHBoxLayout, QRadioButton
from PyQt5.QtCore import Qt

class FileHandler:
    """文件处理"""
    def __init__(self):
        self.filepaths = []

    def load_files(self):
        """选择文件并更新列表"""
        options = QFileDialog.Options()
        filepaths, _ = QFileDialog.getOpenFileNames(None, "选择文件", "", "Excel files (*.xlsx)", options=options)
        self.filepaths = filepaths
        return self.filepaths

class DataProcessor:
    """处理Excel文件的通用方法"""
    @staticmethod
    def read_excel(file, queue):
        """读取Excel文件并返回DataFrame"""
        try:
            df = pd.read_excel(file)
            return df
        except Exception as e:
            queue.put(("error", f"无法读取文件 {os.path.basename(file)}: {str(e)}"))
            return None

    @staticmethod
    def validate_data(df, required_columns, queue):
        """验证DataFrame的列是否完整"""
        for col in required_columns:
            if col not in df.columns:
                queue.put(("error", f"文件缺少必要的列: '{col}'"))
                return False
        return True

    @staticmethod
    def check_duplicate_exam_numbers(current_exam_numbers, all_exam_numbers, queue):
        """检查重复考试编号"""
        duplicate_exam_numbers = set()
        for exam_no in current_exam_numbers:
            if exam_no in all_exam_numbers:
                duplicate_exam_numbers.add(exam_no)
            all_exam_numbers.add(exam_no)
        return duplicate_exam_numbers

class ProgressCalculator:
    """生成进退步系数报表"""
    
    @staticmethod
    def calculate_progress(filepaths, is_canceled_callback, queue):
        exam_data = {}
        exam_numbers = []
        all_exam_numbers = set()  # 用于存储所有考试编号

        # 计算进退步系数
        for file in filepaths:
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return None

            df = DataProcessor.read_excel(file, queue)
            if df is None:
                return None

            if not DataProcessor.validate_data(df, ['考试编号', '姓名', '级名'], queue):
                return None

            current_exam_numbers = set(df['考试编号'])
            duplicate_exam_numbers = DataProcessor.check_duplicate_exam_numbers(current_exam_numbers, all_exam_numbers, queue)
            if duplicate_exam_numbers:
                queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                return None

            exam_number = df['考试编号'].max()  # 获取最大考试编号
            exam_numbers.append(exam_number)

            for _, row in df.iterrows():
                student = row['姓名']
                rank = row['级名']
                if student not in exam_data:
                    exam_data[student] = {}
                exam_data[student][exam_number] = rank

        exam_numbers.sort(reverse=True)
        all_exam_numbers = exam_numbers
        progress_data = []

        for student in exam_data.keys():
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return None

            student_ranks = {exam_no: exam_data[student][exam_no] for exam_no in all_exam_numbers if exam_no in exam_data[student]}
            sorted_ranks = sorted(student_ranks.items())

            if len(sorted_ranks) < 2:
                queue.put(("info", f"学生 {student} 在最近的 2 次考试中仅参加了 {len(sorted_ranks)} 次，将跳过计算"))
                continue

            progress_entry = {'学生姓名': student}
            for exam_no, rank in sorted_ranks:
                progress_entry[f'第{exam_no}次考试排名'] = rank

            last_exam_rank = sorted_ranks[-2][1]
            current_exam_rank = sorted_ranks[-1][1]
            progress_coefficient = (last_exam_rank - current_exam_rank) / last_exam_rank
            progress_entry['进退步系数'] = progress_coefficient
            progress_data.append(progress_entry)

        # 询问保存
        save_directory = QFileDialog.getExistingDirectory(None, "选择保存目录")
        if not save_directory:
            queue.put(("info", "操作已取消"))
            return

        # 输出文件
        output_file = os.path.join(save_directory, "进退步系数.xlsx")
        try:
            pd.DataFrame(progress_data).to_excel(output_file, index=False)
            queue.put(("info", f"进退步系数已保存至 {output_file}"))
        except PermissionError:
            queue.put(("error", f"无法保存文件，因为文件 {output_file} 已被占用或打开。"))

class RankingChartGenerator:
    """生成年级排名折线图"""
    
    @staticmethod
    def generate_ranking_charts(filepaths, save_directory, is_canceled_callback, queue, file_format='pdf'):
        # 设置matplotlib中文支持
        matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 设置中文字体为 SimHei（黑体）
        matplotlib.rcParams['axes.unicode_minus'] = False    # 防止负号显示为方块
        combined_df = pd.DataFrame()
        all_exam_numbers = set()

        for file in filepaths:
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return

            df = DataProcessor.read_excel(file, queue)
            if df is None:
                return None

            if not DataProcessor.validate_data(df, ['考试编号', '姓名', '级名'], queue):
                return None

            current_exam_numbers = set(df['考试编号'])
            duplicate_exam_numbers = DataProcessor.check_duplicate_exam_numbers(current_exam_numbers, all_exam_numbers, queue)
            if duplicate_exam_numbers:
                queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                return None

            combined_df = pd.concat([combined_df, df], ignore_index=True)

        if combined_df.empty:
            queue.put(("warning", "没有有效的数据生成折线图"))
            return

        students = combined_df['姓名'].unique()
        for idx, student in enumerate(students):
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return

            # 确保数据按照考试编号排序（全局排序）
            combined_df['考试编号'] = pd.to_numeric(combined_df['考试编号'], errors='coerce')
            combined_df = combined_df.dropna(subset=['考试编号'])  # 删除无效考试编号的行
            combined_df = combined_df.sort_values(by='考试编号')

            student_data = combined_df[combined_df['姓名'] == student]
            try:
                plt.figure()
                plt.plot(student_data['考试编号'], student_data['级名'], marker='o', label=student)
                plt.title(f'{student} 年级排名折线图')
                plt.xlabel('考试编号')
                plt.ylabel('年级排名')
                # 设置 x 轴刻度为整数
                x_ticks = student_data['考试编号'].astype(int)  # 取整数部分
                plt.xticks(x_ticks)  # 设置 x 轴的刻度为整数
                plt.gca().invert_yaxis()  # 翻转 Y 轴
                plt.legend()
                plt.grid()
                
                # 根据用户选择的文件格式保存文件
                output_file = os.path.join(save_directory, f'{student}_年级排名折线图.{file_format}')
                plt.savefig(output_file, dpi=300)  # 将 dpi 设置为 300
                plt.close()
                queue.put(("progress", ((idx + 1) / len(students)) * 100))
            except Exception as e:
                queue.put(("error", f"生成学生 {student} 的图表时出现错误: {e}"))

        queue.put(("info", "折线图已生成"))

class HistoricalReportGenerator:
    """生成历次考试成绩单"""
    
    @staticmethod
    def generate_report(filepaths, save_directory, is_canceled_callback, queue):
        combined_df = pd.DataFrame()
        all_exam_numbers = set()

        # 合并数据
        for file in filepaths:
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return

            df = DataProcessor.read_excel(file, queue)
            if df is None:
                return None

            if not DataProcessor.validate_data(df, ['考试编号', '姓名', '级名'], queue):
                return None

            current_exam_numbers = set(df['考试编号'])
            duplicate_exam_numbers = DataProcessor.check_duplicate_exam_numbers(current_exam_numbers, all_exam_numbers, queue)
            if duplicate_exam_numbers:
                queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                return None

            combined_df = pd.concat([combined_df, df], ignore_index=True)

        if combined_df.empty:
            queue.put(("warning", "没有有效的数据生成成绩单"))
            return

        # 按学生分类
        students = combined_df['姓名'].unique()

        # 生成报表
        for idx, student in enumerate(students):
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return

            # 筛选每位学生的成绩
            student_data = combined_df[combined_df['姓名'] == student]
            
            # 进行排序
            student_data = student_data.sort_values(by='考试编号', ascending=True)

            # 保存文件
            student_report_path = os.path.join(save_directory, f"{student}_成绩单.xlsx")
            try:
                # 包含所有列
                student_data.to_excel(student_report_path, index=False)
                queue.put(("progress", ((idx + 1) / len(students)) * 100))
            except PermissionError:
                queue.put(("error", f"无法保存文件，因为文件 {student_report_path} 已被占用或打开。"))

        queue.put(("info", "所有学生的成绩单已生成"))

class ExamAnalysisToolGUI(QMainWindow):
    """主页面"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("考试成绩分析工具")
        self.resize(600, 800)  # 设置窗口默认大小
        self.setWindowIcon(QIcon("assets/img/eat.ico"))
        
        self.file_handler = FileHandler()  # 需要实现 FileHandler 类
        self.queue = queue.Queue()
        self.is_canceled = False
        self.setAcceptDrops(True)
        self.is_on_top = False

        self.init_ui()
        self.setup_menu()  # 初始化菜单栏
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.process_queue)
        self.timer.start(100)

    def init_ui(self):
        """初始化 UI"""
        self.central_widget = QWidget()
        layout = QVBoxLayout(self.central_widget)

        self.file_label = QLabel("已选择的成绩文件：")
        layout.addWidget(self.file_label)

        self.file_listbox = QListWidget()
        self.file_listbox.setAcceptDrops(True)  # 允许拖拽
        self.file_listbox.setDragEnabled(False)
        self.file_listbox.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)  # 自定义右键菜单
        self.file_listbox.customContextMenuRequested.connect(self.show_context_menu)  # 连接右键菜单

        layout.addWidget(self.file_listbox)

        self.input_file_button = QPushButton("选择文件")
        self.input_file_button.clicked.connect(self.load_input_files)
        layout.addWidget(self.input_file_button)

        self.analyze_button = QPushButton("生成进退步系数报表")
        self.analyze_button.clicked.connect(self.start_calculate_progress)
        layout.addWidget(self.analyze_button)

        self.chart_button = QPushButton("生成年级排名折线图")
        self.chart_button.clicked.connect(self.start_generate_ranking_charts)
        layout.addWidget(self.chart_button)

        # 添加横向排列的单选按钮
        self.radio_button_group = QtWidgets.QButtonGroup(self)
        self.pdf_radio = QRadioButton("输出为 PDF")
        self.png_radio = QRadioButton("输出为 PNG")
        
        self.pdf_radio.setChecked(True)  # 默认选择PDF

        h_layout = QHBoxLayout()
        h_layout.addWidget(self.pdf_radio)
        h_layout.addWidget(self.png_radio)
        layout.addLayout(h_layout)

        self.report_button = QPushButton("生成历次考试成绩单")
        self.report_button.clicked.connect(self.start_generate_report)
        layout.addWidget(self.report_button)

        self.cancel_button = QPushButton("取消")
        self.cancel_button.setDisabled(True)
        self.cancel_button.clicked.connect(self.cancel_operation)
        layout.addWidget(self.cancel_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)

        self.setCentralWidget(self.central_widget)

    def setup_menu(self):
        """设置菜单栏"""
        menubar = self.menuBar()
        help_menu = menubar.addMenu("帮助")

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

        toggle_top_action = QAction("置顶", self)
        toggle_top_action.triggered.connect(self.toggle_top)
        help_menu.addAction(toggle_top_action)

    def toggle_top(self):
        """切换窗口置顶状态"""
        if self.is_on_top:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)  # 取消置顶
            self.is_on_top = False
        else:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)  # 设置置顶
            self.is_on_top = True
        self.show()  # 需要调用 show() 使窗口更新

    def show_about_dialog(self):
        """显示关于对话框"""
        about_message = """\
        考试成绩分析工具
        版本：1.3.2
        作者: fengyec2
        许可证：GPL-3.0 license
        项目地址：github.com/fengyec2/ExamAnalysisTool
        """
        QMessageBox.information(self, "关于", about_message)

    def dragEnterEvent(self, event):
        """处理拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()  # 接受拖拽操作
        else:
            event.ignore()

    def dropEvent(self, event):
        """处理放置事件"""
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.endswith(".xlsx"):  # 仅接受 .xlsx 文件
                if file_path not in self.file_handler.filepaths:
                    self.file_handler.filepaths.append(file_path)
                    self.file_listbox.addItem(os.path.basename(file_path))

    def show_context_menu(self, position):
        """显示右键菜单"""
        menu = QMenu()
        add_action = QAction("添加...", self)
        delete_action = QAction("删除选定文件", self)
        add_action.triggered.connect(self.load_input_files)
        delete_action.triggered.connect(lambda: self.remove_selected_file())
        menu.addAction(add_action)
        menu.addAction(delete_action)
        menu.exec_(self.file_listbox.mapToGlobal(position))

    def remove_selected_file(self):
        """移除选中的文件"""
        selected_items = self.file_listbox.selectedItems()
        if not selected_items:
            return  # 没有选中项，直接返回

        for item in selected_items:
            # 获取显示的文件名
            filepath = item.text()
            # 获取完整的文件路径
            full_path = None
            for file in self.file_handler.filepaths:
                if os.path.basename(file) == filepath:
                    full_path = file
                    break
            # 匹配完整路径删除
            if full_path:
                self.file_handler.filepaths.remove(full_path)
                self.file_listbox.takeItem(self.file_listbox.row(item))

    def load_input_files(self):
        """文件选择并更新列表"""
        filepaths, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "Excel files (*.xlsx)")
        for filepath in filepaths:
            if filepath not in self.file_handler.filepaths:
                self.file_handler.filepaths.append(filepath)
                self.file_listbox.addItem(os.path.basename(filepath))

    def start_calculate_progress(self):
        """独立线程处理"""
        self.is_canceled = False
        self.progress_bar.setValue(0)
        self.queue.queue.clear()
        self.disable_buttons()
        threading.Thread(target=self.calculate_progress_thread).start()

    def calculate_progress_thread(self):
        """计算进退步系数"""
        ProgressCalculator.calculate_progress(
            self.file_handler.filepaths, lambda: self.is_canceled, self.queue)
        self.enable_buttons()

    def start_generate_ranking_charts(self):
        """独立线程处理"""
        save_directory = QFileDialog.getExistingDirectory(self, "选择 PDF 保存目录")
        if not save_directory:
            return

        # 获取用户选择的文件格式
        file_format = 'pdf' if self.pdf_radio.isChecked() else 'png'

        self.is_canceled = False
        self.progress_bar.setValue(0)
        self.queue.queue.clear()
        self.disable_buttons()
        threading.Thread(target=self.generate_ranking_charts_thread, args=(save_directory, file_format)).start()

    def generate_ranking_charts_thread(self, save_directory, file_format):
        """生成年级排名折线图"""
        RankingChartGenerator.generate_ranking_charts(
            self.file_handler.filepaths, save_directory, lambda: self.is_canceled, self.queue, file_format)
        self.enable_buttons()

    def start_generate_report(self):
        """独立线程处理"""
        save_directory = QFileDialog.getExistingDirectory(self, "选择 Excel 保存目录")
        if not save_directory:
            return

        self.is_canceled = False
        self.progress_bar.setValue(0)
        self.queue.queue.clear()
        self.disable_buttons()
        threading.Thread(target=self.generate_report_thread, args=(save_directory,)).start()

    def generate_report_thread(self, save_directory):
        """生成历次考试成绩单"""
        HistoricalReportGenerator.generate_report(
            self.file_handler.filepaths, save_directory, lambda: self.is_canceled, self.queue)
        self.enable_buttons()

    def cancel_operation(self):
        """取消操作"""
        self.is_canceled = True

    def enable_buttons(self):
        """启用按钮"""
        self.input_file_button.setEnabled(True)
        self.analyze_button.setEnabled(True)
        self.chart_button.setEnabled(True)
        self.report_button.setEnabled(True)
        self.cancel_button.setDisabled(True)
        self.pdf_radio.setEnabled(True)
        self.png_radio.setEnabled(True)

    def disable_buttons(self):
        """禁用按钮"""
        self.input_file_button.setDisabled(True)
        self.analyze_button.setDisabled(True)
        self.chart_button.setDisabled(True)
        self.report_button.setDisabled(True)
        self.cancel_button.setEnabled(True)
        self.pdf_radio.setDisabled(True)
        self.png_radio.setDisabled(True)

    def process_queue(self):
        """信息处理"""
        while not self.queue.empty():
            msg_type, msg_content = self.queue.get()
            if msg_type == "info":
                QMessageBox.information(self, "信息", msg_content)
            elif msg_type == "warning":
                QMessageBox.warning(self, "警告", msg_content)
            elif msg_type == "error":
                QMessageBox.critical(self, "错误", msg_content)
            elif msg_type == "progress":
                self.progress_bar.setValue(int(msg_content))

if __name__ == "__main__":
    import sys
    from PyQt5 import QtWidgets

    app = QtWidgets.QApplication(sys.argv)
    window = ExamAnalysisToolGUI()
    window.show()
    sys.exit(app.exec_())