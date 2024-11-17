# File: exam_analysis_tool.py

import os
import threading
import queue
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from PyQt5 import QtWidgets, QtCore, Qt
from PyQt5.QtWidgets import QMenu, QAction, QMessageBox, QFileDialog, QProgressBar, QListWidget, QPushButton, QLabel, QVBoxLayout, QWidget
import sys

class FileHandler:
    """文件处理"""
    def __init__(self):
        self.filepaths = []

    def load_files(self):
        """选择文件并更新列表"""

        self.filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        return self.filepaths

class ProgressCalculator:
    """生成进退步系数报表"""
    @staticmethod
    def calculate_progress(filepaths, is_canceled_callback, queue):
        exam_data = {}
        exam_numbers = []
        all_exam_numbers = set()  # 用于存储所有考试编号
        duplicate_exam_numbers = set()  # 用于存储发现的重复考试编号

        # 计算进退步系数
        for file in filepaths:
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return None

            try:
                df = pd.read_excel(file)
            except Exception as e:
                queue.put(("error", f"无法读取文件 {os.path.basename(file)}: {str(e)}"))
                return None

            # Data validation
            if '考试编号' not in df.columns or '姓名' not in df.columns or '级名' not in df.columns:
                queue.put(("error", f"文件 {os.path.basename(file)} 缺少必要的列: '考试编号', '姓名', '级名'"))
                return None

            current_exam_numbers = set(df['考试编号'])
            for exam_no in current_exam_numbers:
                if exam_no in all_exam_numbers:
                    duplicate_exam_numbers.add(exam_no)  # 记录重复的考试编号
                all_exam_numbers.add(exam_no)

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
        students = exam_data.keys()

        for student in students:
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
        save_directory = filedialog.askdirectory(title="选择保存目录")
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

        """计算进退步系数"""
        exam_data = {}
        exam_numbers = []
        all_exam_numbers = set()
        duplicate_exam_numbers = set()

        for file in filepaths:
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return None

            try:
                df = pd.read_excel(file)
            except Exception as e:
                queue.put(("error", f"无法读取文件 {os.path.basename(file)}: {str(e)}"))
                return None

            # Data validation
            if '考试编号' not in df.columns or '姓名' not in df.columns or '级名' not in df.columns:
                queue.put(("error", f"文件 {os.path.basename(file)} 缺少必要的列: '考试编号', '姓名', '级名'"))
                return None

            current_exam_numbers = set(df['考试编号'])
            for exam_no in current_exam_numbers:
                if exam_no in all_exam_numbers:
                    duplicate_exam_numbers.add(exam_no)
                all_exam_numbers.add(exam_no)

            if duplicate_exam_numbers:
                queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                return None

            exam_number = df['考试编号'].max()
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
        students = exam_data.keys()

        for student in students:
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

        return progress_data

class RankingChartGenerator:
    """生成年级排名折线图"""
    @staticmethod
    def generate_ranking_charts(filepaths, save_directory, is_canceled_callback, queue):
        # 设置matplotlib中文支持
        matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 设置中文字体为 SimHei（黑体）
        matplotlib.rcParams['axes.unicode_minus'] = False    # 防止负号显示为方块
        combined_df = pd.DataFrame()
        duplicate_exam_numbers = set()
        all_exam_numbers = set()

        for file in filepaths:
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return

            try:
                df = pd.read_excel(file)
            except Exception as e:
                queue.put(("error", f"无法读取文件 {os.path.basename(file)}: {str(e)}"))
                return

            # Data validation
            if '考试编号' not in df.columns or '姓名' not in df.columns or '级名' not in df.columns:
                queue.put(("warning", f"文件 {os.path.basename(file)} 缺少必要的列: '考试编号', '姓名', '级名'"))
                continue

            current_exam_numbers = set(df['考试编号'])
            for exam_no in current_exam_numbers:
                if exam_no in all_exam_numbers:
                    duplicate_exam_numbers.add(exam_no)
                all_exam_numbers.add(exam_no)

            if duplicate_exam_numbers:
                queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                return

            combined_df = pd.concat([combined_df, df], ignore_index=True)

        if combined_df.empty:
            queue.put(("warning", "没有有效的数据生成折线图"))
            return

        students = combined_df['姓名'].unique()
        for idx, student in enumerate(students):
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return

            student_data = combined_df[combined_df['姓名'] == student]
            try:
                plt.figure()
                plt.plot(student_data['考试编号'], student_data['级名'], marker='o', label=student)
                plt.title(f'{student} 年级排名折线图')
                plt.xlabel('考试编号')
                plt.ylabel('年级排名')
                plt.gca().invert_yaxis() # 翻转 Y 轴
                plt.legend()
                plt.grid()
                output_path = os.path.join(save_directory, f'{student}_年级排名折线图.pdf')
                plt.savefig(output_path)
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
        duplicate_exam_numbers = set()
        all_exam_numbers = set()

        # 合并数据
        for file in filepaths:
            if is_canceled_callback():
                queue.put(("info", "操作已取消"))
                return

            try:
                df = pd.read_excel(file)
            except Exception as e:
                queue.put(("error", f"无法读取文件 {os.path.basename(file)}: {str(e)}"))
                return

            # 检验数据
            if '考试编号' not in df.columns or '姓名' not in df.columns or '级名' not in df.columns:
                queue.put(("warning", f"文件 {os.path.basename(file)} 缺少必要的列: '考试编号', '姓名', '级名'"))
                continue

            current_exam_numbers = set(df['考试编号'])
            for exam_no in current_exam_numbers:
                if exam_no in all_exam_numbers:
                    duplicate_exam_numbers.add(exam_no)
                all_exam_numbers.add(exam_no)

            if duplicate_exam_numbers:
                queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                return

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

class ExamAnalysisToolGUI(QWidget):
    """主页面"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("考试成绩分析工具")
        self.resize(600, 800)  # 设置窗口默认大小
        self.file_handler = FileHandler()
        self.queue = queue.Queue()
        self.is_canceled = False

        self.init_ui()
        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.process_queue)
        self.timer.start(100)

    def init_ui(self):
        """初始化 UI"""
        layout = QVBoxLayout()

        self.file_label = QLabel("已选择的成绩文件：")
        layout.addWidget(self.file_label)

        self.file_listbox = QListWidget()
        self.file_listbox.setAcceptDrops(True)  # 允许拖放操作
        self.file_listbox.dragEnterEvent = self.dragEnterEvent  # 设置拖入事件
        self.file_listbox.dropEvent = self.dropEvent  # 设置放下事件
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

        self.setLayout(layout)

    def dragEnterEvent(self, event):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        """拖拽放下事件"""
        for url in event.mimeData().urls():
            filepath = url.toLocalFile()
            if filepath.endswith(".xlsx") and filepath not in self.file_handler.filepaths:
                self.file_handler.filepaths.append(filepath)
                self.file_listbox.addItem(filepath)

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
        for item in selected_items:
            filepath = item.text()
            self.file_handler.filepaths.remove(filepath)
            self.file_listbox.takeItem(self.file_listbox.row(item))

    def load_input_files(self):
        """文件选择并更新列表"""
        filepaths, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "Excel files (*.xlsx)")
        for filepath in filepaths:
            if filepath not in self.file_handler.filepaths:
                self.file_handler.filepaths.append(filepath)
                self.file_listbox.addItem(filepath)

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
        save_directory = QFileDialog.getExistingDirectory(self, "选择PDF保存目录")
        if not save_directory:
            return

        self.is_canceled = False
        self.progress_bar.setValue(0)
        self.queue.queue.clear()
        self.disable_buttons()
        threading.Thread(target=self.generate_ranking_charts_thread, args=(save_directory,)).start()

    def generate_ranking_charts_thread(self, save_directory):
        """生成年级排名折线图"""
        RankingChartGenerator.generate_ranking_charts(
            self.file_handler.filepaths, save_directory, lambda: self.is_canceled, self.queue)
        self.enable_buttons()

    def start_generate_report(self):
        """独立线程处理"""
        save_directory = QFileDialog.getExistingDirectory(self, "选择保存目录")
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

    def disable_buttons(self):
        """禁用按钮"""
        self.input_file_button.setDisabled(True)
        self.analyze_button.setDisabled(True)
        self.chart_button.setDisabled(True)
        self.report_button.setDisabled(True)
        self.cancel_button.setEnabled(True)

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
    app = QtWidgets.QApplication(sys.argv)
    window = ExamAnalysisToolGUI()
    window.show()
    sys.exit(app.exec_())
