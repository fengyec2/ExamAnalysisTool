# File: exam_analysis_tool.py

import os
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import matplotlib.pyplot as plt
import webbrowser

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

class ExamAnalysisToolGUI:
    """主页面"""
    def __init__(self, root):
        self.root = root
        self.root.title("考试成绩分析工具")
        self.file_handler = FileHandler()
        self.queue = queue.Queue()
        self.is_canceled = False 
        
        # 创建菜单栏
        self.create_menu()

        self.create_widgets()
        self.root.after(100, self.process_queue)

    def create_menu(self):
        """创建菜单栏和关于菜单"""
        menu_bar = tk.Menu(self.root)  # 创建菜单栏

        # 创建帮助菜单
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)  # 添加“关于”项
        menu_bar.add_cascade(label="帮助", menu=help_menu)  # 将帮助菜单添加到菜单栏

        self.root.config(menu=menu_bar)  # 配置根窗口使用菜单栏

    def show_about(self):
        """显示关于对话框"""
        about_window = tk.Toplevel(self.root)  # 创建新窗口
        about_window.title("关于")  # 设置窗口标题
        about_window.geometry("400x300")  # 设置窗口大小

        # 显示程序信息
        info_text = """
        考试成绩分析工具

        版本: 1.3.0
        作者: fengyec2
        许可证: GPL-3.0 license
        开源地址: https://github.com/fengyec2/ExamAnalysisTool
        引用的第三方库:
            - pandas
            - matplotlib
            - openpyxl
            - tk
        """

        label = tk.Label(about_window, text=info_text, justify=tk.LEFT, padx=10, pady=10)
        label.pack(fill="both", expand=True)

    def create_widgets(self):
        tk.Label(self.root, text="已选择的成绩文件：").pack(pady=10)
        self.file_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, width=50)
        self.file_listbox.pack(pady=5)

        self.input_file_button = tk.Button(self.root, text="选择文件", command=self.load_input_files)
        self.input_file_button.pack(pady=5)

        self.analyze_button = tk.Button(self.root, text="生成进退步系数报表", command=self.start_calculate_progress)
        self.analyze_button.pack(pady=5)

        self.chart_button = tk.Button(self.root, text="生成年级排名折线图", command=self.start_generate_ranking_charts)
        self.chart_button.pack(pady=5)

        self.report_button = tk.Button(self.root, text="生成历次考试成绩单", command=self.start_generate_report)
        self.report_button.pack(pady=5)

        self.cancel_button = tk.Button(self.root, text="取消", command=self.cancel_operation, state=tk.DISABLED)
        self.cancel_button.pack(pady=5)

        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10)

    def load_input_files(self):
        """文件选择并更新列表"""
        self.file_listbox.delete(0, tk.END)
        filepaths = self.file_handler.load_files()
        for filepath in filepaths:
            self.file_listbox.insert(tk.END, os.path.basename(filepath))

    def start_calculate_progress(self):
        """独立线程处理"""
        self.is_canceled = False
        self.progress_bar['value'] = 0
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
        save_directory = filedialog.askdirectory(title="选择PDF保存目录")
        if not save_directory:
            return

        self.is_canceled = False
        self.progress_bar['value'] = 0
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
        save_directory = filedialog.askdirectory(title="选择保存目录")
        if not save_directory:
            return

        self.is_canceled = False
        self.progress_bar['value'] = 0
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
        self.input_file_button.config(state=tk.NORMAL)
        self.analyze_button.config(state=tk.NORMAL)
        self.chart_button.config(state=tk.NORMAL)
        self.report_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)

    def disable_buttons(self):
        """禁用按钮"""
        self.input_file_button.config(state=tk.DISABLED)
        self.analyze_button.config(state=tk.DISABLED)
        self.chart_button.config(state=tk.DISABLED)
        self.report_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)

    def process_queue(self):
        """信息处理"""
        while not self.queue.empty():
            msg_type, msg_content = self.queue.get()
            if msg_type == "info":
                messagebox.showinfo("信息", msg_content)
            elif msg_type == "warning":
                messagebox.showwarning("警告", msg_content)
            elif msg_type == "error":
                messagebox.showerror("错误", msg_content)
            elif msg_type == "progress":
                self.progress_bar['value'] = msg_content
        self.root.after(100, self.process_queue)
if __name__ == "__main__":
    root = tk.Tk()
    app = ExamAnalysisToolGUI(root)
    root.mainloop()