# File: ExamAnalysisTool.py

import os
import threading
import queue
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog

class FileHandler:
    """文件处理"""
    def __init__(self):
        self.filepaths = []

    def load_files(self):
        """选择文件并更新列表"""
        filepaths = filedialog.askopenfilenames(title="选择文件", filetypes=[("Excel files", "*.xlsx")])
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
        save_directory = filedialog.askdirectory(title="选择保存目录")
        if not save_directory:
            queue.put(("info", "操作已取消"))
            return

        # 输出文件
        output_file = os.path.join(save_directory, "进退步系数.xlsx")
        try:
            pd.DataFrame(progress_data).to_excel(output_file, index=False)
            queue.put(("info", f"进退步系数报表已保存至 {output_file}"))
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
                queue.put(("progress", (idx + 1) / len(students)))
            except Exception as e:
                queue.put(("error", f"生成学生 {student} 的图表时出现错误: {e}"))

        queue.put(("progress", 1.0))
        queue.put(("info", "年级排名折线图已生成"))

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
                queue.put(("progress", (idx + 1) / len(students)))
            except PermissionError:
                queue.put(("error", f"无法保存文件，因为文件 {student_report_path} 已被占用或打开。"))

        queue.put(("progress", 1.0))
        queue.put(("info", "历次考试成绩单已生成"))

class FileCard(ctk.CTkFrame):
    def __init__(self, master, filepath, remove_callback=None):
        super().__init__(master, fg_color=("gray90", "gray13"))
        self.filepath = filepath
        self.remove_callback = remove_callback
        self._create_widgets()

    def _create_widgets(self):
        # 文件图标
        self.icon_label = ctk.CTkLabel(self, text="📄", width=30)
        self.icon_label.pack(side="left", padx=5)

        # 文件名和路径
        text_frame = ctk.CTkFrame(self, fg_color="transparent")
        text_frame.pack(side="left", fill="x", expand=True)
        
        self.name_label = ctk.CTkLabel(text_frame, text=os.path.basename(self.filepath), 
                                      font=ctk.CTkFont(weight="bold"))
        self.name_label.pack(anchor="w")
        
        self.path_label = ctk.CTkLabel(text_frame, text=self.filepath, 
                                      text_color=("gray40", "gray60"), font=ctk.CTkFont(size=12))
        self.path_label.pack(anchor="w")

        # 删除按钮
        self.remove_btn = ctk.CTkButton(self, text="×", width=30, height=30, 
                                      fg_color="transparent", hover_color=("gray80", "gray20"),
                                      command=self._on_remove)
        self.remove_btn.pack(side="right", padx=5)

        # 添加悬停效果
        self.bind("<Enter>", lambda e: self.configure(fg_color=("gray85", "gray15")))
        self.bind("<Leave>", lambda e: self.configure(fg_color=("gray90", "gray13")))

        # 添加文件类型校验图标
        file_ext = os.path.splitext(self.filepath)[1].lower()
        icon = "📊" if file_ext == ".xlsx" else "❓"
        self.icon_label.configure(text=icon)

    def _on_remove(self):
        if self.remove_callback:
            self.remove_callback(self.filepath)
        self.destroy()

class ExamAnalysisToolGUI:
    """主页面"""
    def __init__(self):
        self.root = ctk.CTk()  # 创建 CTk 窗口
        self.root.title("考试成绩分析工具")
        self.root.geometry("800x400")  # 设置窗口默认大小
        
        self.file_handler = FileHandler()  # 需要实现 FileHandler 类
        self.queue = queue.Queue()
        self.is_canceled = False
        self.is_on_top = False

        self.file_format_variable = tk.StringVar(value="pdf")  # 单选按钮变量

        self.init_ui()
        self.setup_menu()  # 初始化菜单栏
        self.timer = self.root.after(100, self.process_queue)
        self.progress_bar.set(0)

    def init_ui(self):
        """初始化 UI"""
        self.central_widget = ctk.CTkFrame(self.root)
        self.central_widget.pack(padx=20, pady=20, fill="both", expand=True)

        # 左侧区域
        left_frame = ctk.CTkFrame(self.central_widget)
        left_frame.pack(side="left", padx=10, pady=10, fill="both", expand=True)

        self.file_label = ctk.CTkLabel(left_frame, text="已选择的成绩文件：")
        self.file_label.pack(pady=10)

        self.file_scrollframe = ctk.CTkScrollableFrame(left_frame, width=250, height=200)
        self.file_scrollframe.pack(padx=10, pady=10, fill="both", expand=True)

        # 右侧区域
        right_frame = ctk.CTkFrame(self.central_widget)
        right_frame.pack(side="right", padx=10, pady=10, fill="both", expand=True)

        self.input_file_button = ctk.CTkButton(right_frame, text="选择文件", command=self.load_input_files)
        self.input_file_button.pack(pady=10)

        self.analyze_button = ctk.CTkButton(right_frame, text="生成进退步系数报表", command=self.start_calculate_progress)
        self.analyze_button.pack(pady=10)

        self.chart_button = ctk.CTkButton(right_frame, text="生成年级排名折线图", command=self.start_generate_ranking_charts)
        self.chart_button.pack(pady=10)

        # 添加单选按钮
        pdf_png_frame = ctk.CTkFrame(right_frame)
        pdf_png_frame.pack(pady=10)

        self.pdf_radio = ctk.CTkRadioButton(pdf_png_frame, text="输出为 PDF", variable=self.file_format_variable, value="pdf")
        self.pdf_radio.pack(side="left", padx=10)

        self.png_radio = ctk.CTkRadioButton(pdf_png_frame, text="输出为 PNG", variable=self.file_format_variable, value="png")
        self.png_radio.pack(side="left", padx=10)

        self.report_button = ctk.CTkButton(right_frame, text="生成历次考试成绩单", command=self.start_generate_report)
        self.report_button.pack(pady=10)

        self.cancel_button = ctk.CTkButton(right_frame, text="取消", state="disabled", command=self.cancel_operation)
        self.cancel_button.pack(pady=10)

        self.progress_bar = ctk.CTkProgressBar(right_frame, width=300)
        self.progress_bar.pack(pady=10)

    def setup_menu(self):
        """设置菜单栏"""
        self.root.option_add("*Font", "SimHei 20")  # 设置全局菜单字体
        menubar = tk.Menu(self.root)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)

        help_menu.add_command(label="关于", command=self.show_about_dialog)
        help_menu.add_command(label="置顶", command=self.toggle_top)

        self.root.config(menu=menubar)

    def toggle_top(self):
        """切换窗口置顶状态"""
        if self.is_on_top:
            self.root.attributes("-topmost", False)
            self.is_on_top = False
        else:
            self.root.attributes("-topmost", True)
            self.is_on_top = True

    def show_about_dialog(self):
        """显示关于对话框"""
        about_message = """\
        考试成绩分析工具
        版本：1.4.1
        作者: fengyec2
        许可证：GPL-3.0 license
        项目地址：github.com/fengyec2/ExamAnalysisTool
        """
        messagebox.showinfo("关于", about_message)

    def load_input_files(self):
        """文件选择并更新列表"""
        filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        self._add_files(filepaths)

    def _add_files(self, filepaths):
        """统一添加文件方法"""
        for fp in filepaths:
            if fp not in self.file_handler.filepaths:
                self.file_handler.filepaths.append(fp)
                card = FileCard(
                    self.file_scrollframe, 
                    fp, 
                    remove_callback=self._remove_file
                )
                card.pack(fill="x", pady=2)

    def _remove_file(self, filepath):
        """删除文件回调"""
        if filepath in self.file_handler.filepaths:
            self.file_handler.filepaths.remove(filepath)

    def start_calculate_progress(self):
        """独立线程处理"""
        self.is_canceled = False
        self.progress_bar.set(0)
        self.queue.queue.clear()
        self.disable_buttons()
        threading.Thread(target=self.calculate_progress_thread).start()

    def calculate_progress_thread(self):
        """计算进退步系数"""
        ProgressCalculator.calculate_progress(self.file_handler.filepaths, lambda: self.is_canceled, self.queue)
        self.enable_buttons()

    def start_generate_ranking_charts(self):
        """独立线程处理"""
        save_directory = filedialog.askdirectory(title="选择 PDF/PNG 保存目录")
        if not save_directory:
            return

        file_format = self.file_format_variable.get()

        self.is_canceled = False
        self.progress_bar.set(0)
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
        save_directory = filedialog.askdirectory(title="选择 Excel 保存目录")
        if not save_directory:
            return

        self.is_canceled = False
        self.progress_bar.set(0)
        self.queue.queue.clear()
        self.disable_buttons()
        threading.Thread(target=self.generate_report_thread, args=(save_directory,)).start()

    def generate_report_thread(self, save_directory):
        """生成历次考试成绩单"""
        HistoricalReportGenerator.generate_report(self.file_handler.filepaths, save_directory, lambda: self.is_canceled, self.queue)
        self.enable_buttons()

    def cancel_operation(self):
        """取消操作"""
        self.is_canceled = True

    def enable_buttons(self):
        """启用按钮"""
        self.input_file_button.configure(state="normal")
        self.analyze_button.configure(state="normal")
        self.chart_button.configure(state="normal")
        self.report_button.configure(state="normal")
        self.cancel_button.configure(state="disabled")
        self.pdf_radio.configure(state="normal")
        self.png_radio.configure(state="normal")

    def disable_buttons(self):
        """禁用按钮"""
        self.input_file_button.configure(state="disabled")
        self.analyze_button.configure(state="disabled")
        self.chart_button.configure(state="disabled")
        self.report_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.pdf_radio.configure(state="disabled")
        self.png_radio.configure(state="disabled")

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
                self.progress_bar.set(msg_content)
        self.timer = self.root.after(100, self.process_queue)

    def run(self):
        """运行应用"""
        self.root.mainloop()

if __name__ == "__main__":
    app = ExamAnalysisToolGUI()
    app.run()
