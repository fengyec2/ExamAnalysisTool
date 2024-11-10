import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading
import queue

class ExamAnalysisTool:
    def __init__(self, root):
        self.root = root
        self.root.title("考试成绩分析工具")
        
        self.create_widgets()
        self.filepaths = []
        self.pdf_save_directory = ""
        self.queue = queue.Queue()
        self.is_canceled = False  # 用于控制取消操作

        # 启动 GUI 更新循环
        self.root.after(100, self.process_queue)

    def create_widgets(self):
        tk.Label(self.root, text="程序版本：v1.2.2\n\n已选择的成绩文件：").pack(pady=10)

        self.file_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, width=50)
        self.file_listbox.pack(pady=5)

        self.input_file_button = tk.Button(self.root, text="选择文件", command=self.load_input_files)
        self.input_file_button.pack(pady=5)

        self.analyze_button = tk.Button(self.root, text="生成进退步系数报表", command=self.start_calculate_progress)
        self.analyze_button.pack(pady=5)

        self.line_chart_button = tk.Button(self.root, text="生成年级排名折线图", command=self.start_generate_ranking_chart)
        self.line_chart_button.pack(pady=5)

        self.create_report_button = tk.Button(self.root, text="生成历次考试成绩单", command=self.start_generate_report)
        self.create_report_button.pack(pady=5)

        self.cancel_button = tk.Button(self.root, text="取消", command=self.cancel_operation, state=tk.DISABLED)
        self.cancel_button.pack(pady=5)

        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10)

    def load_input_files(self):
        self.filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if not self.filepaths:
            return
        self.file_listbox.delete(0, tk.END)
        for filepath in self.filepaths:
            file_name = os.path.basename(filepath)  # 只获取文件名
            self.file_listbox.insert(tk.END, file_name)  # 插入文件名而不是完整路径

    def start_calculate_progress(self):
        self.cancel_button.config(state=tk.NORMAL)
        self.is_canceled = False  # 重置取消标志
        self.progress_bar['value'] = 0  # 重置进度条
        self.queue.queue.clear()  # 清空队列

        self.disable_buttons()  # 禁用功能按钮
        threading.Thread(target=self._calculate_progress_thread).start()

    def _calculate_progress_thread(self):
        if not self.filepaths:
            self.queue.put(("error", "请先选择文件"))
            self.enable_buttons()  # 任务完成后启用按钮
            return

        exam_data = {}
        exam_numbers = []
        all_exam_numbers = set()  # 用于存储所有考试编号
        duplicate_exam_numbers = set()  # 用于存储发现的重复考试编号

        for file in self.filepaths:
            if self.is_canceled:
                self.queue.put(("info", "操作已取消"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

            df = pd.read_excel(file)

            # 数据合法性检查
            if '考试编号' not in df.columns or '姓名' not in df.columns or '级名' not in df.columns:
                self.queue.put(("error", f"文件 {os.path.basename(file)} \n\n缺少必要的列: '考试编号', '姓名', '级名'"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

            # 检查当前文件中的“考试编号”是否重复
            current_exam_numbers = set(df['考试编号'])

            for exam_no in current_exam_numbers:
                if exam_no in all_exam_numbers:
                    duplicate_exam_numbers.add(exam_no)  # 记录重复的考试编号
                all_exam_numbers.add(exam_no)

            if duplicate_exam_numbers:
                self.queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

            exam_number = df['考试编号'].max()  # 获取最大考试编号
            exam_numbers.append(exam_number)

            for _, row in df.iterrows():
                if self.is_canceled:
                    self.queue.put(("info", "操作已取消"))
                    self.enable_buttons()  # 任务完成后启用按钮
                    return
                student = row['姓名']
                rank = row['级名']

                # 将年级排名保存到字典中，以学生姓名为键
                if student not in exam_data:
                    exam_data[student] = {}
                exam_data[student][exam_number] = rank

        # 获取所有考试的编号
        exam_numbers.sort(reverse=True)
        all_exam_numbers = exam_numbers  # 所有考试编号

        progress_data = []
        students = exam_data.keys()

        for student in students:
            if self.is_canceled:
                self.queue.put(("info", "操作已取消"))
                self.enable_buttons()  # 任务完成后启用按钮
                return
            
            student_ranks = {exam_no: exam_data[student][exam_no] for exam_no in all_exam_numbers if exam_no in exam_data[student]}
            sorted_ranks = sorted(student_ranks.items())

            if len(sorted_ranks) < 2:
                self.queue.put(("info", f"学生 {student} 在最近的 2 次考试中仅参加了 {len(sorted_ranks)} 次\n\n将跳过计算"))
                continue

            progress_entry = {'学生姓名': student}
            for exam_no, rank in sorted_ranks:
                progress_entry[f'第{exam_no}次考试排名'] = rank
            
            last_exam_rank = sorted_ranks[-2][1]
            current_exam_rank = sorted_ranks[-1][1]
            progress_coefficient = (last_exam_rank - current_exam_rank) / last_exam_rank
            
            # 直接存储进退步系数值
            progress_entry['进退步系数'] = progress_coefficient
            
            progress_data.append(progress_entry)

            # 更新进度条
            self.queue.put(("progress", (len(progress_data) / len(students)) * 100))  # 保持进度在 0-100 之间

        if not progress_data:
            self.queue.put(("warning", "没有有效的数据进行计算"))
            self.enable_buttons()  # 任务完成后启用按钮
            return

        save_directory = filedialog.askdirectory(title="选择保存目录")
        if not save_directory:  # 如果用户取消选择
            self.enable_buttons()  # 任务完成后启用按钮
            return

        output_file = os.path.join(save_directory, "进退步系数.xlsx")

        # 检查文件是否已存在
        if os.path.exists(output_file):
            response = messagebox.askyesno(
                "文件已存在",
                f"文件 {output_file} 已存在，是否覆盖该文件？"
            )
            if not response:
                self.enable_buttons()  # 任务完成后启用按钮
                return

        try:
            progress_df = pd.DataFrame(progress_data)
            progress_df.to_excel(output_file, index=False)
            self.queue.put(("info", f"进退步系数已保存至 {output_file}"))
        except PermissionError:
            self.queue.put(("error", f"无法保存文件，因为文件 {output_file} 已被占用或打开。"))
            self.enable_buttons()  # 任务完成后启用按钮
            return

        self.progress_bar['value'] = 0
        self.enable_buttons()  # 任务完成后启用按钮

    def start_generate_ranking_chart(self):
        self.cancel_button.config(state=tk.NORMAL)
        self.is_canceled = False  # 重置取消标志
        self.progress_bar['value'] = 0  # 重置进度条
        self.queue.queue.clear()  # 清空队列

        self.disable_buttons()  # 禁用功能按钮
        threading.Thread(target=self._generate_ranking_chart_thread).start()

    def _generate_ranking_chart_thread(self):
        try:
            if not self.filepaths:
                self.queue.put(("error", "请先选择文件"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

            combined_df = pd.DataFrame()
            skip_files = []  # 跳过的文件列表
            all_exam_numbers = set()  # 用于存储所有考试编号
            duplicate_exam_numbers = set()  # 用于存储发现的重复考试编号

            # 预先设定一个读取文件的数量
            total_files = len(self.filepaths)

            for idx, file in enumerate(self.filepaths):
                if self.is_canceled:
                    self.queue.put(("info", "操作已取消"))
                    self.enable_buttons()  # 任务完成后启用按钮
                    return
                
                df = pd.read_excel(file)
                # 数据合法性检查
                if '考试编号' not in df.columns or '姓名' not in df.columns or '级名' not in df.columns:
                    response = messagebox.askyesno(
                        "缺少必要列",
                        f"文件 {os.path.basename(file)} \n\n缺少必要的列: '考试编号', '姓名', '级名'\n\n是否跳过该表格绘制折线图？"
                    )
                    if not response:  # 用户选择不继续
                        self.enable_buttons()  # 任务完成后启用按钮
                        return
                    skip_files.append(file)  # 添加到跳过的文件列表
                    continue
                
                # 检查当前文件中的“考试编号”是否重复
                current_exam_numbers = set(df['考试编号'])

                for exam_no in current_exam_numbers:
                    if exam_no in all_exam_numbers:
                        duplicate_exam_numbers.add(exam_no)  # 记录重复的考试编号
                    all_exam_numbers.add(exam_no)

                if duplicate_exam_numbers:
                    self.queue.put(("error", f"发现重复的考试编号: {', '.join(map(str, duplicate_exam_numbers))}"))
                    self.enable_buttons()  # 任务完成后启用按钮
                    return

                combined_df = pd.concat([combined_df, df], ignore_index=True)

                # 进度条更新，在读取每个文件后更新
                self.queue.put(("progress", ((idx + 1) / total_files) * 100))  # 更新到文件处理进度
                if self.is_canceled:
                    self.queue.put(("info", "操作已取消"))
                    self.enable_buttons()  # 任务完成后启用按钮
                    return

            if combined_df.empty:
                self.queue.put(("warning", "所有文件均缺少必要的数据，无法生成年级排名折线图"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

            from matplotlib import rcParams
            rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
            rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

            students = combined_df['姓名'].unique()
            self.progress_bar['maximum'] = len(students)

            # 询问用户选择保存目录
            if not self.pdf_save_directory:
                self.pdf_save_directory = filedialog.askdirectory(title="选择PDF保存目录")
                if not self.pdf_save_directory:  # 如果用户取消选择
                    self.enable_buttons()  # 任务完成后启用按钮
                    return

            for idx, student in enumerate(students):
                student_data = combined_df[combined_df['姓名'] == student]
                try:
                    plt.figure()
                    plt.plot(student_data['考试编号'], student_data['级名'], marker='o', label=student)
                    plt.title(f'{student} 年级排名折线图')
                    plt.xlabel('考试编号')
                    plt.ylabel('年级排名')
                    plt.gca().invert_yaxis()  # 翻转Y轴
                    plt.legend()
                    plt.grid()
                    output_path = os.path.join(self.pdf_save_directory, f'{student}_年级排名折线图.pdf')
                    print(f"正在保存图表到: {output_path}")
                    plt.savefig(output_path)  # 保存图表
                    plt.close()

                    # 更新进度条
                    self.queue.put(("progress", ((idx + 1) / len(students)) * 100))  # 输出文件后更新进度条

                    if self.is_canceled:
                        self.queue.put(("info", "操作已取消"))
                        self.enable_buttons()  # 任务完成后启用按钮
                        return

                except Exception as e:
                    self.queue.put(("error", f"生成学生 {student} 的 PDF 时出现错误: {e}"))
                    continue

            self.queue.put(("info", "年级排名折线图已生成."))
            self.progress_bar['value'] = 0
            self.enable_buttons()  # 任务完成后启用按钮

            # 生成完毕后清空保存目录
            self.pdf_save_directory = ""  # 这样下次点击时会重新询问

        except Exception as e:
            self.queue.put(("error", f"处理过程中发生错误: {e}"))
            self.enable_buttons()  # 任务完成后启用按钮

    def start_generate_report(self):
        self.cancel_button.config(state=tk.NORMAL)
        self.is_canceled = False  # 重置取消标志
        self.progress_bar['value'] = 0  # 重置进度条
        self.queue.queue.clear()  # 清空队列

        self.disable_buttons()  # 禁用功能按钮
        threading.Thread(target=self._generate_report_thread).start()

    def _generate_report_thread(self):
        if not self.filepaths:
            self.queue.put(("error", "请先选择文件"))
            self.enable_buttons()  # 任务完成后启用按钮
            return

        combined_df = pd.DataFrame()
    
        for file in self.filepaths:
            if self.is_canceled:
                self.queue.put(("info", "操作已取消"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

            try:
                df = pd.read_excel(file)
            except Exception as e:
                self.queue.put(("error", f"无法读取文件 {os.path.basename(file)}: {str(e)}"))
                self.enable_buttons()
                return

            # 数据合法性检查
            if '考试编号' not in df.columns or '姓名' not in df.columns or '级名' not in df.columns:
                self.queue.put(("error", f"文件 {os.path.basename(file)} \n\n缺少必要的列: '考试编号', '姓名', '级名'"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

            # 合并数据
            combined_df = pd.concat([combined_df, df], ignore_index=True)

        if combined_df.empty:
            self.queue.put(("warning", "没有有效的数据进行生成报告"))
            self.enable_buttons()  # 任务完成后启用按钮
            return

        # 对数据进行排序
        combined_df.sort_values(by=['姓名', '考试编号'], inplace=True)

        # 获取所有学生的唯一列表
        students = combined_df['姓名'].unique()

        progress_data = []

        # 为每个学生生成报表
        for student in students:
            student_data = combined_df[combined_df['姓名'] == student]
            progress_data.append(student_data)
        
            if self.is_canceled:
                self.queue.put(("info", "操作已取消"))
                self.enable_buttons()  # 任务完成后启用按钮
                return

        # 询问用户选择保存目录
        save_directory = filedialog.askdirectory(title="选择保存目录")
        if not save_directory:  # 如果用户取消选择
            self.enable_buttons()  # 任务完成后启用按钮
            return

        # 保存每个学生的成绩
        for student in students:
            student_report = combined_df[combined_df['姓名'] == student]
            output_file = os.path.join(save_directory, f'{student}_成绩单.xlsx')

            # 检查文件是否已被占用
            try:
                # 尝试打开文件以检查是否被占用
                if os.path.exists(output_file):
                    with open(output_file, 'a'):
                        pass
                
                # 尝试保存文件
                student_report.to_excel(output_file, index=False)

            except PermissionError:
                self.queue.put(("error", f"无法保存文件，因为文件 {output_file} 已被占用或打开。"))
                continue  # 跳过此文件的保存

            # 更新进度条
            self.queue.put(("progress", (len(progress_data) / len(students)) * 100))  # 保持进度在 0-100 之间

        self.queue.put(("info", "历次考试成绩单已生成。"))
        self.progress_bar['value'] = 0
        self.enable_buttons()  # 任务完成后启用按钮

    def cancel_operation(self):
        self.is_canceled = True  # 设置取消标志
        self.cancel_button.config(state=tk.DISABLED)

    def enable_buttons(self):
        self.input_file_button.config(state=tk.NORMAL)
        self.analyze_button.config(state=tk.NORMAL)
        self.line_chart_button.config(state=tk.NORMAL)
        self.create_report_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)  # 确保取消按钮禁用

    def disable_buttons(self):
        self.input_file_button.config(state=tk.DISABLED)
        self.analyze_button.config(state=tk.DISABLED)
        self.line_chart_button.config(state=tk.DISABLED)
        self.create_report_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)  # 启用取消按钮

    def process_queue(self):
        # 处理队列中的消息
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

        self.root.after(100, self.process_queue)  # 每100毫秒处理队列

if __name__ == "__main__":
    root = tk.Tk()
    app = ExamAnalysisTool(root)
    root.mainloop()