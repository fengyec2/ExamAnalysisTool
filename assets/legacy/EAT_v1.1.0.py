import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading

class ExamAnalysisTool:
    def __init__(self, root):
        self.root = root
        self.root.title("考试成绩分析工具")

        self.create_widgets()
        self.filepaths = []
        self.pdf_save_directory = ""  # 添加一个属性用于保存PDF的目录

    def create_widgets(self):
        tk.Label(self.root, text="选择成绩文件:").pack(pady=10)

        self.file_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, width=50)
        self.file_listbox.pack(pady=5)

        self.input_file_button = tk.Button(self.root, text="选择文件", command=self.load_input_files)
        self.input_file_button.pack(pady=5)

        self.analyze_button = tk.Button(self.root, text="计算进退步系数", command=self.calculate_progress)
        self.analyze_button.pack(pady=5)

        self.line_chart_button = tk.Button(self.root, text="生成年级排名折线图", command=self.generate_ranking_chart)
        self.line_chart_button.pack(pady=5)

        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10)

    def load_input_files(self):
        self.filepaths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if not self.filepaths:
            return
        self.file_listbox.delete(0, tk.END)
        for filepath in self.filepaths:
            self.file_listbox.insert(tk.END, filepath)

    def calculate_progress(self):
        if not self.filepaths:
            messagebox.showwarning("警告", "请先选择文件")
            return

        exam_data = {}
        exam_numbers = []

        for file in self.filepaths:
            df = pd.read_excel(file)

            # 数据合法性检查
            if '考试编号' not in df.columns or '同学' not in df.columns or '年级排名' not in df.columns:
                messagebox.showerror("错误", f"文件 {os.path.basename(file)} \n\n缺少必要的列: '考试编号', '同学', '年级排名'")
                return

            exam_number = df['考试编号'].max()  # 获取最大考试编号
            exam_numbers.append(exam_number)
            for _, row in df.iterrows():
                student = row['同学']
                rank = row['年级排名']

                # 将年级排名保存到字典中，以学生姓名为键
                if student not in exam_data:
                    exam_data[student] = {}
                exam_data[student][exam_number] = rank

        # 获取所有考试的编号
        exam_numbers.sort(reverse=True)
        all_exam_numbers = exam_numbers  # 所有考试编号

        # 检查每个学生在最近两次考试中的出现次数
        progress_data = []
        students = exam_data.keys()

        for student in students:
            # 计算该学生的所有考试排名
            student_ranks = {exam_no: exam_data[student][exam_no] for exam_no in all_exam_numbers if exam_no in exam_data[student]}
            sorted_ranks = sorted(student_ranks.items())

            if len(sorted_ranks) < 2:
                messagebox.showinfo("信息", f"同学 {student} 在最近的 2 次考试中仅参加了 {len(sorted_ranks)} 次\n\n将跳过计算")
                continue

            # 准备数据存储
            progress_entry = {'学生姓名': student}
            
            for exam_no, rank in sorted_ranks:
                progress_entry[f'第{exam_no}次考试排名'] = rank
            
            # 计算进退步系数
            last_exam_rank = sorted_ranks[-2][1]
            current_exam_rank = sorted_ranks[-1][1]
            progress_coefficient = (last_exam_rank - current_exam_rank) / last_exam_rank
            
            if progress_coefficient > 1:
                marked_coefficient = f"🟩{progress_coefficient}"
            elif -1 < progress_coefficient < 1:
                marked_coefficient = f"🟦{progress_coefficient}"
            else:
                marked_coefficient = f"🟥{progress_coefficient}"
            progress_entry['进退步系数'] = marked_coefficient
            
            progress_data.append(progress_entry)

        if not progress_data:
            messagebox.showwarning("警告", "没有有效的数据进行计算")
            return

        # 询问用户选择保存目录
        save_directory = filedialog.askdirectory(title="选择保存目录")
        if not save_directory:  # 如果用户取消选择
            return

        # 构建输出文件路径
        output_file = os.path.join(save_directory, "进退步系数.xlsx")

        # 检查是否覆盖输出文件
        if os.path.exists(output_file):
            if not messagebox.askyesno("确认覆盖", f"文件 {output_file} 已存在，您希望覆盖吗？"):
                return

        progress_df = pd.DataFrame(progress_data)
        progress_df.to_excel(output_file, index=False)
        messagebox.showinfo("信息", f"进退步系数已保存至 {output_file}")

        self.progress_bar['value'] = 0

    def generate_ranking_chart(self):
        if not self.filepaths:
            messagebox.showwarning("警告", "请先选择文件")
            return

        combined_df = pd.DataFrame()
        skip_files = []  # 跳过的文件列表

        for idx, file in enumerate(self.filepaths):
            df = pd.read_excel(file)
            # 数据合法性检查
            if '考试编号' not in df.columns or '同学' not in df.columns or '年级排名' not in df.columns:
                response = messagebox.askyesno(
                    "缺少必要列",
                    f"文件 {os.path.basename(file)} \n\n缺少必要的列: '考试编号', '同学', '年级排名'\n\n是否跳过该表格绘制折线图？"
                )
                if not response:  # 用户选择不继续
                    return
                skip_files.append(file)  # 添加到跳过的文件列表
                continue
            
            combined_df = pd.concat([combined_df, df], ignore_index=True)

        if combined_df.empty:
            messagebox.showwarning("警告", "所有文件均缺少必要的数据，无法生成年级排名折线图")
            return

        if '考试编号' not in combined_df.columns or '同学' not in combined_df.columns or '年级排名' not in combined_df.columns:
            messagebox.showerror("错误", "所有文件必须包含列: '考试编号', '同学', '年级排名'")
            return

        from matplotlib import rcParams
        rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
        rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

        students = combined_df['同学'].unique()
        self.progress_bar['maximum'] = len(students)

        # 询问用户选择保存目录 (仅在本次生成中询问一次)
        if not self.pdf_save_directory:
            self.pdf_save_directory = filedialog.askdirectory(title="选择PDF保存目录")
            if not self.pdf_save_directory:  # 如果用户取消选择
                return

        for student in students:
            student_data = combined_df[combined_df['同学'] == student]
            plt.figure()
            plt.plot(student_data['考试编号'], student_data['年级排名'], marker='o', label=student)
            plt.title(f'{student} 年级排名折线图')
            plt.xlabel('考试编号')
            plt.ylabel('年级排名')
            plt.gca().invert_yaxis()  # 翻转Y轴
            plt.legend()
            plt.grid()
            plt.savefig(os.path.join(self.pdf_save_directory, f'{student}_年级排名折线图.pdf'))  # 使用用户选择的目录
            plt.close()

            self.progress_bar['value'] += 1
            self.root.update_idletasks()

        messagebox.showinfo("信息", "年级排名折线图已生成.")

        self.progress_bar['value'] = 0

        # 生成完毕后清空保存目录
        self.pdf_save_directory = ""  # 这样下次点击时会重新询问

if __name__ == "__main__":
    root = tk.Tk()
    app = ExamAnalysisTool(root)
    root.mainloop()