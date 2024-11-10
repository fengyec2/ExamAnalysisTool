import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

class ExamAnalysisTool:
    def __init__(self, root):
        self.root = root
        self.root.title("考试成绩分析工具")
        
        self.create_widgets()
        self.filepaths = []

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

        # 存储每个文件的考试编号和年级排名
        exam_data = {}
        exam_numbers = []

        # 读取文件，提取考试编号和年级排名
        for file in self.filepaths:
            df = pd.read_excel(file)
            if '考试编号' in df.columns:
                exam_number = df['考试编号'].max()  # 获取最大考试编号
                exam_numbers.append(exam_number)
                for _, row in df.iterrows():
                    student = row['同学']
                    rank = row['年级排名']
                    
                    # 将年级排名保存到字典中，以学生姓名为键
                    if student not in exam_data:
                        exam_data[student] = {}
                    exam_data[student][exam_number] = rank

        # 创建新的 DataFrame
        progress_data = []
        students = exam_data.keys()

        self.progress_bar['maximum'] = len(students)
        for student in students:
            student_ranks = exam_data[student]
            
            # 按考试编号排序
            sorted_ranks = sorted(student_ranks.items())
            
            # 构建行数据，包括所有的年级排名
            progress_entry = {'学生姓名': student}
            for exam_no, rank in sorted_ranks:
                progress_entry[f'第{exam_no}次考试'] = rank
            
            # 计算进退步系数
            if len(sorted_ranks) >= 2:
                last_exam_rank = sorted_ranks[-2][1]  # 倒数第二个即为上次考试
                current_exam_rank = sorted_ranks[-1][1]  # 最新考试
                progress_coefficient = (last_exam_rank - current_exam_rank) / last_exam_rank
                
                # 根据进退步系数的值确定标识符
                if progress_coefficient > 1:
                    marked_coefficient = f"🟩{progress_coefficient}"
                elif -1 < progress_coefficient < 1:
                    marked_coefficient = f"🟦{progress_coefficient}"
                else:
                    marked_coefficient = f"🟥{progress_coefficient}"
                progress_entry['进退步系数'] = marked_coefficient
            
            progress_data.append(progress_entry)

            self.progress_bar['value'] += 1
            self.root.update_idletasks()  # 更新进度条显示

        progress_df = pd.DataFrame(progress_data)
        output_file = "output.xlsx"
        progress_df.to_excel(output_file, index=False)
        messagebox.showinfo("信息", f"进退步系数已保存至 {output_file}")
        
        # 清空进度条
        self.progress_bar['value'] = 0

    def generate_ranking_chart(self):
        if not self.filepaths:
            messagebox.showwarning("警告", "请先选择文件")
            return

        combined_df = pd.DataFrame()

        for file in self.filepaths:
            df = pd.read_excel(file)
            combined_df = pd.concat([combined_df, df], ignore_index=True)

        from matplotlib import rcParams
        rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
        rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

        students = combined_df['同学'].unique()
        self.progress_bar['maximum'] = len(students)  # 设置进度条最大值

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
            plt.savefig(f'{student}_年级排名折线图.pdf')  # 为每个学生生成独立的PDF
            plt.close()

            self.progress_bar['value'] += 1  # 更新进度条
            self.root.update_idletasks()  # 更新进度条显示
    
        messagebox.showinfo("信息", "年级排名折线图已生成.")

        # 清空进度条
        self.progress_bar['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = ExamAnalysisTool(root)
    root.mainloop()