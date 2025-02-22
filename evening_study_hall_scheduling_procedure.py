import random
import re
import tkinter as tk
import tkinter.messagebox
from tkinter import ttk
import openpyxl

teachers = ['教师1', '教师2', '教师3', '教师4', '教师5', '教师6', '教师7', '教师8', '教师9', '教师10']


class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("晚自习安排")
        self.schedule_tables = []
        self.schedule = {}
        self.section_settings = {}

        light_blue = "#ADD8E6"

        self.style = ttk.Style()
        self.style.configure('Custom.TButton', background=light_blue, width=10, height=2)

        button_frame = ttk.Frame(root)
        button_frame.pack(side=tk.BOTTOM, pady=10, anchor=tk.CENTER)

        self.generate_button = ttk.Button(
            button_frame, text="生成", command=self.show_all_schedules, style='Custom.TButton')
        self.generate_button.pack(side=tk.LEFT, padx=5)

        self.settings_button = ttk.Button(
            button_frame, text="课时设置", command=self.open_settings, style='Custom.TButton')
        self.settings_button.pack(side=tk.LEFT, padx=5)

        self.upload_button = ttk.Button(
            button_frame, text="上传名单", command=self.upload_teachers, style='Custom.TButton')
        self.upload_button.pack(side=tk.LEFT, padx=5)

        self.export_button = ttk.Button(
            button_frame, text="导出", command=self.export_all_schedules, style='Custom.TButton')
        self.export_button.pack(side=tk.LEFT, padx=5)

        self.new_table_button = ttk.Button(
            button_frame, text="新建表", command=self.create_new_table, style='Custom.TButton')
        self.new_table_button.pack(side=tk.LEFT, padx=5)

        self.canvas = tk.Canvas(root)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.style.configure('Treeview', font=('等线', 12))
        self.style.configure('Treeview.Heading', font=('等线', 12, 'bold'))

        self.create_new_table()

    def create_new_table(self):
        table_frame = ttk.Frame(self.scrollable_frame)
        table_frame.pack(pady=10)

        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        columns = ['课时/日期']
        for day in ['周日', '周一', '周二', '周三', '周四']:
            columns.append(day)

        table = ttk.Treeview(table_frame, columns=columns,
                             show='headings', yscrollcommand=scrollbar.set)
        scrollbar.config(command=table.yview)
        table.heading('课时/日期', text='课时/日期')
        for col in columns[1:]:
            table.heading(col, text=col)

        table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.schedule_tables.append(table)

    def open_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("课时设置")

        self.section_settings = {}
        for index, day in enumerate(['周日', '周一', '周二', '周三', '周四']):
            tk.Label(settings_window, text=day).grid(row=0, column=index)
            section_var = tk.IntVar()
            section_spinbox = tk.Spinbox(
                settings_window, from_=1, to=5, textvariable=section_var)
            section_spinbox.grid(row=1, column=index)
            self.section_settings[day] = section_var

        save_button = ttk.Button(
            settings_window, text="保存", command=lambda: self.save_settings(settings_window), style='Custom.TButton')
        save_button.grid(row=2, columnspan=5)

    def save_settings(self, settings_window):
        try:
            for day, var in self.section_settings.items():
                self.schedule[day] = []
                for i in range(var.get()):
                    self.schedule[day].append("未安排")
            settings_window.destroy()
        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))

    def show_all_schedules(self):
        if not self.schedule:
            tkinter.messagebox.showwarning("警告", "请先设置课时")
            return

        global_assigned_teachers_per_day = {day: [] for day in ['周日', '周一', '周二', '周三', '周四']}

        for table in self.schedule_tables:
            for item in table.get_children():
                table.delete(item)

            available_teachers = teachers[:]
            if not available_teachers:
                tkinter.messagebox.showerror("错误", "没有可用的教师名单，请上传教师名单")
                return

            max_sections = max([len(self.schedule[day]) for day in self.schedule.keys()])
            for section in range(1, max_sections + 1):
                values = ['第' + str(section) + '节']
                for day in ['周日', '周一', '周二', '周三', '周四']:
                    if day in self.schedule and len(self.schedule[day]) >= section:
                        if section == 1:
                            available_for_day = [teacher for teacher in available_teachers if teacher not in global_assigned_teachers_per_day[day]]
                            if available_for_day:
                                teacher = random.choice(available_for_day)
                                available_teachers.remove(teacher)
                                global_assigned_teachers_per_day[day].append(teacher)
                                for i in range(section, len(self.schedule[day]) + 1):
                                    self.schedule[day][i - 1] = teacher
                            else:
                                teacher = "无可用教师"
                                for i in range(section, len(self.schedule[day]) + 1):
                                    self.schedule[day][i - 1] = teacher
                        else:
                            teacher = self.schedule[day][section - 1]
                    else:
                        teacher = ""
                    values.append(teacher)
                table.insert('', 'end', values=values)

    def upload_teachers(self):
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="选择教师名单文件", filetypes=[("Text files", "*.txt")])
        if file_path:
            try:
                with open(file_path, 'r', encoding='UTF-8') as file:
                    lines = file.readlines()
                    new_teachers = []
                    for line in lines:
                        teacher_names = re.split(r'[，、;,.\s]\s*', line.strip())
                        for teacher_name in teacher_names:
                            if teacher_name:
                                new_teachers.append(teacher_name)
                    global teachers
                    teachers = new_teachers
                    tkinter.messagebox.showinfo("成功", "教师名单上传成功！")
            except Exception as e:
                tkinter.messagebox.showerror("错误", str(e))

    def export_all_schedules(self):
        try:
            wb = openpyxl.Workbook()
            for i, table in enumerate(self.schedule_tables):
                if i == 0:
                    sheet = wb.active
                else:
                    sheet = wb.create_sheet()
                sheet.title = f"晚自习安排表{i + 1}"

                columns = [table.heading(col)['text'] for col in table['columns']]
                sheet.append(columns)
                for row_id in table.get_children():
                    row = table.item(row_id)['values']
                    sheet.append(row)

            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                wb.save(file_path)
                tkinter.messagebox.showinfo("成功", "晚自习安排已成功导出！")
        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


root = tk.Tk()
app = ScheduleApp(root)
root.mainloop()