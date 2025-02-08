import tkinter as tk  # 所需库：tkinter,tkinter.messagebox,random,os,re,openpyxl
from tkinter import ttk
import tkinter.messagebox
import random
import os
import re
import openpyxl


teachers = ['教师1', '教师2', '教师3', '教师4', '教师5']


class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("晚自习安排")
        self.frame = ttk.Frame(root)
        self.frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
        self.schedule = {}
        self.section_settings = {}

        self.generate_button = tk.Button(
            root, text="生成", command=self.show_schedule)
        self.generate_button.pack()

        self.settings_button = tk.Button(
            root, text="课时设置", command=self.open_settings)
        self.settings_button.pack()

        self.upload_button = tk.Button(
            root, text="上传名单", command=self.upload_teachers)
        self.upload_button.pack()

        self.export_button = tk.Button(
            root, text="导出", command=self.export_schedule)
        self.export_button.pack()

        self.style = ttk.Style()
        self.style.configure('Treeview', font=('等线', 12))
        self.style.configure('Treeview.Heading', font=('等线', 12, 'bold'))

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

        save_button = tk.Button(
            settings_window, text="保存", command=lambda: self.save_settings(settings_window))
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

    def show_schedule(self):
        if not self.schedule:
            tkinter.messagebox.showwarning("警告", "请先设置课时")
            return

        available_teachers = teachers[:]
        if not available_teachers:
            tkinter.messagebox.showerror("错误", "没有可用的教师名单，请上传教师名单")
            return
        for day, sections in self.schedule.items():
            teacher = random.choice(
                [t for t in available_teachers if t not in self.schedule.values()])
            for i in range(len(sections)):
                self.schedule[day][i] = teacher
            available_teachers.remove(teacher)

        for widget in self.frame.winfo_children():
            widget.destroy()
        scrollbar = ttk.Scrollbar(self.frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        columns = ['课时/日期']
        for day in ['周日', '周一', '周二', '周三', '周四']:
            columns.append(day)

        table = ttk.Treeview(self.frame, columns=columns,
                             show='headings', yscrollcommand=scrollbar.set)
        scrollbar.config(command=table.yview)
        table.heading('课时/日期', text='课时/日期')
        for col in columns[1:]:
            table.heading(col, text=col)

        max_sections = max([len(self.schedule[day])
                           for day in self.schedule.keys()])
        for section in range(1, max_sections + 1):
            values = ['第' + str(section) + '节']
            for day in ['周日', '周一', '周二', '周三', '周四']:
                if day in self.schedule and len(self.schedule[day]) >= section:
                    teacher = self.schedule[day][section - 1]
                else:
                    teacher = ""
                values.append(teacher)
            table.insert('', 'end', values=values)

        table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

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

    def export_schedule(self):
        try:
            wb = openpyxl.Workbook()
            sheet = wb.active
            sheet.title = "晚自习安排"

            columns = []
            for widget in self.frame.winfo_children():
                if isinstance(widget, ttk.Treeview):
                    columns = [widget.heading(col)['text']
                               for col in widget['columns']]
                    sheet.append(columns)
                    for row_id in widget.get_children():
                        row = widget.item(row_id)['values']
                        sheet.append(row)

            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                wb.save(file_path)
                tkinter.messagebox.showinfo("成功", "晚自习安排已成功导出！")
        except Exception as e:
            tkinter.messagebox.showerror("错误", str(e))


root = tk.Tk()
app = ScheduleApp(root)
root.mainloop()