import random
import re
import tkinter as tk
import tkinter.messagebox
from tkinter import ttk
import openpyxl

teachers = ['教师1', '教师2', '教师3', '教师4', '教师5', '教师6', '教师7', '教师8', '教师9', '教师10','教师11', '教师12', '教师13', '教师14', '教师15', '教师16', '教师17', '教师18', '教师19', '教师20','教师21','教师22','教师23','教师24','教师25','教师26','教师27','教师28','教师29','教师30','教师31','教师32','教师33','教师34','教师35','教师36','教师37','教师38','教师39','教师40','教师41','教师42','教师43','教师44','教师45','教师46','教师47','教师48','教师49','教师50']


class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.schedule_tables = []
        self.schedule = {}
        self.section_settings = {}
        self.current_language = "简体中文"
        self.languages = self.init_language_packs()

        light_blue = "#ADD8E6"
        self.style = ttk.Style()
        self.style.configure('Custom.TButton', background=light_blue, width=10, height=2)

        button_frame = ttk.Frame(root)
        button_frame.pack(side=tk.BOTTOM, pady=10, anchor=tk.CENTER)

        self.generate_button = ttk.Button(
            button_frame, text="", command=self.show_all_schedules, style='Custom.TButton')
        self.settings_button = ttk.Button(
            button_frame, text="", command=self.open_settings, style='Custom.TButton')
        self.upload_button = ttk.Button(
            button_frame, text="", command=self.upload_teachers, style='Custom.TButton')
        self.export_button = ttk.Button(
            button_frame, text="", command=self.export_all_schedules, style='Custom.TButton')
        self.new_table_button = ttk.Button(
            button_frame, text="", command=self.create_new_table, style='Custom.TButton')
        self.language_button = ttk.Button(
            button_frame, text="Language", command=self.show_language_menu, style='Custom.TButton')

        for btn in [self.generate_button, self.settings_button, self.upload_button,
                    self.export_button, self.new_table_button, self.language_button]:
            btn.pack(side=tk.LEFT, padx=5)

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

        self.update_ui_text()
        self.create_new_table()

    def init_language_packs(self):
        return {
            "简体中文": {
                "window_title": "晚自习安排",
                "generate": "生成",
                "settings": "课时设置",
                "upload": "上传名单",
                "export": "导出",
                "new_table": "新建表",
                "days": ["周日", "周一", "周二", "周三", "周四"],
                "section_header": "课时/日期",
                "settings_title": "课时设置",
                "no_teacher": "无可用教师",
                "save": "保存",
                "warnings": {
                    "no_settings": "请先设置课时",
                    "no_teachers": "没有可用的教师名单，请上传教师名单"
                },
                "success": {
                    "upload": "教师名单上传成功！",
                    "export": "晚自习安排已成功导出！"
                },
                "section_format": "第{}节",
                "unassigned": "未安排",
                "file_dialog": {
                    "open": "选择教师名单文件",
                    "save": "保存安排表"
                },
                "error": "错误",
                "warning": "警告"
            },
            "繁体中文": {
                "window_title": "晚自習安排",
                "generate": "生成",
                "settings": "課時設置",
                "upload": "上傳名單",
                "export": "導出",
                "new_table": "新建表",
                "days": ["週日", "週一", "週二", "週三", "週四"],
                "section_header": "課時/日期",
                "settings_title": "課時設置",
                "no_teacher": "無可用教師",
                "save": "保存",
                "warnings": {
                    "no_settings": "請先設置課時",
                    "no_teachers": "沒有可用的教師名單，請上傳教師名單"
                },
                "success": {
                    "upload": "教師名單上傳成功！",
                    "export": "晚自習安排已成功導出！"
                },
                "section_format": "第{}節",
                "unassigned": "未安排",
                "file_dialog": {
                    "open": "選擇教師名單文件",
                    "save": "保存安排表"
                },
                "error": "錯誤",
                "warning": "警告"
            },
            "English": {
                "window_title": "Evening Study Schedule",
                "generate": "Generate",
                "settings": "Sections Setup",
                "upload": "Upload List",
                "export": "Export",
                "new_table": "New Table",
                "days": ["Sun", "Mon", "Tue", "Wed", "Thu"],
                "section_header": "Section/Date",
                "settings_title": "Section Settings",
                "no_teacher": "No Available Teacher",
                "save": "Save",
                "warnings": {
                    "no_settings": "Please set up the class sections first.",
                    "no_teachers": "There is no available teacher list. Please upload the teacher list."
                },
                "success": {
                    "upload": "Teacher list uploaded successfully!",
                    "export": "Evening study schedule exported successfully!"
                },
                "section_format": "Section {}",
                "unassigned": "Unassigned",
                "file_dialog": {
                    "open": "Select teacher list file",
                    "save": "Save schedule"
                },
                "error": "Error",
                "warning": "Warning"
            },
            "Русский": {
                "window_title": "Вечернее расписание",
                "generate": "Создать",
                "settings": "Настройки занятий",
                "upload": "Загрузить список",
                "export": "Экспорт",
                "new_table": "Новая таблица",
                "days": ["Вс", "Пн", "Вт", "Ср", "Чт"],
                "section_header": "Урок/Дата",
                "settings_title": "Настройки уроков",
                "no_teacher": "Нет доступных учителей",
                "save": "Сохранить",
                "warnings": {
                    "no_settings": "Пожалуйста, сначала настройте уроки.",
                    "no_teachers": "Нет доступного списка учителей. Пожалуйста, загрузите список учителей."
                },
                "success": {
                    "upload": "Список учителей успешно загружен!",
                    "export": "Вечернее расписание успешно экспортировано!"
                },
                "section_format": "Урок {}",
                "unassigned": "Не назначено",
                "file_dialog": {
                    "open": "Выберите файл списка учителей",
                    "save": "Сохранить расписание"
                },
                "error": "Ошибка",
                "warning": "Предупреждение"
            },
            "日本語": {
                "window_title": "自習室スケジュール",
                "generate": "生成",
                "settings": "時間設定",
                "upload": "リストアップロード",
                "export": "エクスポート",
                "new_table": "新規テーブル",
                "days": ["日曜", "月曜", "火曜", "水曜", "木曜"],
                "section_header": "時間/日付",
                "settings_title": "時間設定",
                "no_teacher": "利用可能な教師がいません",
                "save": "保存",
                "warnings": {
                    "no_settings": "まず時間設定を行ってください。",
                    "no_teachers": "利用可能な教師リストがありません。教師リストをアップロードしてください。"
                },
                "success": {
                    "upload": "教師リストのアップロードに成功しました！",
                    "export": "自習室スケジュールのエクスポートに成功しました！"
                },
                "section_format": "セクション{}",
                "unassigned": "未割り当て",
                "file_dialog": {
                    "open": "教師リストのファイルを選択してください",
                    "save": "スケジュールを保存"
                },
                "error": "エラー",
                "warning": "警告"
            }
        }

    def show_language_menu(self):
        lang_menu = tk.Menu(self.root, tearoff=0)
        for lang in self.languages:
            lang_menu.add_command(label=lang, command=lambda l=lang: self.change_language(l))
        x = self.language_button.winfo_rootx()
        y = self.language_button.winfo_rooty() + self.language_button.winfo_height()
        lang_menu.post(x, y)

    def change_language(self, lang):
        self.current_language = lang
        self.update_ui_text()

    def update_ui_text(self):
        lang = self.languages[self.current_language]
        self.root.title(lang["window_title"])
        self.generate_button.config(text=lang["generate"])
        self.settings_button.config(text=lang["settings"])
        self.upload_button.config(text=lang["upload"])
        self.export_button.config(text=lang["export"])
        self.new_table_button.config(text=lang["new_table"])
        self.language_button.config(text="Language")

        new_columns = [lang["section_header"]] + lang["days"]
        for table in self.schedule_tables:
            table["columns"] = new_columns
            for col in new_columns:
                table.heading(col, text=col)

    def create_new_table(self):
        lang = self.languages[self.current_language]
        table_frame = ttk.Frame(self.scrollable_frame)
        table_frame.pack(pady=10)

        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        columns = [lang["section_header"]] + lang["days"]

        table = ttk.Treeview(table_frame, columns=columns,
                             show='headings', yscrollcommand=scrollbar.set)
        scrollbar.config(command=table.yview)
        table.heading(lang["section_header"], text=lang["section_header"])
        for col in columns[1:]:
            table.heading(col, text=col)

        table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.schedule_tables.append(table)

    def open_settings(self):
        lang = self.languages[self.current_language]
        settings_window = tk.Toplevel(self.root)
        settings_window.title(lang["settings_title"])

        self.section_settings = {}
        for index, day in enumerate(lang["days"]):
            tk.Label(settings_window, text=day).grid(row=0, column=index)
            section_var = tk.IntVar()
            section_spinbox = tk.Spinbox(
                settings_window, from_=1, to=5, textvariable=section_var)
            section_spinbox.grid(row=1, column=index)
            self.section_settings[day] = section_var

        save_button = ttk.Button(
            settings_window, text=lang["save"], command=lambda: self.save_settings(settings_window), style='Custom.TButton')
        save_button.grid(row=2, columnspan=5)

    def save_settings(self, settings_window):
        try:
            lang = self.languages[self.current_language]
            for day, var in self.section_settings.items():
                self.schedule[day] = []
                for i in range(var.get()):
                    self.schedule[day].append(lang["unassigned"])
            settings_window.destroy()
        except Exception as e:
            self.show_error(str(e))

    def show_all_schedules(self):
        lang = self.languages[self.current_language]
        if not self.schedule:
            self.show_warning(lang["warnings"]["no_settings"])
            return

        global_assigned_teachers_per_day = {day: [] for day in lang["days"]}

        for table in self.schedule_tables:
            for item in table.get_children():
                table.delete(item)

            available_teachers = teachers[:]
            if not available_teachers:
                self.show_error(lang["warnings"]["no_teachers"])
                return

            max_sections = max([len(self.schedule[day]) for day in self.schedule.keys()])
            for section in range(1, max_sections + 1):
                values = [lang["section_format"].format(section)]
                for day in lang["days"]:
                    if day in self.schedule and len(self.schedule[day]) >= section:
                        if section == 1:
                            available_for_day = [teacher for teacher in available_teachers if teacher not in
                                                 global_assigned_teachers_per_day[day]]
                            if available_for_day:
                                teacher = random.choice(available_for_day)
                                available_teachers.remove(teacher)
                                global_assigned_teachers_per_day[day].append(teacher)
                                for i in range(section, len(self.schedule[day]) + 1):
                                    self.schedule[day][i - 1] = teacher
                            else:
                                teacher = lang["no_teacher"]
                                for i in range(section, len(self.schedule[day]) + 1):
                                    self.schedule[day][i - 1] = teacher
                        else:
                            teacher = self.schedule[day][section - 1]
                    else:
                        teacher = ""
                    values.append(teacher)
                table.insert('', 'end', values=values)

    def upload_teachers(self):
        lang = self.languages[self.current_language]
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title=lang["file_dialog"]["open"], filetypes=[("Text files", "*.txt")])
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
                self.show_info(lang["success"]["upload"])
            except Exception as e:
                self.show_error(str(e))

    def export_all_schedules(self):
        try:
            lang = self.languages[self.current_language]
            wb = openpyxl.Workbook()
            for i, table in enumerate(self.schedule_tables):
                if i == 0:
                    sheet = wb.active
                else:
                    sheet = wb.create_sheet()
                sheet.title = f"{lang['window_title']}{i + 1}"

                columns = [table.heading(col)['text'] for col in table['columns']]
                sheet.append(columns)
                for row_id in table.get_children():
                    row = table.item(row_id)['values']
                    sheet.append(row)

            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title=lang["file_dialog"]["save"])
            if file_path:
                wb.save(file_path)
                self.show_info(lang["success"]["export"])
        except Exception as e:
            self.show_error(str(e))

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def show_error(self, message):
        lang = self.languages[self.current_language]
        tkinter.messagebox.showerror(lang["error"], message)

    def show_warning(self, message):
        lang = self.languages[self.current_language]
        tkinter.messagebox.showwarning(lang["warning"], message)

    def show_info(self, message):
        lang = self.languages[self.current_language]
        tkinter.messagebox.showinfo(lang["success"], message)


root = tk.Tk()
app = ScheduleApp(root)
root.mainloop()