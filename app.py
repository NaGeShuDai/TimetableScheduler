import json
import os
import queue
import sys
import threading
import traceback
import ctypes
import tkinter as tk
import tkinter.font as tkfont
from pathlib import Path
from tkinter import ttk, filedialog, messagebox
from core import Person, SLOTS, DAYS, DEPARTMENTS, parse_timetable, available, candidates, rule_matches, solve_schedule, export_grid, export_template, time_range, weeks_from_text, total_free_cells, compact_weeks

ROOT = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
DATA = (Path(os.environ.get('LOCALAPPDATA', str(ROOT))) / 'TimetableScheduler' / 'data' / 'workspace.json'
        if (ROOT / 'installed.flag').exists() else ROOT / 'data' / 'workspace.json')
BG, INK, GREEN, ORANGE = '#F4F6F7', '#263238', '#157347', '#B65E00'
if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.tk.call('tk', 'scaling', 4 / 3)
        self.title('空课与排班')
        self.geometry('1320x860')
        self.minsize(1000, 700)
        self.configure(bg=BG)
        self.people, self.rules, self.schedules = [], [], {}
        self.slots = [list(s) for s in SLOTS]
        self.custom_slots = [list(s) for s in SLOTS]
        self.slot_mode = 'default'
        self.default_department = '办公室'
        self.demands = {f'{d}:{s}': 1 for d in range(5) for s in range(5)}
        self.busy = False
        self.fairness_threshold = 100
        self.events = queue.Queue()
        self.load()
        font_path = ROOT / 'resources' / 'FangSong_GB2312.ttf'
        if not font_path.exists() and getattr(sys, 'frozen', False):
            font_path = Path(sys._MEIPASS) / 'resources' / 'FangSong_GB2312.ttf'
        if sys.platform == 'win32' and font_path.exists():
            ctypes.windll.gdi32.AddFontResourceExW(str(font_path), 0x10, 0)
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('.', font=('Microsoft YaHei UI', 10))
        style.configure('TFrame', background=BG)
        style.configure('TLabel', background=BG, foreground=INK)
        style.configure('TButton', padding=(10, 6))
        style.configure('Treeview', rowheight=30, background='white', fieldbackground='white')
        style.configure('Treeview.Heading', padding=7)
        style.configure('TNotebook.Tab', padding=(24, 10))
        header = ttk.Frame(self, padding=(18, 12))
        header.pack(fill='x')
        ttk.Label(header, text='空课与排班', font=('Microsoft YaHei UI', 19, 'bold')).pack(side='left')
        ttk.Label(header, text='教学周').pack(side='left', padx=(32, 6))
        self.week = tk.IntVar(value=1)
        spin = ttk.Spinbox(header, from_=1, to=60, width=5, textvariable=self.week, command=self.refresh_all)
        spin.pack(side='left')
        spin.bind('<Return>', lambda e: self.refresh_all())
        spin.bind('<FocusOut>', lambda e: self.refresh_all())
        ttk.Button(header, text='时段设置', command=self.edit_times).pack(side='right')
        self.term_label = ttk.Label(header, text=self.term())
        self.term_label.pack(side='right', padx=16)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=16)
        self.person_page, self.free_page, self.schedule_page = [ttk.Frame(self.notebook) for _ in range(3)]
        for page, title in [(self.person_page, '人员与课表'), (self.free_page, '空课表'), (self.schedule_page, '排班表')]:
            self.notebook.add(page, text=title)
        self.build_people()
        self.build_free()
        self.build_schedule()
        self.status = tk.StringVar(value='就绪。请批量导入课表。')
        ttk.Label(self, textvariable=self.status, padding=(18, 10)).pack(fill='x')
        self.protocol('WM_DELETE_WINDOW', self.close)
        self.after(100, self.poll)
        self.refresh_all()

    def term(self):
        return next((p.term for p in self.people if p.term), '未导入学期')

    def current_week(self):
        try:
            w = self.week.get()
            if not 1 <= w <= 60:
                raise ValueError()
            return w
        except (ValueError, tk.TclError):
            self.week.set(1)
            return 1

    def load(self):
        if DATA.exists():
            try:
                data = json.loads(DATA.read_text(encoding='utf-8'))
                self.people = [Person(**p) for p in data.get('people', [])]
                for p in self.people:
                    p.notes = []
                    p.notes_confirmed = True
                self.rules = data.get('rules', [])
                self.schedules = data.get('schedules', {})
                self.slots = data.get('slots', self.slots)
                self.custom_slots = data.get('custom_slots', self.slots)
                self.slot_mode = data.get('slot_mode', 'default' if self.slots == [list(s) for s in SLOTS] else 'custom')
                self.default_department = data.get('default_department', '办公室')
                self.demands = data.get('demands', self.demands)
                self.fairness_threshold = data.get('fairness_threshold', 100)
            except Exception as exc:
                messagebox.showwarning('数据读取失败', f'已有数据未被覆盖。\n{exc}')
                self.destroy()
                raise

    def save(self):
        DATA.parent.mkdir(parents=True, exist_ok=True)
        tmp = DATA.with_suffix('.tmp')
        tmp.write_text(json.dumps(dict(people=[p.to_dict() for p in self.people], rules=self.rules,
            schedules=self.schedules, slots=self.slots, custom_slots=self.custom_slots, slot_mode=self.slot_mode, demands=self.demands, default_department=self.default_department,
            fairness_threshold=self.threshold.get() if hasattr(self, 'threshold') else self.fairness_threshold), ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(tmp, DATA)

    def persist(self):
        try:
            self.save()
        except OSError as exc:
            messagebox.showerror('保存失败', str(exc))

    def close(self):
        if self.busy:
            messagebox.showinfo('正在计算', '请等待当前排班计算完成后关闭。')
            return
        try:
            self.save()
        except OSError as exc:
            messagebox.showerror('保存失败', str(exc))
            return
        self.destroy()

    def guarded(self):
        if self.busy:
            messagebox.showinfo('正在计算', '排班计算期间，请等待计算完成后再修改资料和规则。')
            return True
        return False

    def invalidate(self):
        for schedule in self.schedules.values():
            schedule['stale'] = True
        self.persist()
        self.refresh_all()

    def build_people(self):
        toolbar = ttk.Frame(self.person_page, padding=(0, 12))
        toolbar.pack(fill='x')
        for title, command in [('批量导入', self.import_people), ('查看课表 / 确认异常', self.person_details), ('增加课程', self.add_course),
                               ('参与 / 停用', self.toggle_person), ('移除人员', self.remove_person)]:
            ttk.Button(toolbar, text=title, command=command).pack(side='left', padx=(0, 8))
        self.person_tree = ttk.Treeview(self.person_page, columns=('name', 'id', 'term', 'group', 'courses', 'state'), show='headings')
        for key, title, width in [('name', '姓名', 130), ('id', '人员编号', 140), ('term', '学期', 130),
                                  ('group', '班级', 260), ('courses', '课程段数', 90), ('state', '状态', 220)]:
            self.person_tree.heading(key, text=title)
            self.person_tree.column(key, width=width)
        self.person_tree.pack(fill='both', expand=True)
        department_bar = ttk.Frame(self.person_page, padding=(0, 8))
        department_bar.pack(fill='x', before=self.person_tree)
        ttk.Label(department_bar, text='默认导入部门').pack(side='left')
        self.import_department = tk.StringVar(value=self.default_department)
        box = ttk.Combobox(department_bar, values=DEPARTMENTS, textvariable=self.import_department, state='readonly', width=16)
        box.pack(side='left', padx=8)
        box.bind('<<ComboboxSelected>>', lambda e: self.set_default_department())
        ttk.Button(department_bar, text='将选中人员改为此部门', command=self.bulk_department).pack(side='left')
        self.person_tree.configure(selectmode='extended')
        self.person_tree['columns'] = ('name', 'id', 'term', 'group', 'courses', 'state', 'department')
        self.person_tree.heading('department', text='部门')
        self.person_tree.column('department', width=110)
        self.person_tree.bind('<Double-1>', lambda e: self.person_details())
        self.person_tree.tag_configure('warning', foreground=ORANGE)

    def selected_person(self):
        selection = self.person_tree.selection()
        if not selection:
            messagebox.showinfo('选择人员', '请先选择一位人员。')
            return None
        return next(p for p in self.people if p.id == selection[0])

    def set_default_department(self):
        self.default_department = self.import_department.get()
        self.persist()

    def bulk_department(self):
        if self.guarded():
            return
        selection = set(self.person_tree.selection())
        if not selection:
            messagebox.showinfo('选择人员', '请先选择人员，可使用 Ctrl 或 Shift 多选。')
            return
        for p in self.people:
            if p.id in selection:
                p.department = self.import_department.get()
        self.invalidate()
        self.status.set(f'已更新 {len(selection)} 人的部门为 {self.import_department.get()}。')

    def import_people(self):
        if self.guarded():
            return
        paths = filedialog.askopenfilenames(title='批量导入个人课表', filetypes=[('Excel 课表', '*.xls *.xlsx')])
        success, failures = 0, []
        for path in paths:
            try:
                p = parse_timetable(path)
                p.department = self.default_department
                if self.people and p.term != self.term():
                    raise ValueError(f'学期 {p.term or "未知"} 与当前学期 {self.term()} 不一致')
                old = next((x for x in self.people if x.id == p.id), None)
                if old:
                    if not messagebox.askyesno('重复人员', f'{p.name}（{p.id}）已存在，是否替换课表？'):
                        continue
                    self.people.remove(old)
                self.people.append(p)
                success += 1
            except Exception as exc:
                failures.append(Path(path).name + '：' + str(exc))
        if paths:
            self.invalidate()
            self.status.set(f'导入完成：成功 {success} 份，失败 {len(failures)} 份。')
        if failures:
            messagebox.showwarning('部分文件导入失败', '\n\n'.join(failures))

    def person_details(self):
        p = self.selected_person()
        if not p:
            return
        dialog = tk.Toplevel(self)
        dialog.title(p.name + ' · 课表与导入信息')
        dialog.geometry('1000x650')
        body = ttk.Frame(dialog, padding=16)
        body.pack(fill='both', expand=True)
        ttk.Label(body, text=f'{p.name}  |  {p.id}  |  {p.term}', font=('Microsoft YaHei UI', 14, 'bold')).pack(anchor='w')
        assignment = ttk.Frame(body)
        assignment.pack(fill='x', pady=8)
        ttk.Label(assignment, text='校区').pack(side='left')
        campus = ttk.Combobox(assignment, values=['海甸', '观澜湖'], state='readonly', width=12)
        campus.set(p.campus)
        campus.pack(side='left', padx=8)
        ttk.Label(assignment, text='部门').pack(side='left')
        department = ttk.Combobox(assignment, values=DEPARTMENTS, state='readonly', width=16)
        department.set(p.department)
        department.pack(side='left', padx=8)
        ttk.Label(body, text='晚间第 9～10 节暂按整行 19:20–21:55 占用；可在时段设置中查看时间。').pack(anchor='w', pady=8)
        tree = ttk.Treeview(body, columns=('course', 'day', 'time', 'weeks', 'place'), show='headings', height=10)
        for key, title, width in [('course', '课程', 250), ('day', '星期', 80), ('time', '时间', 120), ('weeks', '周次', 230), ('place', '地点', 160)]:
            tree.heading(key, text=title)
            tree.column(key, width=width)
        tree.pack(fill='both', expand=True)
        for c in p.courses:
            clock = f'{c["start"] // 60:02}:{c["start"] % 60:02}-{c["end"] // 60:02}:{c["end"] % 60:02}'
            tree.insert('', 'end', values=(c['name'], DAYS[c['day']], clock, ','.join(map(str, c['weeks'])), c['location']))
        confirmed = tk.BooleanVar(value=p.errors_confirmed)
        if p.errors:
            notes = tk.Text(body, height=5, wrap='word', font=('Microsoft YaHei UI', 10))
            notes.pack(fill='x', pady=12)
            notes.insert('1.0', '解析异常：\n' + '\n'.join(p.errors))
            notes.configure(state='disabled')
            ttk.Checkbutton(body, text='已核实异常并补齐遗漏课程，允许参与计算', variable=confirmed).pack(anchor='w')
        def commit():
            if self.guarded():
                return
            p.errors_confirmed = confirmed.get()
            p.campus, p.department = campus.get(), department.get()
            self.invalidate()
            dialog.destroy()
        ttk.Button(body, text='保存', command=commit).pack(anchor='e', pady=8)

    def add_course(self):
        if self.guarded():
            return
        p = self.selected_person()
        if not p:
            return
        dialog = tk.Toplevel(self)
        dialog.title(p.name + ' · 增加课程')
        body = ttk.Frame(dialog, padding=20)
        body.pack()
        fields = {}
        for i, (key, label, default) in enumerate([('name', '课程名称', ''), ('weeks', '周次', '1-18'),
                                                  ('time', '时间', '07:40-09:20'), ('location', '地点', '')]):
            ttk.Label(body, text=label).grid(row=i, column=0, padx=8, pady=8)
            entry = ttk.Entry(body, width=36)
            entry.insert(0, default)
            entry.grid(row=i, column=1)
            fields[key] = entry
        ttk.Label(body, text='星期').grid(row=4, column=0, pady=8)
        day = ttk.Combobox(body, values=DAYS, state='readonly', width=34)
        day.current(0)
        day.grid(row=4, column=1)
        def commit():
            if self.guarded():
                return
            try:
                name = fields['name'].get().strip()
                if not name:
                    raise ValueError('请输入课程名称')
                weeks = weeks_from_text(fields['weeks'].get())
                start, end = time_range(fields['time'].get().strip())
                p.courses.append(dict(name=name, teacher='', weeks=weeks, day=day.current(),
                    periods=[], start=start, end=end, location=fields['location'].get().strip(),
                    raw='', cell='手动', conservative=False, manual=True))
                self.invalidate()
                dialog.destroy()
            except ValueError as exc:
                messagebox.showerror('课程无效', str(exc), parent=dialog)
        ttk.Button(body, text='增加课程', command=commit).grid(row=5, column=1, sticky='e', pady=12)

    def toggle_person(self):
        if self.guarded():
            return
        p = self.selected_person()
        if p:
            p.enabled = not p.enabled
            self.invalidate()

    def remove_person(self):
        if self.guarded():
            return
        p = self.selected_person()
        if p and messagebox.askyesno('移除人员', '确定移除 ' + p.name + ' 的课表及关联规则？'):
            self.people.remove(p)
            self.rules = [r for r in self.rules if r['person'] != p.id]
            self.invalidate()

    def grid_container(self, page):
        frame = ttk.Frame(page)
        frame.pack(fill='both', expand=True)
        canvas = tk.Canvas(frame, bg='white', highlightthickness=0)
        xscroll, yscroll = ttk.Scrollbar(frame, orient='horizontal', command=canvas.xview), ttk.Scrollbar(frame, command=canvas.yview)
        canvas.configure(xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
        canvas.grid(row=0, column=0, sticky='nsew')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        inner = tk.Frame(canvas, bg='white')
        window = canvas.create_window(0, 0, window=inner, anchor='nw')
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>', lambda e: canvas.itemconfigure(window, width=max(e.width, 800)))
        def wheel(event):
            canvas.yview_scroll(-int(event.delta / 120), 'units')
        canvas.bind('<MouseWheel>', wheel)
        return inner

    def build_free(self):
        toolbar = ttk.Frame(self.free_page, padding=(0, 12))
        toolbar.pack(fill='x')
        ttk.Button(toolbar, text='刷新', command=self.refresh_all).pack(side='left')
        ttk.Button(toolbar, text='导出空课表', command=lambda: self.export(False)).pack(side='left', padx=8)
        self.free_mode = tk.StringVar(value='本周空课表')
        mode = ttk.Combobox(toolbar, values=['本周空课表', '总空课表'], textvariable=self.free_mode, state='readonly', width=12)
        mode.pack(side='left', padx=6)
        mode.bind('<<ComboboxSelected>>', lambda e: self.refresh_free())
        self.total_weeks = tk.StringVar(value='1-18')
        ttk.Label(toolbar, text='汇总周次').pack(side='left', padx=(8, 4))
        entry = ttk.Entry(toolbar, textvariable=self.total_weeks, width=16)
        entry.pack(side='left')
        entry.bind('<Return>', lambda e: self.refresh_free())
        entry.bind('<FocusOut>', lambda e: self.refresh_free())
        self.free_summary = ttk.Label(toolbar)
        self.free_summary.pack(side='right')
        self.free_grid = self.grid_container(self.free_page)

    def build_schedule(self):
        toolbar = ttk.Frame(self.schedule_page, padding=(0, 12))
        toolbar.pack(fill='x')
        self.generate_button = ttk.Button(toolbar, text='生成建议排班', command=self.generate)
        self.generate_button.pack(side='left')
        for title, command in [('班次人数', self.edit_demands), ('撤销', lambda: self.history(False)),
                               ('重做', lambda: self.history(True)), ('恢复建议', self.restore), ('导出排班', lambda: self.export(True))]:
            ttk.Button(toolbar, text=title, command=command).pack(side='left', padx=(6, 0))
        ttk.Button(toolbar, text='规则面板 ▸', command=self.toggle_rules).pack(side='right')
        self.schedule_note = ttk.Label(self.schedule_page, foreground=ORANGE, padding=(0, 0, 0, 8))
        self.schedule_note.pack(fill='x')
        self.split = ttk.Panedwindow(self.schedule_page, orient='horizontal')
        self.split.pack(fill='both', expand=True)
        left = ttk.Frame(self.split)
        self.split.add(left, weight=4)
        self.schedule_grid = self.grid_container(left)
        self.stats = ttk.Treeview(left, columns=('name', 'count', 'hours', 'state'), show='headings', height=4)
        for key, title in [('name', '人员'), ('count', '班次'), ('hours', '总时长'), ('state', '负担提醒')]:
            self.stats.heading(key, text=title)
            self.stats.column(key, width=120)
        self.stats.tag_configure('unfair', foreground=ORANGE)
        self.stats.pack(fill='x', pady=(8, 0))
        self.rule_panel = ttk.Frame(self.split, width=320, padding=12)
        self.rules_visible = True
        self.split.add(self.rule_panel, weight=1)
        rule_canvas = tk.Canvas(self.rule_panel, width=320, bg=BG, highlightthickness=0)
        rule_scroll = ttk.Scrollbar(self.rule_panel, command=rule_canvas.yview)
        rule_scroll.pack(side='right', fill='y')
        rule_canvas.pack(side='left', fill='both', expand=True)
        rule_canvas.configure(yscrollcommand=rule_scroll.set)
        rule_body = ttk.Frame(rule_canvas)
        rule_window = rule_canvas.create_window(0, 0, anchor='nw', window=rule_body)
        rule_body.bind('<Configure>', lambda e: rule_canvas.configure(scrollregion=rule_canvas.bbox('all')))
        rule_canvas.bind('<Configure>', lambda e: rule_canvas.itemconfigure(rule_window, width=e.width))
        ttk.Label(rule_body, text='排班规则', font=('Microsoft YaHei UI', 14, 'bold')).pack(anchor='w')
        ttk.Label(rule_body, text='自定义规则 · 从上到下优先', padding=(0, 10)).pack(anchor='w')
        self.rule_list = tk.Listbox(rule_body, height=7, width=34, font=('Microsoft YaHei UI', 10), exportselection=False,
                                    selectbackground='#D9EDE5', selectforeground=INK)
        self.rule_list.pack(fill='x')
        self.rule_list.bind('<Double-1>', lambda e: self.edit_rule(True))
        row = ttk.Frame(rule_body)
        row.pack(fill='x', pady=8)
        for title, cmd in [('新增', lambda: self.edit_rule(False)), ('编辑', lambda: self.edit_rule(True)), ('删除', self.delete_rule)]:
            ttk.Button(row, text=title, command=cmd, width=6).pack(side='left', padx=2)
        row = ttk.Frame(rule_body)
        row.pack(fill='x')
        for title, cmd in [('↑', lambda: self.move_rule(-1)), ('↓', lambda: self.move_rule(1)), ('启用 / 停用', self.toggle_rule)]:
            ttk.Button(row, text=title, command=cmd, width=9).pack(side='left', padx=2)
        ttk.Separator(rule_body).pack(fill='x', pady=16)
        ttk.Label(rule_body, text='默认规则', font=('Microsoft YaHei UI', 11, 'bold')).pack(anchor='w')
        for text in ['1. 满足班次人数', '2. 均衡总时长', '3. 均衡晚班次数', '4. 均衡周末次数', '5. 减少一天多班']:
            ttk.Label(rule_body, text=text, padding=(0, 5)).pack(anchor='w')
        ttk.Separator(rule_body).pack(fill='x', pady=12)
        ttk.Label(rule_body, text='公平性提醒阈值（超过平均时长）').pack(anchor='w')
        self.threshold = tk.IntVar(value=self.fairness_threshold)
        box = ttk.Spinbox(rule_body, from_=0, to=1000, increment=10, textvariable=self.threshold, width=8,
                          command=self.refresh_schedule)
        box.pack(anchor='w', pady=6)
        box.bind('<Return>', lambda e: self.refresh_schedule())
        box.bind('<FocusOut>', lambda e: self.refresh_schedule())
        ttk.Label(rule_body, text='分钟', padding=(0, 0, 0, 12)).pack(anchor='w')
        for text, color in [('黄色边框：已修改', INK), ('绿色姓名：本次加入', GREEN), ('橙色姓名：负担偏高', ORANGE), ('红色边框：缺员或冲突', '#BD3636')]:
            ttk.Label(rule_body, text=text, foreground=color, padding=(0, 4)).pack(anchor='w')

    def toggle_rules(self):
        if self.rules_visible:
            self.split.forget(self.rule_panel)
        else:
            self.split.add(self.rule_panel, weight=1)
        self.rules_visible = not self.rules_visible

    def refresh_all(self):
        self.current_week()
        self.term_label.configure(text=self.term())
        self.person_tree.delete(*self.person_tree.get_children())
        for p in self.people:
            state = '参与' if p.enabled else '已停用'
            if p.errors and not p.errors_confirmed:
                state = '解析异常，不能参与'
            elif p.errors:
                state = '异常已确认' if p.enabled else '已停用'
            self.person_tree.insert('', 'end', iid=p.id, values=(p.name, p.id, p.term, p.group, len(p.courses), state, p.department),
                                    tags=('warning',) if p.errors and not p.errors_confirmed else ())
        self.refresh_rules()
        self.refresh_free()
        self.refresh_schedule()

    def draw_headers(self, grid):
        for child in grid.winfo_children():
            child.destroy()
        for col, name in enumerate(['节次 / 时间'] + DAYS):
            tk.Label(grid, text=name, bg='#E6EEF0', fg=INK, font=('Microsoft YaHei UI', 10, 'bold'), padx=8, pady=12).grid(row=0, column=col, sticky='nsew', padx=1, pady=1)
            grid.columnconfigure(col, weight=1, minsize=95)
        for s, (label, clock) in enumerate(self.slots):
            tk.Label(grid, text=label + '\n' + clock, bg='#EEF2F3', fg=INK, font=('Microsoft YaHei UI', 10), padx=5, pady=12).grid(row=s + 1, column=0, sticky='nsew', padx=1, pady=1)
            grid.rowconfigure(s + 1, minsize=86)

    def refresh_free(self):
        self.draw_headers(self.free_grid)
        w = self.current_week()
        eligible = sum(p.enabled and (not p.errors or p.errors_confirmed) for p in self.people)
        self.free_summary.configure(text=f'第 {w} 周 · 可参与 {eligible} / {len(self.people)} 人')
        totals = None
        if self.free_mode.get() == '总空课表':
            try:
                weeks = weeks_from_text(self.total_weeks.get())
                totals = total_free_cells(self.people, weeks, self.slots)
                font = tkfont.Font(family='Microsoft YaHei UI', size=10)
                for d in range(7):
                    width = max(font.measure(line) for s in range(len(self.slots)) for line in totals[f'{d}:{s}'].splitlines()) + 20
                    self.free_grid.columnconfigure(d + 1, minsize=max(95, width))
                self.free_summary.configure(text=f'第 {compact_weeks(weeks)} 周 · {eligible} 人')
            except ValueError:
                self.free_summary.configure(text='请输入有效汇总周次，例如 1-18')
                return
        for s in range(len(self.slots)):
            for d in range(7):
                names = [p.name for p in self.people if available(p, w, d, s, self.slots)]
                text = totals[f'{d}:{s}'] if totals is not None else f'{len(names)} 人\n' + ('、'.join(names) or '—')
                label = tk.Label(self.free_grid, text=text, bg='white', fg=GREEN if text != '无' else '#8B9498',
                                 font=('Microsoft YaHei UI', 10), wraplength=0 if totals is not None else 88, padx=6, pady=10)
                label.grid(row=s + 1, column=d + 1, sticky='nsew', padx=1, pady=1)

    def schedule(self):
        return self.schedules.get(str(self.current_week()))

    def loads(self, cells):
        eligible = set()
        for key in self.demands:
            eligible.update(candidates(self.people, self.current_week(), *map(int, key.split(':')), self.rules, self.slots))
        loads = {p.id: 0 for p in self.people if p.id in eligible}
        counts = {pid: 0 for pid in loads}
        for key, ids in cells.items():
            _, slot = map(int, key.split(':'))
            a, b = time_range(self.slots[slot][1])
            for pid in ids:
                if pid in loads:
                    loads[pid] += b - a
                    counts[pid] += 1
        mean = sum(loads.values()) / len(loads) if loads else 0
        try:
            threshold = max(0, self.threshold.get())
        except tk.TclError:
            threshold = 100
        unfair = {pid for pid, load in loads.items() if load > mean + threshold}
        return loads, counts, unfair, mean

    def refresh_schedule(self):
        self.draw_headers(self.schedule_grid)
        state = self.schedule()
        cells = state['cells'] if state else {}
        baseline = state.get('baseline', {}) if state else {}
        last = state.get('last_added', {}) if state else {}
        loads, counts, unfair, mean = self.loads(cells)
        names = {p.id: p.name for p in self.people}
        missing = sum(max(0, count - len(cells.get(k, []))) for k, count in self.demands.items())
        text = '尚未生成本周排班。' if not state else f'第 {self.current_week()} 周 · 缺员 {missing} 个名额 · 人均 {mean / 60:.1f} 小时'
        if state and state.get('stale'):
            text += ' · 资料或规则已变化，请重新生成后编辑与导出。'
        if state and state.get('diagnostics'):
            text += ' · ' + '；'.join(state['diagnostics'])
        self.schedule_note.configure(text=text)
        for s in range(len(self.slots)):
            for d in range(7):
                key = f'{d}:{s}'
                ids = cells.get(key, [])
                count = self.demands.get(key, 0)
                changed = state is not None and ids != baseline.get(key, [])
                bad = count > len(ids) or self.cell_conflict(key, ids, cells)
                border = '#CC4949' if state and bad else '#D3A423' if changed else '#DCE2E4'
                frame = tk.Frame(self.schedule_grid, bg='white' if count else '#F5F6F7', highlightbackground=border, highlightthickness=2, cursor='hand2' if count else '')
                frame.grid(row=s + 1, column=d + 1, sticky='nsew', padx=1, pady=1)
                label = tk.Label(frame, text=f'{len(ids)} / {count} 人' if count else '不排班', bg=frame['bg'], fg='#7A858A', font=('Microsoft YaHei UI', 9))
                label.pack(pady=(8, 3))
                widgets = [frame, label]
                for pid in ids:
                    added = pid in last.get(key, [])
                    lbl = tk.Label(frame, text=('＋' if added else '') + names.get(pid, '已移除人员') + (' · 偏高' if pid in unfair else ''),
                                   bg='#E6F4EB' if added else frame['bg'], fg=ORANGE if pid in unfair else GREEN if added else INK,
                                   font=('Microsoft YaHei UI', 10), wraplength=90)
                    lbl.pack(fill='x', padx=3, pady=2)
                    widgets.append(lbl)
                if not ids and count:
                    empty = tk.Label(frame, text='待安排', bg=frame['bg'], fg='#939CA0')
                    empty.pack(pady=6)
                    widgets.append(empty)
                for widget in widgets:
                    widget.bind('<Button-1>', lambda e, k=key: self.edit_cell(k))
        self.stats.delete(*self.stats.get_children())
        for pid in sorted(loads, key=lambda x: (-loads[x], names[x])):
            self.stats.insert('', 'end', values=(names[pid], counts[pid], f'{loads[pid] / 60:.2f} 小时',
                f'超平均 {loads[pid] - mean:.0f} 分钟' if pid in unfair else ''), tags=('unfair',) if pid in unfair else ())

    def cell_conflict(self, key, ids, cells):
        day, slot = map(int, key.split(':'))
        pool = candidates(self.people, self.current_week(), day, slot, self.rules, self.slots)
        if len(set(ids)) != len(ids) or any(pid not in pool for pid in ids):
            return True
        fixed = {r['person'] for r in self.rules if r['kind'] == 'fixed' and rule_matches(r, self.current_week(), day, slot)}
        if not fixed.issubset(ids):
            return True
        a, b = time_range(self.slots[slot][1])
        for other, assigned in cells.items():
            if other == key:
                continue
            od, oslot = map(int, other.split(':'))
            x, y = time_range(self.slots[oslot][1])
            if day == od and a < y and x < b and set(ids).intersection(assigned):
                return True
        return False

    def refresh_rules(self):
        self.rule_list.delete(0, 'end')
        names = {p.id: p.name for p in self.people}
        kinds = {'exclude': '禁止', 'fixed': '固定', 'prefer': '优先'}
        for i, r in enumerate(self.rules):
            scope = (DAYS[r['day']] if r['day'] is not None else '每天') + ' / ' + (self.slots[r['slot']][0] if r['slot'] is not None else '全部节次')
            weeks = ','.join(map(str, r.get('weeks', []))) or '全部'
            self.rule_list.insert('end', f'{i + 1}. {"✓" if r.get("enabled", True) else "停"} {kinds[r["kind"]]} {names.get(r["person"], "未知")}\n{scope} · 周次 {weeks}')

    def rule_index(self):
        selection = self.rule_list.curselection()
        return selection[0] if selection else None

    def edit_rule(self, existing):
        if self.guarded():
            return
        if not self.people:
            messagebox.showinfo('没有人员', '请先导入课表。')
            return
        index = self.rule_index() if existing else None
        if existing and index is None:
            return
        r = self.rules[index] if index is not None else dict(kind='exclude', person=self.people[0].id, day=None, slot=None, weeks=[], enabled=True)
        dialog = tk.Toplevel(self)
        dialog.title('编辑规则' if existing else '新增规则')
        dialog.resizable(False, False)
        body = ttk.Frame(dialog, padding=20)
        body.pack()
        kinds = ['禁止安排', '固定安排', '优先安排']
        keys = ['exclude', 'fixed', 'prefer']
        person_choices = [f'{p.name}（{p.id}）' for p in self.people]
        values = [kinds[keys.index(r['kind'])], person_choices[next(i for i, p in enumerate(self.people) if p.id == r['person'])],
                  '全部星期' if r['day'] is None else DAYS[r['day']], '全部节次' if r['slot'] is None else self.slots[r['slot']][0]]
        boxes = []
        for i, (label, choices) in enumerate([('类型', kinds), ('人员', person_choices), ('星期', ['全部星期'] + DAYS), ('节次', ['全部节次'] + [s[0] for s in self.slots])]):
            ttk.Label(body, text=label).grid(row=i, column=0, sticky='w', padx=8, pady=8)
            box = ttk.Combobox(body, values=choices, state='readonly', width=32)
            box.set(values[i])
            box.grid(row=i, column=1, pady=8)
            boxes.append(box)
        ttk.Label(body, text='周次（空白为全部）').grid(row=4, column=0, padx=8, pady=8)
        week_text = ttk.Entry(body, width=34)
        week_text.insert(0, ','.join(map(str, r.get('weeks', []))))
        week_text.grid(row=4, column=1)
        enabled = tk.BooleanVar(value=r.get('enabled', True))
        ttk.Checkbutton(body, text='启用规则', variable=enabled).grid(row=5, column=1, sticky='w', pady=10)
        def commit():
            if self.guarded():
                return
            try:
                new = dict(kind=keys[boxes[0].current()], person=self.people[boxes[1].current()].id,
                           day=None if boxes[2].current() == 0 else boxes[2].current() - 1,
                           slot=None if boxes[3].current() == 0 else boxes[3].current() - 1,
                           weeks=weeks_from_text(week_text.get()) if week_text.get().strip() else [], enabled=enabled.get())
                if index is None:
                    self.rules.append(new)
                else:
                    self.rules[index] = new
                self.invalidate()
                dialog.destroy()
            except ValueError as exc:
                messagebox.showerror('规则无效', str(exc), parent=dialog)
        ttk.Button(body, text='保存规则', command=commit).grid(row=6, column=1, sticky='e')

    def delete_rule(self):
        if self.guarded():
            return
        index = self.rule_index()
        if index is not None:
            self.rules.pop(index)
            self.invalidate()

    def move_rule(self, direction):
        if self.guarded():
            return
        i = self.rule_index()
        if i is not None and 0 <= i + direction < len(self.rules):
            self.rules[i], self.rules[i + direction] = self.rules[i + direction], self.rules[i]
            self.invalidate()
            self.rule_list.selection_set(i + direction)

    def toggle_rule(self):
        if self.guarded():
            return
        i = self.rule_index()
        if i is not None:
            self.rules[i]['enabled'] = not self.rules[i].get('enabled', True)
            self.invalidate()

    def generate(self):
        if self.guarded():
            return
        if not self.people:
            messagebox.showinfo('没有课表', '请先导入人员课表。')
            return
        w = self.current_week()
        old = self.schedules.get(str(w))
        if old and not messagebox.askyesno('重新生成', '重新生成会替换本周已有排班及手动修改，是否继续？'):
            return
        self.busy = True
        self.generate_button.configure(state='disabled')
        self.status.set(f'正在计算第 {w} 周建议排班…')
        people = [Person(**p.to_dict()) for p in self.people]
        rules = json.loads(json.dumps(self.rules))
        demands, slots = dict(self.demands), [list(s) for s in self.slots]
        def worker():
            try:
                result, diagnostics = solve_schedule(people, w, demands, rules, slots)
                self.events.put(('result', w, result, diagnostics))
            except Exception as exc:
                self.events.put(('error', str(exc)))
        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            event = self.events.get_nowait()
            self.busy = False
            self.generate_button.configure(state='normal')
            if event[0] == 'error':
                self.status.set('生成失败，请调整规则。')
                messagebox.showerror('无法生成排班', event[1])
            else:
                _, w, cells, diagnostics = event
                self.schedules[str(w)] = dict(cells=cells, baseline=json.loads(json.dumps(cells)), last_added={}, undo=[], redo=[], stale=False, diagnostics=diagnostics)
                self.persist()
                self.refresh_schedule()
                self.status.set(f'第 {w} 周排班已生成。')
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def edit_cell(self, key):
        if self.guarded() or not self.demands.get(key):
            return
        state = self.schedule()
        if not state or state.get('stale'):
            messagebox.showinfo('需要生成排班', '请先生成当前资料和规则下的本周建议排班。')
            return
        w = self.current_week()
        day, slot = map(int, key.split(':'))
        count = self.demands[key]
        ids = state['cells'].get(key, [])
        loads, _, _, _ = self.loads(state['cells'])
        pool = candidates(self.people, w, day, slot, self.rules, self.slots)
        # Exclude people already assigned to an overlapping different cell.
        a, b = time_range(self.slots[slot][1])
        for other, assigned in state['cells'].items():
            od, oslot = map(int, other.split(':'))
            x, y = time_range(self.slots[oslot][1])
            if other != key and day == od and a < y and x < b:
                pool = [pid for pid in pool if pid not in assigned]
        def rank(pid):
            priorities = [i for i, r in enumerate(self.rules) if r['person'] == pid and r['kind'] == 'prefer' and rule_matches(r, w, day, slot)]
            return (min(priorities, default=len(self.rules)), loads.get(pid, 0), pid)
        pool.sort(key=rank)
        names = {p.id: p.name for p in self.people}
        labels = ['（空缺）']
        for pid in pool:
            added = 0 if pid in ids else b - a
            preferred = any(r['person'] == pid and r['kind'] == 'prefer' and rule_matches(r, w, day, slot) for r in self.rules)
            labels.append(f'{names[pid]}（{pid}） · 当前 {loads.get(pid, 0) / 60:.1f}h → { (loads.get(pid, 0) + added) / 60:.1f}h' + (' · 优先' if preferred else ''))
        dialog = tk.Toplevel(self)
        dialog.title(f'第 {w} 周 · {DAYS[day]} · {self.slots[slot][0]}')
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        body = ttk.Frame(dialog, padding=20)
        body.pack()
        ttk.Label(body, text=f'可选 {len(pool)} 人 · 需要 {count} 人').pack(anchor='w', pady=(0, 12))
        boxes = []
        for i in range(count):
            box = ttk.Combobox(body, values=labels, state='readonly', width=66)
            pid = ids[i] if i < len(ids) else None
            box.current(pool.index(pid) + 1 if pid in pool else 0)
            box.pack(pady=6)
            boxes.append(box)
        def commit():
            new = [pool[box.current() - 1] for box in boxes if box.current() > 0]
            tentative = dict(state['cells'])
            tentative[key] = new
            if self.cell_conflict(key, new, tentative):
                messagebox.showerror('安排冲突', '存在重复人员、固定规则未满足或时间冲突。', parent=dialog)
                return
            if new != ids:
                state['undo'].append(json.loads(json.dumps(state['cells'])))
                state['undo'] = state['undo'][-50:]
                state['redo'] = []
                state['cells'][key] = new
                state['last_added'] = {key: [pid for pid in new if pid not in ids]}
                self.persist()
                self.refresh_schedule()
            dialog.destroy()
        ttk.Button(body, text='保存安排', command=commit).pack(anchor='e', pady=(16, 0))
        if boxes:
            boxes[0].focus_set()
            dialog.after(100, lambda: boxes[0].event_generate('<Down>') if boxes[0].winfo_exists() else None)

    def history(self, redo):
        if self.guarded():
            return
        state = self.schedule()
        if not state or state.get('stale'):
            return
        source, target = ('redo', 'undo') if redo else ('undo', 'redo')
        if state[source]:
            old = state['cells']
            state[target].append(json.loads(json.dumps(old)))
            state['cells'] = state[source].pop()
            state['last_added'] = {k: [p for p in v if p not in old.get(k, [])] for k, v in state['cells'].items()}
            self.persist()
            self.refresh_schedule()

    def restore(self):
        if self.guarded():
            return
        state = self.schedule()
        if state and not state.get('stale') and messagebox.askyesno('恢复建议', '恢复到本周最初生成的建议排班？'):
            state['undo'].append(json.loads(json.dumps(state['cells'])))
            state['cells'] = json.loads(json.dumps(state['baseline']))
            state['last_added'], state['redo'] = {}, []
            self.persist()
            self.refresh_schedule()

    def edit_demands(self):
        if self.guarded():
            return
        dialog = tk.Toplevel(self)
        dialog.title('班次人数 · 0 表示不排班')
        body = ttk.Frame(dialog, padding=16)
        body.pack()
        for d, name in enumerate(DAYS):
            ttk.Label(body, text=name).grid(row=0, column=d + 1, padx=8)
        inputs = {}
        for s, (label, _) in enumerate(self.slots):
            ttk.Label(body, text=label).grid(row=s + 1, column=0, padx=8, pady=8)
            for d in range(7):
                key = f'{d}:{s}'
                var = tk.StringVar(value=str(self.demands.get(key, 0)))
                ttk.Spinbox(body, from_=0, to=20, textvariable=var, width=5).grid(row=s + 1, column=d + 1, padx=5)
                inputs[key] = var
        def commit():
            if self.guarded():
                return
            try:
                values = {k: int(v.get()) for k, v in inputs.items()}
                if any(not 0 <= n <= 20 for n in values.values()):
                    raise ValueError()
                self.demands = {k: n for k, n in values.items() if n}
                self.invalidate()
                dialog.destroy()
            except ValueError:
                messagebox.showerror('人数无效', '人数必须是 0～20 的整数。', parent=dialog)
        ttk.Button(body, text='保存', command=commit).grid(row=len(self.slots) + 1, column=7, pady=12)

    def edit_times(self):
        if self.guarded():
            return
        dialog = tk.Toplevel(self)
        dialog.title('排班时段设置')
        body = ttk.Frame(dialog, padding=20)
        body.pack()
        mode = tk.StringVar(value=self.slot_mode)
        ttk.Radiobutton(body, text='默认课表时段', variable=mode, value='default').pack(anchor='w')
        ttk.Radiobutton(body, text='自定义时段', variable=mode, value='custom').pack(anchor='w', pady=8)
        editor = tk.Text(body, width=48, height=14, font=('Microsoft YaHei UI', 10))
        editor.pack(fill='both', expand=True)
        editor.insert('1.0', '\n'.join(f'{label} | {clock}' for label, clock in self.custom_slots))
        def commit():
            if self.guarded():
                return
            try:
                custom = []
                for line in editor.get('1.0', 'end').splitlines():
                    if not line.strip():
                        continue
                    parts = line.split('|')
                    if len(parts) != 2 or not parts[0].strip():
                        raise ValueError('每行格式：时段名称 | HH:MM-HH:MM')
                    label, clock = [part.strip() for part in parts]
                    time_range(clock)
                    custom.append([label, clock])
                if not 1 <= len(custom) <= 24 or len({s[0] for s in custom}) != len(custom):
                    raise ValueError('请设置 1～24 个名称不重复的时段')
                slots = [list(s) for s in SLOTS] if mode.get() == 'default' else custom
                if slots != self.slots:
                    if any(r['slot'] is not None for r in self.rules) and not messagebox.askyesno('时段变更', '时段变更会停用绑定具体节次的规则，并重置班次人数。是否继续？', parent=dialog):
                        return
                    for r in self.rules:
                        if r['slot'] is not None:
                            r['enabled'] = False
                            r['slot'] = None
                    self.demands = {f'{d}:{s}': 1 for d in range(5) for s in range(min(5, len(slots)))}
                    self.schedules = {}
                self.custom_slots, self.slot_mode, self.slots = custom, mode.get(), slots
                self.invalidate()
                dialog.destroy()
            except ValueError as exc:
                messagebox.showerror('时间无效', str(exc), parent=dialog)
        ttk.Button(body, text='保存', command=commit).pack(anchor='e', pady=12)

    def export(self, schedule):
        w = self.current_week()
        state = self.schedule()
        if schedule and (not state or state.get('stale')):
            messagebox.showinfo('不能导出', '请先生成当前规则下的本周排班。')
            return
        if schedule:
            names = {p.id: p.name for p in self.people}
            cells = {k: '\n'.join(names.get(pid, '已移除人员') for pid in ids) for k, ids in state['cells'].items()}
            for key, count in self.demands.items():
                missing = count - len(state['cells'].get(key, []))
                if missing:
                    cells[key] = (cells.get(key, '') + f'\n【缺 {missing} 人】').strip()
        else:
            cells = {f'{d}:{s}': '\n'.join(p.name for p in self.people if available(p, w, d, s, self.slots)) or '无'
                     for d in range(7) for s in range(len(self.slots))}
        title = '建议排班表' if schedule else '空课表'
        total = False
        if not schedule and self.free_mode.get() == '总空课表':
            try:
                weeks = weeks_from_text(self.total_weeks.get())
            except ValueError as exc:
                messagebox.showerror('周次无效', str(exc))
                return
            cells = total_free_cells(self.people, weeks, self.slots)
            w = compact_weeks(weeks)
            title = '总空课表'
            total = True
        filename = f'建议排班表_第{w}周.xlsx' if schedule else f'第{w}周空课表.xlsx'
        if total:
            import re
            match = re.fullmatch(r'(\d{4})-(\d{4})-(\d)', self.term())
            filename = (f'{match[1]}-{match[2]}学年第{match[3]}学期校新媒体空课表.xlsx' if match else '总空课表.xlsx')
        path = filedialog.asksaveasfilename(title='导出 ' + title, initialfile=filename, defaultextension='.xlsx', filetypes=[('Excel 工作簿', '*.xlsx')])
        if path:
            try:
                if schedule:
                    export_grid(path, title, w, cells, self.slots, self.term())
                else:
                    export_template(path, self.people, weeks if total else [self.current_week()], total)
                self.status.set('已导出：' + path)
                messagebox.showinfo('导出完成', '已保存到：\n' + path)
            except Exception as exc:
                messagebox.showerror('导出失败', str(exc))


def main():
    if '--self-test' in sys.argv:
        import tempfile
        from openpyxl import load_workbook
        index = sys.argv.index('--self-test')
        source, report_path = sys.argv[index + 1:index + 3]
        report = {}
        try:
            person = parse_timetable(source)
            assert len(person.courses) == 27 and not person.errors
            report['import'] = {'person': person.name, 'courses': len(person.courses)}
            persons = [Person(str(i), f'Test {i}', '', '', '', notes_confirmed=True) for i in range(3)]
            cells, diagnostics = solve_schedule(persons, 1, {'0:0': 1, '1:0': 1, '2:0': 1}, [])
            assert len({p for ids in cells.values() for p in ids}) == 3
            report['solver'] = 'passed'
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / 'export.xlsx'
                export_grid(path, '空课表', 1, {'0:0': 'Test 0'})
                book = load_workbook(path)
                assert book.active['B4'].value == 'Test 0'
                book.close()
            report['export'] = 'passed'
            report['status'] = 'passed'
        except Exception:
            report['status'] = 'failed'
            report['error'] = traceback.format_exc()
        Path(report_path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        return
    app = Application()
    def report(exc_type, exc, tb):
        log = ROOT / 'data' / 'error.log'
        log.parent.mkdir(parents=True, exist_ok=True)
        with log.open('a', encoding='utf-8') as handle:
            traceback.print_exception(exc_type, exc, tb, file=handle)
        messagebox.showerror('操作失败', f'{exc}\n详细信息已写入 data/error.log。')
    app.report_callback_exception = report
    app.mainloop()


if __name__ == '__main__':
    main()
