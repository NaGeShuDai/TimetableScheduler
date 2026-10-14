import re
from dataclasses import dataclass, field, asdict
from pathlib import Path

SLOTS = [
    ('1、2节', '07:40-09:20'), ('3、4节', '09:45-11:25'),
    ('5、6节', '14:30-16:10'), ('7、8节', '16:35-18:15'),
    ('9、10、11节', '19:20-21:55'), ('中午节次(一)', '11:35-13:05'),
    ('中午节次(二)', '13:05-14:20'),
]
DAYS = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
DEPARTMENTS = ['办公室', '采编部', '运营部', '技术部', '设计部']


def weeks_from_text(text):
    text = text.replace('，', ',').replace('、', ',').replace('－', '-').replace('—', '-')
    result = set()
    for part in text.split(','):
        part = part.strip()
        if not part:
            continue
        if re.fullmatch(r'\d+', part):
            result.add(int(part))
        elif re.fullmatch(r'\d+\s*-\s*\d+', part):
            lo, hi = map(int, re.split(r'\s*-\s*', part))
            if lo > hi:
                raise ValueError('周次范围起点大于终点')
            result.update(range(lo, hi + 1))
        else:
            raise ValueError('无法识别周次：' + part)
    if not result or min(result) < 1 or max(result) > 60:
        raise ValueError('周次必须在 1～60 内')
    return sorted(result)


def minutes(clock):
    h, m = map(int, clock.split(':'))
    if not 0 <= h <= 23 or not 0 <= m <= 59:
        raise ValueError('时间必须为 HH:MM')
    return h * 60 + m


def time_range(text):
    a, b = text.split('-')
    start, end = minutes(a), minutes(b)
    if start >= end:
        raise ValueError('结束时间必须晚于开始时间')
    return start, end


@dataclass
class Person:
    id: str
    name: str
    term: str
    group: str
    source: str
    courses: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    notes_confirmed: bool = True
    errors_confirmed: bool = False
    enabled: bool = True
    campus: str = '海甸'
    department: str = '办公室'

    def to_dict(self):
        return asdict(self)


def read_rows(path):
    if Path(path).suffix.lower() == '.xls':
        import xlrd
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        return [sheet.row_values(r) for r in range(sheet.nrows)]
    from openpyxl import load_workbook
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        return [list(row) for row in book.worksheets[0].iter_rows(values_only=True)]
    finally:
        book.close()


def parse_timetable(path):
    rows = [[str(v or '').strip() for v in row] for row in read_rows(path)]
    if len(rows) < 4:
        raise ValueError('文件没有完整课表')
    title = rows[0][0]
    match = re.search(r'\s+(\S+)\s+学生个人课表', title)
    name = match.group(1) if match else Path(path).stem
    meta = ' '.join(rows[1])
    term = re.search(r'学年学期[：:]\s*(\S+)', meta)
    group = re.search(r'班级[：:]\s*(.*?)\s+(?:专业|院系)[：:]', meta)
    sid = re.search(r'(\d{8,})', Path(path).stem)
    person = Person(sid.group(1) if sid else name, name, term.group(1) if term else '',
                    group.group(1) if group else '', str(path))
    header = next((i for i, row in enumerate(rows) if '星期一' in row), None)
    if header is None:
        raise ValueError('未找到星期一至星期日表头')
    day_columns = {d: rows[header].index(label) for d, label in enumerate(DAYS)
                   if label in rows[header]}
    if len(day_columns) != 7:
        raise ValueError('星期表头不完整')
    for r in range(header + 1, len(rows)):
        row = rows[r]
        clock = re.search(r'(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})', row[0])
        if not clock:
            continue
        start, end = time_range(clock.group(1) + '-' + clock.group(2))
        for day, col in day_columns.items():
            text = row[col] if col < len(row) else ''
            if not text:
                continue
            # Week/period lines delimit records even when blank separators are lost.
            lines = [s.strip() for s in text.splitlines() if s.strip()]
            markers = [i for i, line in enumerate(lines) if re.search(r'\(.*?周.*?\).*?\[', line)]
            if not markers:
                person.errors.append(f'第{r + 1}行/{DAYS[day]}：无法识别课程时间：{text}')
                continue
            previous = 0
            for index, marker in enumerate(markers):
                try:
                    line = lines[marker]
                    m = re.fullmatch(r'([\d\s,，、\-－—]+)\(([^)]*周[^)]*)\)\[([^]]+)\]', line)
                    if not m or marker - previous != 2:
                        raise ValueError('课程记录结构不符合名称/教师/周次/地点格式')
                    weeks = weeks_from_text(m.group(1))
                    parity = m.group(2)
                    if '单' in parity:
                        weeks = [w for w in weeks if w % 2]
                    elif '双' in parity:
                        weeks = [w for w in weeks if not w % 2]
                    periods = list(map(int, re.findall(r'\d+', m.group(3))))
                    if len(periods) == 2 and periods[1] > periods[0]:
                        periods = list(range(periods[0], periods[1] + 1))
                    location_end = markers[index + 1] - 2 if index + 1 < len(markers) else len(lines)
                    person.courses.append(dict(name=lines[marker - 2], teacher=lines[marker - 1],
                        weeks=weeks, day=day, periods=periods, start=start, end=end,
                        location=' '.join(lines[marker + 1:location_end]), raw=text,
                        cell=f'{chr(65 + col)}{r + 1}', conservative=True))
                    previous = location_end
                except ValueError as exc:
                    person.errors.append(f'第{r + 1}行/{DAYS[day]}：{exc}')
    if not person.courses:
        raise ValueError('没有成功识别的课程')
    return person


def available(person, week, day, slot, slots=SLOTS):
    if not person.enabled or (person.errors and not person.errors_confirmed):
        return False
    start, end = time_range(slots[slot][1])
    return not any(c['day'] == day and week in c['weeks'] and start < c['end'] and c['start'] < end
                   for c in person.courses)


def rule_matches(rule, week, day, slot):
    return rule.get('enabled', True) and (not rule.get('weeks') or week in rule['weeks']) and \
        (rule['day'] is None or rule['day'] == day) and (rule['slot'] is None or rule['slot'] == slot)


def candidates(people, week, day, slot, rules, slots=SLOTS):
    pool = [p.id for p in people if available(p, week, day, slot, slots)]
    exclusions = {r['person'] for r in rules if r['kind'] == 'exclude' and rule_matches(r, week, day, slot)}
    return [pid for pid in pool if pid not in exclusions]


def solve_schedule(people, week, demands, rules, slots=SLOTS, locked=None):
    from ortools.sat.python import cp_model
    model = cp_model.CpModel()
    ids = [p.id for p in people]
    variables, pools = {}, {}
    for key, count in demands.items():
        day, slot = map(int, key.split(':'))
        pool = candidates(people, week, day, slot, rules, slots)
        pools[key] = pool
        for pid in pool:
            variables[key, pid] = model.new_bool_var(f'{key}_{pid}')
        model.add(sum(variables[key, pid] for pid in pool) <= count)
    for key, count in demands.items():
        day, slot = map(int, key.split(':'))
        fixed = {r['person'] for r in rules if r['kind'] == 'fixed' and rule_matches(r, week, day, slot)}
        if locked and key in locked:
            fixed.update(locked[key])
        if len(fixed) > count:
            raise ValueError(f'{DAYS[day]} {slots[slot][0]}：固定人员超过所需人数')
        for pid in fixed:
            if pid not in pools[key]:
                raise ValueError(f'{DAYS[day]} {slots[slot][0]}：固定人员不可用或已被排除')
            model.add(variables[key, pid] == 1)
        if locked and key in locked:
            for pid in pools[key]:
                if pid not in locked[key]:
                    model.add(variables[key, pid] == 0)
    keys = list(demands)
    for pid in ids:
        for i, a in enumerate(keys):
            ad, at = map(int, a.split(':'))
            astart, aend = time_range(slots[at][1])
            for b in keys[i + 1:]:
                bd, bt = map(int, b.split(':'))
                bs, be = time_range(slots[bt][1])
                if ad == bd and astart < be and bs < aend and (a, pid) in variables and (b, pid) in variables:
                    model.add(variables[a, pid] + variables[b, pid] <= 1)
    total_bound = sum(count * (time_range(slots[int(k.split(':')[1])][1])[1] -
                              time_range(slots[int(k.split(':')[1])][1])[0]) for k, count in demands.items())
    loads = {}
    active_ids = [pid for pid in ids if any(p == pid for _, p in variables)]
    for pid in active_ids:
        loads[pid] = model.new_int_var(0, total_bound, 'load_' + pid)
        model.add(loads[pid] == sum(v * (time_range(slots[int(k.split(':')[1])][1])[1] -
                                       time_range(slots[int(k.split(':')[1])][1])[0])
                                    for (k, p), v in variables.items() if p == pid))
    objectives = []
    for r in rules:
        if r['kind'] == 'prefer' and r.get('enabled', True):
            expr = sum(v for (k, p), v in variables.items() if p == r['person'] and
                       rule_matches(r, week, *map(int, k.split(':'))))
            objectives.append(('自定义优先安排', expr, True))
    objectives.append(('满足班次人数', sum(variables.values()), True))
    if loads:
        spread_terms = []
        for i, pid in enumerate(active_ids):
            for other in active_ids[i + 1:]:
                diff = model.new_int_var(0, total_bound, 'diff_' + pid + '_' + other)
                model.add_abs_equality(diff, loads[pid] - loads[other])
                spread_terms.append(diff)
        objectives.append(('均衡总时长', sum(spread_terms), False))
        for label, predicate in [('均衡晚班', lambda d, s: time_range(slots[s][1])[0] >= 18 * 60), ('均衡周末', lambda d, s: d >= 5)]:
            counts = []
            for pid in active_ids:
                n = model.new_int_var(0, len(keys), label + pid)
                model.add(n == sum(v for (k, p), v in variables.items() if p == pid and predicate(*map(int, k.split(':')))))
                counts.append(n)
            hi, lo = model.new_int_var(0, len(keys), label + 'hi'), model.new_int_var(0, len(keys), label + 'lo')
            model.add_max_equality(hi, counts)
            model.add_min_equality(lo, counts)
            objectives.append((label, hi - lo, False))
    excess = []
    for pid in active_ids:
        for day in range(7):
            n = sum(v for (k, p), v in variables.items() if p == pid and int(k.split(':')[0]) == day)
            e = model.new_int_var(0, len(keys), f'excess_{pid}_{day}')
            model.add(e >= n - 1)
            excess.append(e)
    objectives.append(('减少一天多班', sum(excess), False))
    # Each proven optimum becomes a constraint before optimizing the next priority.
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 5
    solver.parameters.num_search_workers = 1
    diagnostics = []
    answer = None
    for label, expr, maximize in objectives:
        model.maximize(expr) if maximize else model.minimize(expr)
        status = solver.solve(model)
        if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            if answer is None:
                raise ValueError('规则存在冲突，无法生成排班；请检查固定安排和重叠时段')
            diagnostics.append(label + '未完成优化，保留上一阶段有效方案')
            break
        answer = {k: [pid for pid in pools[k] if solver.value(variables[k, pid])] for k in keys}
        if status != cp_model.OPTIMAL:
            diagnostics.append(label + '得到可行方案，但尚未证明最优；后续优先级未优化')
            break
        model.add(expr == round(solver.objective_value))
    return answer or {}, diagnostics


def compact_weeks(weeks):
    values = sorted(set(weeks))
    groups = []
    for value in values:
        if groups and value == groups[-1][1] + 1:
            groups[-1][1] = value
        else:
            groups.append([value, value])
    return ', '.join(str(a) if a == b else f'{a}-{b}' for a, b in groups)


def total_free_cells(people, weeks, slots=SLOTS):
    return {f'{d}:{s}': '\n'.join(f'{p.name} ({compact_weeks(free)})'
            for p in people if (free := [w for w in weeks if available(p, w, d, s, slots)])) or '无'
            for d in range(7) for s in range(len(slots))}


def export_template(path, people, weeks, total=False):
    import sys
    import zipfile
    import xml.etree.ElementTree as ET
    root = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    template = root / 'resources' / 'free_template.xlsx'
    if not template.exists() and getattr(sys, 'frozen', False):
        template = Path(sys._MEIPASS) / 'resources' / 'free_template.xlsx'
    ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    replacements = {}
    # Preserve every ZIP member except the values of the 140 roster cells per sheet.
    with zipfile.ZipFile(template) as source:
        strings = []
        if 'xl/sharedStrings.xml' in source.namelist():
            strings = [''.join(si.itertext()) for si in ET.fromstring(source.read('xl/sharedStrings.xml'))]
        for sheet_index, campus in enumerate(['海甸', '观澜湖'], 1):
            member = f'xl/worksheets/sheet{sheet_index}.xml'
            data = source.read(member).decode('utf-8')
            xml = ET.fromstring(data)
            existing = {c.attrib['r']: c for c in xml.findall('.//m:sheetData/m:row/m:c', ns)}
            for slot in range(5):
                for row in range(2 + slot * 5, 7 + slot * 5):
                    department_cell = existing.get(f'B{row}')
                    department = ''.join(department_cell.itertext()) if department_cell is not None else ''
                    if department_cell is not None and department_cell.get('t') == 's':
                        department = strings[int(department)]
                    if department not in DEPARTMENTS:
                        raise ValueError('模板部门无法识别：' + department)
                    for day in range(7):
                        coord = f'{chr(67 + day)}{row}'
                        lines = []
                        for person in people:
                            if person.campus != campus or person.department != department:
                                continue
                            free = [w for w in weeks if available(person, w, day, slot, SLOTS)]
                            if free:
                                lines.append(f'{person.name} ({compact_weeks(free)})' if total else person.name)
                        cell = existing.get(coord)
                        if cell is None:
                            raise ValueError('模板缺少名单单元格：' + coord)
                        style = cell.get('s', '0')
                        from xml.sax.saxutils import escape
                        value = escape('\n'.join(lines))
                        replacement = f'<c r="{coord}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{value}</t></is></c>'
                        pattern = rf'<c\b[^>]*\br="{coord}"[^>]*?(?:/>|>.*?</c>)'
                        data, count = re.subn(pattern, lambda _: replacement, data, count=1, flags=re.S)
                        if count != 1:
                            raise ValueError('无法匹配模板单元格：' + coord)
            replacements[member] = data.encode('utf-8')
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as target:
            for info in source.infolist():
                target.writestr(info, replacements.get(info.filename, source.read(info.filename)))


def export_grid(path, title, week, cells, slots=SLOTS, term=''):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
    from openpyxl.utils import get_column_letter
    book = Workbook()
    ws = book.active
    ws.title = '课表'
    ws.merge_cells('A1:H1')
    ws['A1'] = title
    ws.merge_cells('A2:H2')
    ws['A2'] = f'学期：{term}    第 {week} 教学周    时间采用课表整行范围'
    ws.append([''] + DAYS)
    for slot, (label, clock) in enumerate(slots):
        ws.append([label + '\n' + clock] + [cells.get(f'{day}:{slot}', '') for day in range(7)])
    footer = len(slots) + 4
    ws.merge_cells(start_row=footer, start_column=2, end_row=footer, end_column=8)
    ws.cell(footer, 2, '按教学周计算；未确认解析异常的人员不列为空闲。')
    border = Border(*( [Side(style='thin', color='C8CDD2')] * 4 ))
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.font = Font(name='Microsoft YaHei', size=11)
            cell.border = border
            if cell.row <= 3 or cell.column == 1:
                cell.fill = PatternFill('solid', fgColor='E8F0F3')
    ws['A1'].font = Font(name='Microsoft YaHei', size=16, bold=True)
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 26
    ws.row_dimensions[3].height = 26
    for slot in range(len(slots)):
        longest = max((len(cells.get(f'{d}:{slot}', '').splitlines()) for d in range(7)), default=1)
        ws.row_dimensions[slot + 4].height = max(66, 19 * longest + 12)
    ws.column_dimensions['A'].width = 23
    for col in range(2, 9):
        longest = max((sum(2 if ord(ch) > 127 else 1 for ch in line)
                       for s in range(len(slots)) for line in cells.get(f'{col - 2}:{s}', '').splitlines()), default=0)
        ws.column_dimensions[get_column_letter(col)].width = max(22, longest + 3)
    ws.row_dimensions[footer].height = 30
    ws.freeze_panes = 'B4'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.print_area = f'A1:H{footer}'
    book.save(path)
