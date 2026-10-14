import tempfile
import unittest
from pathlib import Path
from openpyxl import load_workbook
from core import Person, SLOTS, parse_timetable, weeks_from_text, available, candidates, solve_schedule, export_grid, compact_weeks, total_free_cells


def people():
    return [Person(str(i), name, '2026-2027-1', '', '', notes_confirmed=True)
            for i, name in enumerate(['甲', '乙', '丙'])]


def rule(kind, pid, day=0, slot=0, weeks=None):
    return dict(kind=kind, person=pid, day=day, slot=slot, weeks=weeks or [], enabled=True)


class CoreTests(unittest.TestCase):
    def test_source(self):
        source = Path(__file__).parent / 'examples' / 'sample_20243006951.xls'
        p = parse_timetable(str(source))
        self.assertEqual(p.name, '陈锦瀚')
        self.assertEqual(p.id, '20243006951')
        self.assertEqual(len(p.courses), 27)
        self.assertEqual(p.errors, [])
        self.assertEqual(len([c for c in p.courses if c['cell'] == 'F7']), 4)
        self.assertEqual(p.notes, [])
        self.assertTrue(available(p, 1, 5, 0))
        p.notes_confirmed = True
        self.assertTrue(available(p, 1, 5, 0))
        self.assertFalse(available(p, 1, 0, 0))
        self.assertTrue(available(p, 6, 0, 0))
        self.assertFalse(available(p, 12, 0, 0))

    def test_weeks(self):
        self.assertEqual(weeks_from_text('1-3,7,10-12'), [1, 2, 3, 7, 10, 11, 12])
        for text in ['0', '3-1', 'foo', '61']:
            with self.assertRaises(ValueError):
                weeks_from_text(text)

    def test_total_free_and_custom_slots(self):
        p = people()[0]
        p.courses = [dict(day=0, weeks=[10, 11, 13], start=480, end=540)]
        slots = [('自定义', '08:30-10:00')]
        cells = total_free_cells([p], range(1, 19), slots)
        self.assertEqual(cells['0:0'], '甲 (1-9, 12, 14-18)')
        self.assertEqual(compact_weeks([4, 1, 3, 2, 8]), '1-4, 8')
        result, _ = solve_schedule([p], 10, {'0:0': 1}, [], slots)
        self.assertEqual(result['0:0'], [])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'total.xlsx'
            export_grid(path, '总空课表', '1-18', cells, slots)
            book = load_workbook(path)
            self.assertEqual(book.active['B4'].value, '甲 (1-9, 12, 14-18)')
            self.assertEqual(book.active.print_area, "'课表'!$A$1:$H$5")
            book.close()

    def test_exception_confirmation(self):
        p = people()[0]
        p.errors = ['测试异常']
        self.assertFalse(available(p, 1, 0, 0))
        p.errors_confirmed = True
        self.assertTrue(available(p, 1, 0, 0))

    def test_template_preservation(self):
        import zipfile
        import re
        from core import export_template
        p = people()[0]
        p.courses = [dict(day=0, weeks=[10, 11, 13], start=460, end=560)]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'template.xlsx'
            export_template(path, [p], list(range(1, 19)), True)
            book = load_workbook(path)
            self.assertEqual(book['海甸']['C2'].value, '甲 (1-9, 12, 14-18)')
            self.assertEqual(book['海甸']['C2'].font.name, '仿宋_GB2312')
            self.assertEqual(book['海甸']['C3'].value, '')
            book.close()
            with zipfile.ZipFile(Path(__file__).parent / 'resources/free_template.xlsx') as before, zipfile.ZipFile(path) as after:
                for name in before.namelist():
                    a, b = before.read(name), after.read(name)
                    if name in ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']:
                        pattern = rb'<c\b[^>]*\br="[C-I](?:[2-9]|1[0-9]|2[01])"[^>]*(?:/>|>.*?</c>)'
                        a, b = re.sub(pattern, b'', a, flags=re.S), re.sub(pattern, b'', b, flags=re.S)
                    self.assertEqual(a, b, name)

    def test_interval_boundaries(self):
        p = people()[0]
        p.courses = [dict(day=0, weeks=[1], start=440, end=460)]
        self.assertTrue(available(p, 1, 0, 0))
        p.courses[0]['end'] = 461
        self.assertFalse(available(p, 1, 0, 0))

    def test_exclusion_and_fixed(self):
        ps = people()
        rules = [rule('exclude', '0'), rule('fixed', '1')]
        result, notes = solve_schedule(ps, 1, {'0:0': 1}, rules)
        self.assertEqual(result['0:0'], ['1'])
        self.assertEqual(notes, [])
        with self.assertRaises(ValueError):
            solve_schedule(ps, 1, {'0:0': 1}, [rule('fixed', '0'), rule('exclude', '0')])
        with self.assertRaises(ValueError):
            solve_schedule(ps, 1, {'0:0': 1}, [rule('fixed', '0'), rule('fixed', '1')])

    def test_preference_priority(self):
        result, _ = solve_schedule(people(), 1, {'0:0': 1}, [rule('prefer', '2'), rule('prefer', '1')])
        self.assertEqual(result['0:0'], ['2'])
        result, _ = solve_schedule(people(), 1, {'0:0': 1}, [rule('prefer', '1'), rule('prefer', '2')])
        self.assertEqual(result['0:0'], ['1'])

    def test_balance_and_shortage(self):
        result, _ = solve_schedule(people(), 1, {'0:0': 1, '1:0': 1, '2:0': 1}, [])
        self.assertEqual(sorted(pid for ids in result.values() for pid in ids), ['0', '1', '2'])
        result, _ = solve_schedule(people()[:1], 1, {'0:0': 2}, [])
        self.assertEqual(result['0:0'], ['0'])

    def test_overlap_and_fixed_conflict(self):
        slots = list(SLOTS)
        slots[1] = ('重叠时段', '08:00-10:00')
        result, _ = solve_schedule(people()[:1], 1, {'0:0': 1, '0:1': 1}, [], slots)
        self.assertEqual(sum(len(ids) for ids in result.values()), 1)
        with self.assertRaises(ValueError):
            solve_schedule(people()[:1], 1, {'0:0': 1, '0:1': 1}, [rule('fixed', '0', slot=0), rule('fixed', '0', slot=1)], slots)

    def test_week_specific_rules(self):
        ps = people()
        r = rule('exclude', '0', weeks=[2])
        self.assertIn('0', candidates(ps, 1, 0, 0, [r]))
        self.assertNotIn('0', candidates(ps, 2, 0, 0, [r]))

    def test_export_roundtrip(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'table.xlsx'
            export_grid(path, '空课表', 5, {'0:0': '甲\n乙'})
            b = load_workbook(path)
            s = b.active
            self.assertEqual(s['B4'].value, '甲\n乙')
            self.assertEqual(s['A8'].value, '9、10、11节\n19:20-21:55')
            self.assertIn('A1:H1', [str(r) for r in s.merged_cells.ranges])
            self.assertEqual(s.freeze_panes, 'B4')
            b.close()


if __name__ == '__main__':
    unittest.main()
