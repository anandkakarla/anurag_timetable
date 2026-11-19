"""Small synthetic checks. Never reads or solves the production timetable."""
import importlib.util
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location('changed_solver', Path(__file__).with_name('solve_timetable.py'))
solver = importlib.util.module_from_spec(spec)
spec.loader.exec_module(solver)


def activity(identifier, kind='Theory', duration=1, faculty='F1'):
    return dict(id=identifier, dept='TEST', semester=4, code=identifier, name=identifier,
                kind=kind, groups=['TEST-S4-A'], faculty=[faculty], faculty_names=faculty,
                hours=duration, sessions=1, duration=duration, rooms=['R1'], parallel='',
                eslot='', source='', assumptions='', course_id=identifier, home_rooms=['R1'])


def data(activities, busy=None):
    rules = solver.default_constraints()
    for rule in rules:
        rule['enabled'] = rule['name'] in {
            'NO_SECTION_OVERLAP', 'NO_FACULTY_CLASH', 'NO_ROOM_CLASH',
            'EXISTING_SCHEDULE_CLASH', 'MAX_FACULTY_THEORY_PER_DAY',
            'MAX_FACULTY_TOTAL_PER_DAY', 'FIXED_SLOT'}
    return dict(days=['Monday'], times=[(0, 1), (1, 2), (2, 3), (3, 4), (5, 6), (6, 7), (7, 8)],
                activities=activities, busy=busy or [], faculty_limits={}, constraints=rules,
                fixed_sessions=[], settings={})


def rule(model_data, name):
    return next(row for row in model_data['constraints'] if row['name'] == name)


class ConstraintTests(unittest.TestCase):
    def placed(self, model_data):
        return solver.solve(model_data, 2, workers=1)['placed_periods']

    def test_theory_cap(self):
        self.assertEqual(self.placed(data([activity(str(i)) for i in range(5)])), 4)

    def test_lab_counts_every_period_in_total(self):
        model_data = data([activity(str(i)) for i in range(3)] + [activity('L', 'Lab', 3)])
        self.assertEqual(self.placed(model_data), 5)

    def test_existing_unknown_counts_as_theory(self):
        busy = [dict(faculty=['F1'], day='Monday', start=0, end=2, component='Unknown')]
        self.assertEqual(self.placed(data([activity(str(i)) for i in range(3)], busy)), 2)

    def test_existing_lab_counts_only_toward_total(self):
        busy = [dict(faculty=['F1'], day='Monday', start=0, end=2, component='Lab')]
        self.assertEqual(self.placed(data([activity(str(i)) for i in range(3)], busy)), 3)

    def test_soft_theory_cap_does_not_drop_periods(self):
        model_data = data([activity(str(i)) for i in range(5)])
        cap = rule(model_data, 'MAX_FACULTY_THEORY_PER_DAY')
        cap.update(type='SOFT', value=1)
        self.assertEqual(self.placed(model_data), 5)

    def test_disabled_theory_cap(self):
        model_data = data([activity(str(i)) for i in range(5)])
        rule(model_data, 'MAX_FACULTY_THEORY_PER_DAY')['enabled'] = False
        self.assertEqual(self.placed(model_data), 5)

    def test_faculty_override(self):
        model_data = data([activity(str(i)) for i in range(5)])
        cap = dict(rule(model_data, 'MAX_FACULTY_TOTAL_PER_DAY'))
        cap.update(scope='Faculty', target='F1', value=2)
        model_data['constraints'].append(cap)
        self.assertEqual(self.placed(model_data), 2)

    def test_required_periods_reports_infeasible_with_conflicting_cap(self):
        model_data = data([activity(str(i)) for i in range(5)])
        rule(model_data, 'PERIODS_PER_WEEK')['enabled'] = True
        self.assertEqual(solver.solve(model_data, 2, workers=1)['status'], 'INFEASIBLE')

    def test_fixed_session(self):
        model_data = data([activity('X')])
        model_data['fixed_sessions'] = [dict(activity_id='X', session=1, day='Monday', period=6, room='R1')]
        result = solver.solve(model_data, 2, workers=1)
        self.assertEqual(result['schedule'][0]['period'], 6)

    def test_all_quality_handlers_build(self):
        model_data = data([activity('T'), activity('T2'), activity('L', 'Lab', 2), activity('L2', 'Lab', 2)])
        model_data['constraints'] = solver.default_constraints()
        model, _ = solver.build_cp_model(model_data)
        self.assertEqual(model.validate(), '')

    def test_hard_subject_daily_limit(self):
        activities = [activity('T1'), activity('T2')]
        for item in activities:
            item['code'] = 'SAME'
        model_data = data(activities)
        rule(model_data, 'MAX_SAME_SUBJECT_PER_DAY').update(enabled=True, type='HARD', value=1)
        self.assertEqual(self.placed(model_data), 1)

    def test_hard_student_gap_with_fixed_sessions(self):
        model_data = data([activity('A'), activity('B')])
        rule(model_data, 'COMPACT_SCHEDULE').update(enabled=True, type='HARD')
        model_data['fixed_sessions'] = [
            dict(activity_id='A', session=1, day='Monday', period=1, room='R1'),
            dict(activity_id='B', session=1, day='Monday', period=3, room='R1')]
        self.assertEqual(solver.solve(model_data, 2, workers=1)['status'], 'INFEASIBLE')

    def test_hard_room_preference(self):
        item = activity('A')
        item['rooms'] = ['R2']
        model_data = data([item])
        rule(model_data, 'PREFER_HOME_ROOM').update(enabled=True, type='HARD')
        self.assertEqual(self.placed(model_data), 0)

    def test_fixed_load_preflight(self):
        model_data = data([activity('A')], [dict(faculty=['F1'], day='Monday', start=0, end=4, component='Theory')])
        self.assertEqual(solver.existing_cap_warnings(model_data), [])
        model_data['busy'].append(dict(faculty=['F1'], day='Monday', start=5, end=6, component='Theory'))
        self.assertTrue(any('THEORY' in message for message in solver.existing_cap_warnings(model_data)))

    def test_joe_exception_keeps_six_period_day_closed_to_extra_teaching(self):
        busy = [dict(faculty=['F1'], day='Monday', start=0, end=4, component='JOE Lab', daily_total_exception=6),
                dict(faculty=['F1'], day='Monday', start=5, end=7, component='JOE Lab', daily_total_exception=6)]
        model_data = data([activity('A')], busy)
        self.assertEqual(solver.existing_cap_warnings(model_data), [])
        self.assertEqual(self.placed(model_data), 0)

    def test_doe_and_project_days_get_six_period_cap(self):
        for component in ('DOE', 'DOE Lab', 'JOE', 'Project'):
            model_data = data([activity('A')], [dict(faculty=['F1'], day='Monday', start=0, end=3, component=component, daily_total_exception=None)])
            self.assertEqual(solver.existing_daily_exceptions(model_data), {('F1', 'Monday'): 6})
            self.assertEqual(solver.existing_cap_warnings(model_data), [])

    def test_more_than_six_existing_project_periods_accepted_without_additions(self):
        model_data = data([activity('A')], [dict(faculty=['F1'], day='Monday', start=0, end=8, component='DOE', daily_total_exception=6)])
        self.assertEqual(solver.existing_daily_exceptions(model_data), {('F1', 'Monday'): 6})
        self.assertEqual(self.placed(model_data), 0)

    def test_existing_theory_overload_blocks_new_labs_too(self):
        busy = [dict(faculty=['F1'], day='Monday', start=0, end=4, component='Theory'),
                dict(faculty=['F1'], day='Monday', start=5, end=6, component='Theory')]
        self.assertEqual(self.placed(data([activity('L', 'Lab')], busy)), 0)

    def test_existing_total_overload_blocks_additions_even_for_soft_cap(self):
        busy = [dict(faculty=['F1'], day='Monday', start=0, end=4, component='Lab'),
                dict(faculty=['F1'], day='Monday', start=5, end=7, component='Lab')]
        model_data = data([activity('A')], busy)
        rule(model_data, 'MAX_FACULTY_TOTAL_PER_DAY')['type'] = 'SOFT'
        self.assertEqual(self.placed(model_data), 0)

    def test_overload_does_not_block_another_day(self):
        busy = [dict(faculty=['F1'], day='Monday', start=0, end=4, component='Theory'),
                dict(faculty=['F1'], day='Monday', start=5, end=6, component='Theory')]
        model_data = data([activity('A')], busy)
        model_data['days'].append('Tuesday')
        result = solver.solve(model_data, 2, workers=1)
        self.assertEqual(result['placed_periods'], 1)
        self.assertEqual(result['schedule'][0]['day'], 'Tuesday')

    def test_ordinary_lab_cannot_claim_joe_exception(self):
        model_data = data([activity('A')], [dict(faculty=['F1'], day='Monday', start=0, end=2, component='Lab', daily_total_exception=6)])
        with self.assertRaises(solver.InputError):
            solver.existing_cap_warnings(model_data)

    def test_settings_validation(self):
        class Sheet:
            def __init__(self, rows):
                self.rows = rows
            def iter_rows(self, values_only=True):
                return iter(self.rows)
        class Workbook:
            sheetnames = ['Constraints']
            def __init__(self, rows):
                self.sheet = Sheet(rows)
            def __getitem__(self, name):
                return self.sheet
        headers = ['Name', 'Enabled', 'Type', 'Value', 'Weight', 'Scope', 'Target']
        rows = [[r['name'], 'YES' if r['enabled'] else 'NO', r['type'], r['value'], r['weight'], r['scope'], r['target']]
                for r in solver.default_constraints()]
        self.assertEqual(len(solver.read_constraints(Workbook([headers] + rows), [activity('X')], [], 7)), 20)
        for column, bad in [(0, 'MISTYPED_RULE'), (1, 'MAYBE'), (2, 'SOFT'), (3, 2), (4, -1), (5, 'Unknown')]:
            modified = [list(row) for row in rows]
            modified[0][column] = bad
            with self.subTest(column=column):
                with self.assertRaises(solver.InputError):
                    solver.read_constraints(Workbook([headers] + modified), [activity('X')], [], 7)
        with self.assertRaises(solver.InputError):
            solver.read_constraints(Workbook([headers] + rows + [rows[0]]), [activity('X')], [], 7)
        with self.assertRaises(solver.InputError):
            solver.read_constraints(Workbook([headers] + rows[1:]), [activity('X')], [], 7)


if __name__ == '__main__':
    unittest.main()
