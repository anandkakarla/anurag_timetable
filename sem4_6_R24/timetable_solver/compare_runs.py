"""Compare finished timetable runs side by side.

    python compare_runs.py OUT_FOLDER NAME [NAME ...]

For every NAME it reads OUT_FOLDER/NAME.xlsx (Sessions sheet) and, if present,
OUT_FOLDER/NAME_log.txt, and prints one row of quality counts. Lower is better
for everything except periods placed.
"""
import collections
import json
import os
import re
import sys

import openpyxl

BACK_TO_BACK = [(1, 2), (2, 3), (3, 4), (5, 6), (6, 7)]  # P4 -> P5 is across lunch


def building(room):
    room = str(room or '').strip().upper()
    if not room or room.startswith('TBC'):
        return None
    match = re.match(r'^([A-Z])-', room)
    return match.group(1) if match else room


def measure(folder, name):
    path = os.path.join(folder, name + '.xlsx')
    if not os.path.exists(path):
        return None
    status, placed, required = '?', None, None
    log = os.path.join(folder, name + '_log.txt')
    if os.path.exists(log):
        for line in open(log, encoding='utf-8', errors='replace'):
            if line.startswith('{"status"'):
                info = json.loads(line.split('}')[0] + '}')
                status = info.get('status', '?')
                placed, required = info.get('placed_periods'), info.get('required_periods')
    sheet = openpyxl.load_workbook(path, read_only=True)['Sessions']
    section = collections.defaultdict(dict)
    teacher = collections.defaultdict(dict)
    theory_days = collections.Counter()
    tbc_rooms = 0
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or row[2] is None:
            continue
        day, period, duration, code, component, groups, faculty, room = (
            row[2], int(row[3]), int(row[6]), row[7], str(row[9]), str(row[10]), str(row[12] or ''), str(row[13] or ''))
        if room.upper().startswith('TBC'):
            tbc_rooms += 1
        roots = {g.strip().split(':')[0] for g in groups.split(';') if g.strip()}
        if component.lower().startswith('theory'):
            for g in groups.split(';'):
                if g.strip():
                    theory_days[g.strip(), code, day] += 1
        for p in range(period, period + duration):
            for g in roots:
                section[g, day][p] = room
            for f in re.split(r'[;,]\s*', faculty):
                if f and not f.startswith('TBC'):
                    teacher[f, day][p] = room

    def changes(occupancy, by_building):
        count = 0
        for slots in occupancy.values():
            for a, b in BACK_TO_BACK:
                if a in slots and b in slots:
                    ra, rb = slots[a], slots[b]
                    if by_building:
                        ba, bb = building(ra), building(rb)
                        count += bool(ba and bb and ba != bb)
                    else:
                        count += bool(ra and rb and not ra.upper().startswith('TBC') and ra != rb)
        return count

    gaps = 0
    for slots in section.values():
        used = sorted(slots)
        gaps += sum(1 for p in range(used[0], used[-1]) if p not in slots and p != 4) if used else 0
    return dict(status=status,
                placed=f'{placed}/{required}' if placed is not None else '?',
                tbc_rooms=tbc_rooms,
                teacher_building_hops=changes(teacher, True),
                section_building_hops=changes(section, True),
                section_room_changes=changes(section, False),
                same_subject_twice=sum(1 for n in theory_days.values() if n > 1),
                section_gaps=gaps)


if __name__ == '__main__':
    folder, names = sys.argv[1], sys.argv[2:]
    columns = ['status', 'placed', 'tbc_rooms', 'teacher_building_hops', 'section_building_hops',
               'section_room_changes', 'same_subject_twice', 'section_gaps']
    print('run'.ljust(34) + ''.join(c.ljust(24) for c in columns))
    for name in names:
        result = measure(folder, name)
        if result is None:
            print(name.ljust(34) + 'no output file')
        else:
            print(name.ljust(34) + ''.join(str(result[c]).ljust(24) for c in columns))
