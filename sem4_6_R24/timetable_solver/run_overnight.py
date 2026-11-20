"""Overnight run: try each attempt in order and stop at the first one that finds a
timetable. If an attempt is proven impossible, diagnose_rooms.py runs right after it
and writes which room list is the bottleneck. Logs go to ..\\output\\<name>_log.txt,
one-line results to ..\\output\\overnight_summary.txt.

Run from the timetable_solver folder:
    python run_overnight.py
"""
import datetime
import os
import re
import subprocess
import sys

OUT = os.path.join('..', 'output')
HINT = os.path.join(OUT, 'timetable_nov19_v21.xlsx')

# (name, solver script, input workbook, seconds, extra solver arguments).
# All five run every time (about 4.5 hours in total), then compare_runs.py writes one
# table so the timetables can be compared side by side.
ATTEMPTS = [
    # A. Agreed setup, current weights: hops 500, home room 300, fewer rooms 200,
    #    same subject per day 8.
    ('timetable_nov20_A_current', 'solve_timetable_p4p7_hop.py',
     'workload_all_departments_R24_nov19_v34.xlsx', 3000, []),
    # B. Same, room and hop weights halved (easier to search).
    ('timetable_nov20_B_halfroom', 'solve_timetable_p4p7_hop.py',
     'workload_all_departments_R24_nov19_v34_weightsB.xlsx', 3000, []),
    # C. Same as A, but same subject twice a day weighted 2000.
    ('timetable_nov20_C_samesubject', 'solve_timetable_p4p7_hop.py',
     'workload_all_departments_R24_nov19_v34_weightsC.xlsx', 3000, []),
    # E. v30: same rules, but no assigned home rooms (whole department room list).
    ('timetable_nov20_E_v30', 'solve_timetable_p4p7_hop.py',
     'workload_all_departments_R24_nov19_v30.xlsx', 3600, []),
    # D. Safety net: as A, but room lists soft (a TBC room costs 5000 instead of
    #    making the run fail).
    ('timetable_nov20_D_softrooms', 'solve_timetable_p4p7_hop.py',
     'workload_all_departments_R24_nov19_v34.xlsx', 3000, ['--soft-rooms', '5000']),
]


def summary(line):
    stamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    text = f'[{stamp}] {line}'
    print(text, flush=True)
    with open(os.path.join(OUT, 'overnight_summary.txt'), 'a', encoding='utf-8') as f:
        f.write(text + '\n')


def stream(command, log_path, watch):
    """Run a command, copy its output to a log, return the lines matching watch."""
    hits = []
    with open(log_path, 'w', encoding='utf-8') as log:
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                   text=True, encoding='utf-8', errors='replace')
        for line in process.stdout:
            log.write(line)
            log.flush()
            if 'EXISTING LOAD' not in line and 'UNPLACED' not in line:
                print(line, end='', flush=True)
            if watch(line):
                hits.append(line.rstrip())
        process.wait()
    return process.returncode, hits


def run(name, script, workbook, seconds, extra):
    summary(f'START {name}: {script} on {workbook}, {seconds}s')
    command = [sys.executable, script, '--input', os.path.join(OUT, workbook), '--hint', HINT,
               '--seconds', str(seconds), '--workers', '8', '--out', OUT, '--name', name] + extra
    code, hits = stream(command, os.path.join(OUT, f'{name}_log.txt'),
                        lambda l: '"status"' in l or l.startswith(('Building hops', 'Finished after')))
    status = 'NO STATUS'
    for line in hits:
        match = re.search(r'"status":\s*"(\w+)"', line)
        if match:
            status = match.group(1)
        else:
            summary(f'  {name}: {line}')
    summary(f'END {name}: status {status}, exit code {code}')
    if status == 'INFEASIBLE' and not extra:
        summary(f'  {name}: proven impossible -- diagnosing room lists (up to 15 min)')
        _, found = stream([sys.executable, 'diagnose_rooms.py', script, os.path.join(OUT, workbook), '900'],
                          os.path.join(OUT, f'{name}_room_diagnosis.txt'),
                          lambda l: 'Room diagnosis' in l or l.startswith('  ROOMS'))
        for line in found:
            summary(f'  {name}: {line.strip()}')
    elif status == 'UNKNOWN':
        summary(f'  {name}: no timetable found in time, but not proven impossible (more time may work)')
    return status in ('FEASIBLE', 'OPTIMAL')


if __name__ == '__main__':
    diagnosed = False
    for name, script, workbook, seconds, extra in ATTEMPTS:
        if diagnosed and not extra and 'v34' in workbook:
            # A, B and C share the same hard rules: if one is proven impossible,
            # so are the others. Skip them (E and D still run).
            summary(f'SKIP {name}: same hard rules as an attempt already proven impossible')
            continue
        found = run(name, script, workbook, seconds, extra)
        if not found and os.path.exists(os.path.join(OUT, f'{name}_room_diagnosis.txt')):
            diagnosed = True
    names = ['timetable_nov19_v21'] + [attempt[0] for attempt in ATTEMPTS]
    _, table = stream([sys.executable, 'compare_runs.py', OUT] + names,
                      os.path.join(OUT, 'weight_comparison.txt'), lambda l: True)
    summary('COMPARISON (lower is better except placed), also in weight_comparison.txt:')
    for line in table:
        summary('  ' + line)
    summary('DONE.')
