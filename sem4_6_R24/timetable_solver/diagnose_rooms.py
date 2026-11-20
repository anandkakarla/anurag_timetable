"""Find which room list makes a timetable impossible.

Every class may also fall back to a "TBC" room, but each group of classes sharing
the same room list gets a switch that forbids the fallback. CP-SAT then names the
smallest set of switches that cannot all be on together: those room lists are the
real bottleneck.

    python diagnose_rooms.py SOLVER.py WORKBOOK.xlsx [SECONDS]
"""
import contextlib
import importlib.util
import io
import sys

from ortools.sat.python import cp_model

solver_path, workbook = sys.argv[1], sys.argv[2]
seconds = float(sys.argv[3]) if len(sys.argv) > 3 else 900
spec = importlib.util.spec_from_file_location('timetable_solver', solver_path)
solver = importlib.util.module_from_spec(spec)
spec.loader.exec_module(solver)

with contextlib.redirect_stdout(io.StringIO()):
    data = solver.read_input(workbook)
for activity in data['activities']:
    if not activity['rooms'][0].startswith('TBC'):
        activity['rooms'] = activity['rooms'] + ['TBC-ROOM-' + activity['id']]
with contextlib.redirect_stdout(io.StringIO()):
    model, entries = solver.build_cp_model(data)

switches = {}
for activity, _, _, _, rooms, _ in entries:
    real = [room for room, _ in rooms if not room.startswith('TBC')]
    if not real:
        continue
    key = ' | '.join(real)
    if key not in switches:
        switches[key] = model.new_bool_var('rooms_' + key)
    for room, use in rooms:
        if room.startswith('TBC'):
            model.add(use == 0).only_enforce_if(switches[key])
model.clear_objective()
model.add_assumptions(list(switches.values()))

cp = cp_model.CpSolver()
cp.parameters.max_time_in_seconds = seconds
cp.parameters.num_search_workers = 1  # needed to get the blocking set back
status = cp.solve(model)
print('Room diagnosis:', cp.status_name(status), f'{cp.wall_time:.0f}s')
if status == cp_model.INFEASIBLE:
    names = {var.index: key for key, var in switches.items()}
    core = [names.get(i, str(i)) for i in cp.sufficient_assumptions_for_infeasibility()]
    if core:
        print('Room diagnosis: these room lists cannot fit everything assigned to them:')
        for key in core:
            classes = sorted({a['id'] for a, _, _, _, rooms, _ in entries
                              if ' | '.join(r for r, _ in rooms if not r.startswith('TBC')) == key})
            print(f'  ROOMS {key}: {len(classes)} classes, e.g. {", ".join(classes[:6])}')
    else:
        print('Room diagnosis: impossible even with TBC rooms allowed, so the cause is not rooms '
              '(check teachers, fixed sessions or section hours).')
elif status in (cp_model.FEASIBLE, cp_model.OPTIMAL):
    print('Room diagnosis: room lists are not the problem; a timetable exists with them.')
else:
    print('Room diagnosis: no conclusion within the time limit.')
