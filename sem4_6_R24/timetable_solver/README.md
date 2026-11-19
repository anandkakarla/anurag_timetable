# B.Tech Timetable Solver: Semesters 4 & 6

Builds a clash-free weekly timetable for all departments from one workbook. Each department coordinator fills in their part; one script solves the timetable and exports it.

## Files

| File | Purpose |
|---|---|
| `workload_all_departments.xlsx` | **The only input.** It holds the workload for every department. |
| `solve_timetable.py` | Reads the input, builds the timetable and writes the outputs. |
| `run_timetable.ps1` | Optional Windows wrapper that runs the script with the bundled Python packages. |
| `timetable_output.xlsx` | Output: timetable sheets per department, plus Sessions, Unplaced and Assumptions. |
| `timetable_output.html` | Output: a browsable timetable with Master, Section, Teacher, Room and Course views. |

## Workflow

1. Coordinators enter their department's data on the **Workload Intake** page, or edit their sheets in the workbook directly.
2. The admin downloads the combined workbook as `workload_all_departments.xlsx`.
3. Run the solver:
   ```
   python solve_timetable.py                    # solve (default 600 s), write Excel + HTML
   python solve_timetable.py --excel            # solve, write Excel only
   python solve_timetable.py --html             # solve, write HTML only
   python solve_timetable.py --html-from-excel  # no solving: build HTML from timetable_output.xlsx
   python solve_timetable.py --html-from-excel run2.xlsx   # ...or from another generated Excel
   python solve_timetable.py --check-only       # validate the input and constraints only
   ```
4. Open `timetable_output.html` and review it. Fix any placeholders or unplaced items in the input, then run again.

If a previous `timetable_output.xlsx` is in the folder, the solver uses it as a warm start, so re-runs are faster. You need Python with `ortools` and `openpyxl` installed.

## Input workbook: what each sheet holds

| Sheet | Content |
|---|---|
| Sections | Department, semester, section letters and home rooms. |
| Rooms | Classrooms and labs for each department. |
| Faculty | Employee ID, name and department. |
| `<DEPT> S<n>` (e.g. `AI S6`) | One row per course, with lab courses on separate rows. Columns: code, name, type (Theory / Lab / IP / PSE), periods per week, lab block length, faculty ID for each section, rooms, and note. |
| Merged classes | Sections taught together as one class (code, type, sections). |
| Lab batches | Labs split into B1/B2 batches, and which lab each one swaps with. |
| Electives | Elective pools: the pooled sections and the subjects offered. All subjects in a pool run at the same time. |
| Faculty busy | Existing commitments, such as first-year classes, which are blocked out. |
| Settings | Days, period times, title and solve time. |
| Curriculum | Reference periods for each course, used for cross-checks. |

## Rules the solver enforces

- Week: Monday to Friday, 7 periods a day, lunch from 12:40 to 13:20. No session crosses lunch.
- Labs run as one unbroken block (2–4 periods). A 4-period lab can only go in the morning.
- No teacher, section, lab batch or room is in two places at once.
- Each teacher's first-year busy times are blocked.
- Merged sections are scheduled together.
- Swapped lab batches (B1/B2) run at the same time in different labs.
- All subjects in an elective pool run in the same slot.
- Goal: place every period, and prefer real rooms over provisional (TBC) rooms.

## Configurable constraints

Edit the Constraints sheet. Every row has a working handler in solve_timetable.py.

- Enabled is YES or NO. Type is HARD or SOFT where the rule supports it.
- Value is a positive integer for a cap or numbered period. Leave it empty for rules without a value.
- Weight is a positive integer. Higher SOFT weights give that preference more importance.
- Scope is Global, Department, Faculty, Section or Course where supported. Target is empty for Global and otherwise the exact ID/code from the input. A section target looks like DS-S4-A.
- Faculty-specific rows replace the Global row for that same faculty rule. Other scoped preferences add to the global preferences.
- Unknown names, unsupported type/scope combinations, duplicate rows, missing rules and invalid values are rejected. Disable a rule with Enabled=NO instead of deleting it.

Default daily faculty limits are HARD: four theory periods and five total periods. Every lab/IP/PSE period counts toward total. Existing commitments count toward the limits. Faculty busy has an optional Component column. Unknown commitments count as theory conservatively; label known lab commitments as Lab. Fixed loads already exceeding a hard cap are reported before a production solve. Project work (Component = Project, JOE, JOE Lab, DOE or DOE Lab) counts as non-theory, and a day with existing project work may carry up to 6 periods in total instead of 5; more than 6 project periods on one day is rejected.

PERIODS_PER_WEEK is enabled in this input to require a complete timetable. Disable it for a diagnostic partial run with an explicit Unplaced list. Older workbooks without Constraints use backward-compatible partial placement plus the new default daily faculty caps and quality preferences.

Optional Fixed sessions columns are Activity ID, Session, Day, Period and Room. Room may be blank. FIXED_SLOT controls enforcement.

Use --check-only to validate settings and build/check the CP-SAT model without searching or exporting a timetable. This also reports fixed-load conflicts. It cannot prove that a full timetable is feasible.

Theory spreading, gap reduction, faculty balancing, home-room preference, room minimization and timing preferences are configurable. Coverage always takes priority over all soft penalties combined.

DS lab blocks are specified in the input as three consecutive periods for the whole section, alongside four theory periods. The solver has no department-specific DS exception.

Room seating-capacity checks remain outside the current scope.

## Placeholders in the output

- **`TBC-FAC-…`**: no faculty has been allocated yet in the input.
- **`TBC-ROOM-…`**: no listed room was free or suitable, so the class got a provisional room.
- **Unplaced**: a session that could not fit. The reason is shown in the Unplaced sheet and at the top of the HTML page.

Each placeholder is fixed by correcting the input workbook, not the output.
