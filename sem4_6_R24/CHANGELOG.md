# Changelog

## 2026-11-19 (late morning): Civil rooms and missing course (v13)

`input/workload_all_departments_R24_nov19_v13_noICS.xlsx`, from the Civil "New Work Load and Room Numbers" file:
- Concrete Tech lab in C-G2, Hydraulics lab in C-G5, Environmental Engg lab in C-304, IP-II in C-106, PSE-II in Fab Lab Zone-A. Before this, every Civil lab was in C-106 (58 periods a week for one room).
- Tekla is now a 3-period lab in C-106 (it was entered as theory).
- Added Civil S4 AS&G (T&P) Value Addition Course, 3 periods: A taught by 19128 and 19601, B by 1101 and 19602. The department's load file had it; our input didn't.
- Civil S6 VACR rooms set to C-301/C-302.
- Civil S4 = 34 periods, Civil S6 = 25. Each department solved on its own: both feasible.

## 2026-11-19 (morning): only P4 and P7 may be free; colour-coded grids

- `timetable_solver/solve_timetable_p4p7.py`: P1, P2, P3, P5 and P6 must be occupied for every section (P4 and P7 may be free). Exception: Mech S6 has 23 periods for 25 must-fill slots, so P5/P6 are top-priority soft there (weight 100000). Tested department by department: all feasible; AIML S4/S6 and ECE S4 were undecided, as on earlier versions.
- `input/workload_all_departments_R24_nov19_v12_noICS.xlsx`: v11 with Civil S6 VACR moved from Monday P5–P7 to Monday P1–P3, so Civil S6 (25 periods, 25 must-fill slots) can meet the rule.
- Grid colours (Excel and HTML): blue = theory, green = lab, purple = PSE/IP, orange = fixed training (Technical Training, QALR, VACR, FSJ, Monday off), shown with its time. Teacher view greys out existing first-year load. Fixed garbled characters in the HTML.
- Run: `python solve_timetable_p4p7.py --input ..\input\workload_all_departments_R24_nov19_v12_noICS.xlsx --hint <previous timetable.xlsx> --seconds 3600 --workers 8 --out <dir> --name timetable_nov19_v12_p4p7`

## 2026-11-19 — first feasible R24 Sem 4/6 timetable

**Result:** `output/timetable_nov19_v11_p1p3_labelled.xlsx`. FEASIBLE; all 3,030 periods placed. Every section has a class in P1–P3 on every working day. Training blocks are labelled in the grid. ICS for ECE S6 is not included.

**Input:** `input/workload_all_departments_R24_nov19_v11_noICS.xlsx`. `input/workload_all_departments_R24_nov19_v10.xlsx` is the same input with ECE S6 ICS still in it.

### Cut to make the timetable fit
| Area | Change |
|---|---|
| CSE C1/C2/C3 S4 (14 sections) | DBMS, Python and Web Tech labs: 3 → 2 periods |
| Mech S4 | MP, ATD, FM and CAD labs: 4 → 3 periods |
| ECE S6 | DSP and ESIoT labs: 3 → 2 periods |
| ECE S6 + ECM | ICS (Intro to Cyber Security) parked. PE-II and the two ICS common slots need 9 common periods; only 7 exist. |

### Input corrections
- CSE-C2 S6: the two electives are now two whole-section classes (was one pooled slot).
- CSE-C3 S6 R/T/V: added Professional Excellence Through Soft Skills (EAEXX08). The English department's file lists it for sections "A/B/C", read as R/T/V.
- ECE S6: PE-II is pooled across A–G (2 Analog VLSI, 1 RTOS, 2 CNIP). H gets CNIP with 2072.
- ECE S6 ICS (in v10): Set 1 = A+G and C+F; Set 2 = B, E+H and D+ECM.
- EEE S6 ASDC: A+B on Wed P1–P3; C on Wed P5–P7.

### Faculty swaps
- Mech S4 IP-II A: 4370 → 4463.
- Mech S4 ATD Lab A: 4433 → 4435.
- AIML S6 MERN lab: Navatha (J → F) ↔ Meghana 12083 (F → J).
- CSE-C3 S6 TBCS: 7311 (P → N) ↔ 7226 (N → P).

### Faculty slot rules
- No P1: 3034 (Dinesh, EEE) and 4463 (Naga Kishore, Mech).
- P2 only, Mon–Wed: 4001 (Chalapathi) and 4002 (Sikindar Baba).

### Solver changes
- Theory rows accept a `Parallel group`, so several merged classes can be held in one common slot.
- `Merged classes` accepts a section from another department (`D+ECM.A`).
- A merged group with no faculty on a row is skipped; it no longer becomes a TBC activity.
- The Excel grid shows `Section busy` blocks (Technical Training, Monday off) as grey "(fixed)" cells.
- New `--redraw XLSX`: rebuilds the outputs from an existing timetable without solving.
- Variants: `solve_timetable_p1p3.py` (P1–P3 must be occupied), `_p1p3p6.py` (adds P6), `_p1p3p5p6.py` (adds P5 and P6). Mech S6 has 23 periods for 25 must-fill slots, so the last variant is infeasible until the CAE-II lab is added.

### Open
- ECE S6 ICS: run it on Monday, or re-time a VACR/FSJ block.
- Mech S6: the Computer Aided Engineering-II lab (2 periods in the grid) is missing from the input.
- No-P1 rule: IDs still needed for "Jyothi madam" (Mech) and one other faculty member.
- 92 activities have TBC faculty and 94 have TBC rooms. Paired S/W–H/W labs still book only one room.

## 2026-11-20 — v35 timetable published (FEASIBLE, 2897/2897 periods)
- Solver `solve_timetable_v38b.py`: department room lists are hard (TBC room only where none was given);
  building-hop penalties for teachers and sections; per-day / half-day room-change penalties for floating
  sections; existing-load gap penalty; PSE-II venues shared two sections at a time; TBCS only in P4/P6/P7;
  at most 3 consecutive periods per teacher (soft); optional `--soft-rooms`; quieter console output.
- Inputs: v35 (published timetable) and v38 (latest: + Swathi 12040 mornings only, V. Jyothi 1008 no P1,
  max-3-consecutive rule). Home rooms: dedicated rooms as given; pooled departments assigned in list order,
  AIML 70% fixed / 30% floating. IP-II in department computer labs; PSE-II in the OIA venue list.
- Tools: run_overnight.py (attempt ladder), compare_runs.py (quality table), diagnose_rooms.py (room bottleneck).
