# Changelog

## 2026-11-22 (morning) — v78l timetable published (in-run snapshot, all periods placed)

**Published:** `output/timetable_nov22_v78l.xlsx` / `.html`, `index.html` (latest) and `2026-11-22_1025/index.html` (fixed copy, 22 Nov 10:25 IST). Server run (16 workers), warm start from v71; this is the 15-minute save at 11,619 s (solution #1245). The run continues to about 11:54 IST and may improve slightly.

**Input `input/workload_all_departments_R24_nov22_v78l.xlsx`** = v71 input plus:
- Mech: HT lab in two batches, CAE one full-class session within two consecutive days, PSE unpinned, II-B earlier teachers, Baba/Chalapati Mon-Wed P2.
- EEE: Project and Risk Management (EMD4205) only at P4/P6/P7.
- Faculty busy has a Room column (existing load filled from the first-year timetable and department files; MBA/BBA/Pharmacy/MCA left blank).
- Weights: teacher building change back-to-back 20000; class next to existing load in another building 20000; two labs in one morning 15000; second technical lab in a day 5000; ECE lab batches same room 10000 (x2); more than 3 periods in a row 5000. Teacher floor rule off.

| Lower is better | v64 | v71 | v78l |
|---|---:|---:|---:|
| Teacher building changes (back-to-back) | 9 | 9 | 4 |
| Class next to existing load, other building | 53 | 57 | 14 |
| Class next to existing load, total | 84 | 93 | 34 |
| Teacher periods beyond 3 in a row | 94 | 105 | 64 |
| Sections with 2 labs in one morning | 22 | 24 | 15 |
| Sections with 2 technical labs in one day | 44 | 46 | 36 |
| ECE lab batches in different rooms | 9 | 2 | 1 |
| Same subject twice in a day | 85 | 130 | 116 |
| Teacher floor changes (same building) | 124 | 106 | 119 |

Remaining two-morning-lab days are CSE-C1 S4 (6), CSE-C2 S4 (6) and ECE S4 (3); these sections are fully packed.
Solver: `timetable_solver/solve_timetable_v71.py` (v70 + existing-load gap compares the real building via the Faculty busy Room column).

## 2026-11-21 (afternoon) — v64 timetable published (FEASIBLE, 2922/2922)

**Published:** `output/timetable_nov20_v64.xlsx` / `.html`, `index.html` (latest) and `2026-11-21_1405/index.html` (fixed copy, published 21 Nov 14:05 IST). Server run, 16 workers, warm start from v60; saved at 10021 s. All 2922 periods placed, 0 unplaced, 0 TBC rooms.

**Input `input/workload_all_departments_R24_nov20_v64.xlsx`** = v60 plus:
- IP-II (EVA2201) 2 -> 3 periods for AI, AIML, CSE-C1, CSE-C2, CSE-C3 and DS S4 (all IP is now 3).
- CSE C1 20-11 file: CN theory III-A 10616, CN Lab III-A 10616, CN Lab III-D 10678, DBMS Lab II-C 10656 (CN labs stay 2).
- Civil S4 English for Professionals: A 7303 Priyanka Reddy, B 11553 V Padmaja (her MBA load added).
- English through Theatre Arts (CS/IT S4): both English teachers per section, theory and lab.
- No floor rule.

| Lower is better | v35 | v60 | v64 |
|---|---:|---:|---:|
| Same subject twice in a day | 353 | 83 | 85 |
| Teacher building changes (back-to-back) | 12 | 2 | 9 |
| Section building changes | 19 | 8 | 7 |
| Section room changes | 297 | 283 | 275 |
| Student gaps | 115 | 94 | 108 |
| Teacher floor changes (same building) | — | 143 | 124 |

CSE S4 sections are now at 34 of 35 periods a week (Friday is training), so there is almost no slack there.
Solver: `timetable_solver/solve_timetable_v60.py` (v59 + `--free-depts` / `--free-faculty` to re-solve only some departments with the rest frozen to a previous timetable).

## 2026-11-21 — v60 timetable published (FEASIBLE, 2888/2888); v63 input and solver v59

**Published:** `output/timetable_nov20_v60.xlsx` / `.html`, also `index.html` (always the latest) and `2026-11-21_0916/index.html` (fixed copy, published 21 Nov 09:16 IST; `2026-11-21/` shows the same). Solved on the server (16 workers), saved at 4614 s, warm-started from v58. All 2888 periods placed, 0 unplaced, 0 TBC rooms.

| Lower is better | v35 | v58 | v60 |
|---|---:|---:|---:|
| Same subject twice in a day | 339 | 263 | 83 |
| Teacher building changes (back-to-back) | 11 | 5 | 2 |
| Section building changes | 19 | 11 | 8 |
| Section room changes | 280 | 274 | 283 |
| Student gaps | 121 | 96 | 94 |

**Input `input/workload_all_departments_R24_nov20_v60.xlsx`** (what v60 was solved from) = v35 plus:
- Swathi 12040 P1–P4 only; Mech S4 ATD theory 4; Mech S6 HT and CAE-II labs 3+3 (CAE-II theory row removed).
- EEE ASDC (ESE3202) trainers 9303 Koteshwara Rao and 3259 Kamalakar, both slots.
- CSE-C2 S4 Python Lab A-434; ESIoT theory III-E 2501, III-F 2074; CSE CN labs 2 periods.
- PSE-II venues: CSE C1 A-302, C2 A-402, C3 A-217; CS/IT H-811; AI/AIML/DS FAB Zone C; ECE/ECM/EEE Zone B; Civil/Mech Zone A or B. AIML PSE one teacher per section.
- No first hour (P1): 12029, 12008, 12001, 6002, 1029, 1093, 2507, 2491, 3073, 11018, 4429, 7061; 19156 Pallavi Badry Monday only (plus existing 3034, 4463).
- Weights: high 5000 (building changes, same subject consecutive, theory room switch), medium 1000 (faculty > 5/day, student gaps, home room, distinct rooms), low 100; Mech/Civil S6 P5–P6 idle 10000; same subject max 2/day hard, second period of a subject in a day soft 5000.

**Input `input/workload_all_departments_R24_nov20_v63.xlsx`** (not yet solved) = v60 plus:
- CSE C1 20-11 file: CN theory III-A 10616, CN Lab III-A 10616, CN Lab III-D 10678, DBMS Lab II-C 10656 (CN labs stay 2 periods).
- Civil S4 English for Professionals: A 7303 Priyanka Reddy, B 11553 V Padmaja (her MBA load added as existing load).
- English through Theatre Arts (CS/IT S4): both listed English teachers per section, theory and lab.
- Floor rule: back-to-back periods on different floors of the same building cost 5000, teachers and sections (v60 has 143 teacher / 162 section floor changes).

**Solver `timetable_solver/solve_timetable_v59.py`:** saves the first timetable at once (`<name>_first`) and every 15 min if improved; Ctrl+C stops and saves the best. Adds PREFER_ONE_SUBJECT_PER_DAY, NO_SECTION_ROOM_SWITCH, FACULTY_BACK_TO_BACK, NO_FACULTY_FLOOR_HOP, NO_SECTION_FLOOR_HOP. HTML shows sections taught together as one cell (AIML III-A/B/C) and no course codes in elective tags.
Run: `python solve_timetable_v59.py --input ..\input\workload_all_departments_R24_nov20_v63.xlsx --hint ..\output\timetable_nov20_v60.xlsx --seconds 14400 --workers <cores> --out <dir> --name timetable_nov20_v63`

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
