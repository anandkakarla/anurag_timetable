# Changelog

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
