# Overnight run for the Nov 16 input. Usage:  .\run_overnight.ps1        (about 4 hours at most)
#   Stage 1: relaxed rounds (place as much as possible, quality rules off) until every period is placed.
#   Stage 2: full rules, starting from the stage 1 result.
# Results: timetable_nov16.xlsx / .html (final) or the last stage1_roundN files if stage 2 finds nothing.
param([int]$Rounds = 5, [int]$RoundSeconds = 1500, [int]$FinalSeconds = 5400)
$ErrorActionPreference = 'Continue'
Set-Location $PSScriptRoot
$py = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
if (-not (Test-Path $py)) { $py = 'python' }
$deps = Join-Path $PSScriptRoot '..\support\semester_project\python_deps'
if (Test-Path $deps) { $env:PYTHONPATH = (Resolve-Path $deps).Path }
$workers = [Math]::Max(2, [Environment]::ProcessorCount - 1)
$log = Join-Path $PSScriptRoot 'overnight_log.txt'
"Started $(Get-Date)  workers=$workers" | Out-File $log -Encoding utf8
$hint = 'timetable_v2_2317.xlsx'
$complete = $false
for ($i = 1; $i -le $Rounds; $i++) {
    $name = "stage1_round$i"
    "--- Stage 1 round $i  $(Get-Date)  hint=$hint" | Out-File $log -Append -Encoding utf8
    $out = & $py solve_timetable.py --input workload_all_departments_nov16_relaxed.xlsx --seconds $RoundSeconds --workers $workers --name $name --hint $hint 2>&1
    $out | Out-File "$name.log" -Encoding utf8
    $status = $out | Where-Object { $_ -match '"status"' } | Select-Object -Last 1
    "$status" | Out-File $log -Append -Encoding utf8
    if (Test-Path "$name.xlsx") { $hint = "$name.xlsx" }
    if ("$status" -match '"unplaced_periods": 0\b') { $complete = $true; break }
}
"--- Stage 2 (full rules)  $(Get-Date)  hint=$hint  stage1 complete=$complete" | Out-File $log -Append -Encoding utf8
$out = & $py solve_timetable.py --input workload_all_departments_nov16.xlsx --seconds $FinalSeconds --workers $workers --name timetable_nov16 --hint $hint 2>&1
$out | Out-File 'timetable_nov16.log' -Encoding utf8
$status = $out | Where-Object { $_ -match '"status"' } | Select-Object -Last 1
"$status" | Out-File $log -Append -Encoding utf8
"Finished $(Get-Date)" | Out-File $log -Append -Encoding utf8
Get-Content $log
