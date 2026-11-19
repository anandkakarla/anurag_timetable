# Usage:  .\run_timetable.ps1                 (600 s, timetable_input.xlsx)
#         .\run_timetable.ps1 -Seconds 300
#         .\run_timetable.ps1 -CheckOnly      (validate the input only)
param([int]$Seconds = 600, [string]$InputFile = "timetable_input.xlsx", [string]$Name = "timetable_output", [switch]$CheckOnly)
$ErrorActionPreference = 'Stop'
$py = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
if (-not (Test-Path $py)) { $py = 'python' }
$deps = Join-Path $PSScriptRoot '..\support\semester_project\python_deps'
if (Test-Path $deps) { $env:PYTHONPATH = (Resolve-Path $deps).Path }
$script = Join-Path $PSScriptRoot 'solve_timetable.py'
$inp = Join-Path $PSScriptRoot $InputFile
$args2 = @($script, '--input', $inp, '--seconds', $Seconds, '--name', $Name)
if ($CheckOnly) { $args2 += '--check-only' }
& $py @args2
if ($LASTEXITCODE -ne 0) { throw "Timetable run failed - see messages above." }
