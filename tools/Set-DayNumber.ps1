<#
  Set-DayNumber.ps1  -  move an already-built puzzle to a different day number
  and date. Used when the running order is reshuffled: the puzzle content
  (theme, letters, words) is untouched, only its slot in the calendar changes.

  Updates: PUZZLE.day, PUZZLE.date, and the "Day N - DATE" hero CTA.
  Does NOT touch schedule.js - update that separately.

  -Out lets you write to a different filename (a reorder is usually a
  permutation, so write to temp names first, then move them into place).

  ASCII-only on purpose: PowerShell 5.1 reads a BOM-less .ps1 as ANSI. The
  middot / em dash / arrow in the CTA are preserved from the source file by
  capturing them in a regex group instead of being typed here.

  Usage:
    powershell -File "tools\Set-DayNumber.ps1" -File "day40.html" -Day 37 -Date "August 7, 2026" -Out "day37.html"
#>
param(
  [Parameter(Mandatory=$true)][string]$File,
  [Parameter(Mandatory=$true)][int]$Day,
  [Parameter(Mandatory=$true)][string]$Date,
  [string]$Out
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object System.Text.UTF8Encoding($false)

$path = if (Test-Path $File) { (Resolve-Path $File).Path } else { Join-Path $root $File }
if (-not (Test-Path $path)) { throw "not found: $File" }
$t = [System.IO.File]::ReadAllText($path)

$was = [regex]::Match($t, 'day:\s*(\d+),').Groups[1].Value

$t = [regex]::Replace($t, 'day:\s*\d+,',      "day: $Day,")
$t = [regex]::Replace($t, "date:\s*'[^']*',", "date: '$($Date -replace "'", "\'")',")
# \s* before the entity is consumed and one space re-emitted, so spacing is
# identical whether or not the source already had it
$t = [regex]::Replace($t, '(class="hero-cta">Day )\d+( . )[^&]*?\s*(&nbsp;)',
                      ('${1}' + $Day + '${2}' + $Date + ' ${3}'))

$outPath = if ($Out) { Join-Path $root $Out } else { $path }
[System.IO.File]::WriteAllText($outPath, $t, $utf8)
Write-Output ("{0}: day {1} -> {2}, {3}  =>  {4}" -f (Split-Path $path -Leaf), $was, $Day, $Date, (Split-Path $outPath -Leaf))
