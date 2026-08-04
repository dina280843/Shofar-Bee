<#
  Export-PuzzleData.ps1  -  read every dayN.html and emit one JSON file with
  everything the print book needs. The web pages stay the single source of
  truth; nothing is retyped.

  Emits (per puzzle): day, date, hero, credit, card, cite, center, outer[],
  pangram, words[], queen.

  HTML entities are decoded to real characters here, because Word wants the
  character, not the entity.

  Usage:
    powershell -File "tools\Export-PuzzleData.ps1" -Out "tools\book_data.json"
#>
param(
  [string]$Out = 'tools\book_data.json',
  [int]$From = 1,
  [int]$To   = 50
)
$ErrorActionPreference = 'Stop'
$SL   = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent
# BOM on purpose: PowerShell 5.1's Get-Content falls back to the machine's
# ANSI codepage (Hebrew here) for BOM-less files and mangles every dash and
# quote. ReadAllText and ConvertFrom-Json both strip the BOM harmlessly.
$utf8 = New-Object System.Text.UTF8Encoding($true)

function Grab($text, $pattern, $group = 1) {
  $m = [regex]::Match($text, $pattern, $SL)
  if (-not $m.Success) { return '' }
  [System.Net.WebUtility]::HtmlDecode($m.Groups[$group].Value).Trim()
}

$all = @()
for ($n = $From; $n -le $To; $n++) {
  $path = Join-Path $root "day$n.html"
  if (-not (Test-Path $path)) { Write-Warning "missing day$n.html"; continue }
  $t = [System.IO.File]::ReadAllText($path)

  $outer = @()
  $om = [regex]::Match($t, "outer:\s*\[(.*?)\]", $SL)
  foreach ($g in [regex]::Matches($om.Groups[1].Value, "'([A-Z])'")) { $outer += $g.Groups[1].Value }

  $words = @()
  $wm = [regex]::Match($t, 'words: new Set\(\[(.*?)\]\)', $SL)
  foreach ($g in [regex]::Matches($wm.Groups[1].Value, "'([A-Z]+)'")) { $words += $g.Groups[1].Value }

  $all += [pscustomobject]@{
    day     = [int](Grab $t 'day:\s*(\d+),')
    date    = Grab $t "date:\s*'([^']*)'"
    hero    = Grab $t '<h2 class="hero-theme">(.*?)</h2>'
    credit  = Grab $t '<p class="hero-credit">(.*?)</p>'
    card    = Grab $t '<div class="theme-card">.*?<p>(.*?)</p>'
    cite    = Grab $t '<cite>(.*?)</cite>'
    center  = Grab $t "center:\s*'([A-Z])'"
    outer   = $outer
    pangram = Grab $t 'if \(w === "([A-Z]+)"\)'
    words   = ($words | Sort-Object)
    queen   = [int](Grab $t 'maxScore:\s*(\d+),')
  }
}

$outPath = if ([System.IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }
[System.IO.File]::WriteAllText($outPath, ($all | ConvertTo-Json -Depth 5), $utf8)

$tw = ($all | ForEach-Object { $_.words.Count } | Measure-Object -Sum).Sum
Write-Output ("exported {0} puzzles, {1} answer words -> {2}" -f $all.Count, $tw, $Out)
