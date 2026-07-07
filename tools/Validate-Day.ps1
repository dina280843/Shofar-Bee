<#
  Validate-Day.ps1  —  sanity-check a Shofar Bee puzzle page (or draft).

  Confirms: every word contains the centre letter; every word uses ONLY the
  puzzle's 7 letters; the pangram (revealWords highlight) is in the word set;
  reports word count and any pangrams (words using all 7 letters).

  Usage:
    powershell -File "tools\Validate-Day.ps1" -File "DRAFT_day11_x.html"
    powershell -File "tools\Validate-Day.ps1" -File "index.html"
#>
param([Parameter(Mandatory=$true)][string]$File)
$ErrorActionPreference = 'Stop'
$SL = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent
$path = if (Test-Path $File) { $File } else { Join-Path $root $File }
$t = [System.IO.File]::ReadAllText($path)

$center = [regex]::Match($t,"center:\s*'([A-Z])'").Groups[1].Value
$outer  = [regex]::Matches([regex]::Match($t,"outer:\s*\[(.*?)\]").Groups[1].Value,"'([A-Z])'") | ForEach-Object { $_.Groups[1].Value }
$allowed = @($center) + $outer | Sort-Object -Unique
$highlight = [regex]::Match($t,'if \(w === "([A-Z]+)"\)').Groups[1].Value
$setBlock = [regex]::Match($t,'words: new Set\(\[(.*?)\]\)',$SL).Groups[1].Value
$words = [regex]::Matches($setBlock,"'([A-Z]+)'") | ForEach-Object { $_.Groups[1].Value }

Write-Output ("center={0}  allowed={1}  highlight={2}  words={3}" -f $center, ($allowed -join ''), $highlight, $words.Count)

$problems = 0
foreach ($w in $words) {
  if ($w -notmatch $center) { Write-Output "  NO-CENTER: $w"; $problems++ }
  $bad = ($w.ToCharArray() | Where-Object { $allowed -notcontains "$_" }) | Select-Object -Unique
  if ($bad) { Write-Output ("  BAD-LETTER: {0} (has {1})" -f $w, ($bad -join ',')); $problems++ }
}
if ($words -notcontains $highlight) { Write-Output "  HIGHLIGHT '$highlight' not in word set!"; $problems++ }

$pangrams = $words | Where-Object { $wu = $_; ($allowed | Where-Object { $wu -notmatch $_ }).Count -eq 0 }
Write-Output ("pangrams: {0}" -f ($pangrams -join ', '))
if ($problems -eq 0) { Write-Output "OK - no problems." } else { Write-Output "$problems PROBLEM(S) found." }
