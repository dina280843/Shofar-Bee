<#
  Set-Teaser.ps1  -  rewrite the hero heading and the teaser line on existing
  puzzle pages, from a JSON file.

  The early puzzles were captioned rather than teased ("Israel Museum . June
  2026 - March 2027"), and Day 2's heading printed its own answer. That barely
  showed on the site, where each day stands alone, but it is conspicuous in a
  bound book where they sit next to the later pages.

  Refuses to write a teaser that contains the puzzle's own pangram, so this
  can never reintroduce the Day 2 problem.

  Prose lives in the UTF-8 JSON (HTML entities are fine and render correctly);
  this file stays ASCII because PowerShell 5.1 reads a BOM-less .ps1 as ANSI.

  Usage:
    powershell -File "tools\Set-Teaser.ps1" -DataFile "tools\teasers_1-6.json"
#>
param(
  [string]$DataFile = 'tools\teasers_1-6.json',
  [switch]$WhatIfOnly
)
$ErrorActionPreference = 'Stop'
$SL   = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object System.Text.UTF8Encoding($false)   # match the day pages

$path = if ([System.IO.Path]::IsPathRooted($DataFile)) { $DataFile } else { Join-Path $root $DataFile }
$items = [System.IO.File]::ReadAllText($path) | ConvertFrom-Json

foreach ($it in $items) {
  $file = Join-Path $root ("day{0}.html" -f $it.day)
  if (-not (Test-Path $file)) { Write-Warning ("missing day{0}.html" -f $it.day); continue }
  $t = [System.IO.File]::ReadAllText($file)

  # guard: never let the answer appear in the heading or the teaser
  $pan = [regex]::Match($t, 'if \(w === "([A-Z]+)"\)').Groups[1].Value
  $plain = [System.Net.WebUtility]::HtmlDecode($it.hero + ' ' + $it.credit)
  if ($pan -and $plain -match ('(?i)\b' + [regex]::Escape($pan) + '\b')) {
    throw ("day{0}: the new text contains the pangram {1}" -f $it.day, $pan)
  }
  if ($it.pangram -and $pan -and $it.pangram -ne $pan) {
    throw ("day{0}: expected pangram {1} but the page says {2}" -f $it.day, $it.pangram, $pan)
  }

  $oldHero   = [regex]::Match($t, '<h2 class="hero-theme">(.*?)</h2>', $SL).Groups[1].Value
  $oldCredit = [regex]::Match($t, '<p class="hero-credit">(.*?)</p>', $SL).Groups[1].Value

  $t = [regex]::Replace($t, '(<h2 class="hero-theme">).*?(</h2>)', ('${1}' + $it.hero   + '${2}'), $SL)
  $t = [regex]::Replace($t, '(<p class="hero-credit">).*?(</p>)',  ('${1}' + $it.credit + '${2}'), $SL)

  if (-not $WhatIfOnly) { [System.IO.File]::WriteAllText($file, $t, $utf8) }
  Write-Output ("day{0}: hero  [{1}] -> [{2}]" -f $it.day, $oldHero, $it.hero)
  Write-Output ("        teaser was [{0}]" -f $oldCredit)
}
Write-Output ("{0} pages updated{1}" -f $items.Count, $(if ($WhatIfOnly) { ' (dry run)' } else { '' }))
