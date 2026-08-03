<#
  New-Day.ps1  -  build one or more dayN.html pages from a template + a JSON
  data file, filling every per-day slot. Replaces the old throwaway
  gen*.ps1 scratch scripts.

  The JSON is an ARRAY of objects, one per puzzle:
    {
      "day":    44,
      "date":   "August 16, 2026",
      "file":   "day44.html",
      "theme":  "Blossom Before Spring - Almonds",   // PUZZLE.theme
      "hero":   "Blossom Before Spring",             // <h2 class=hero-theme>
      "credit": "One sentence teasing the answer.",  // <p class=hero-credit>
      "card":   "The Today's-Inspiration paragraph.",
      "cite":   "- Numbers 17:8",
      "center": "D",
      "outer":  ["A","L","M","O","N","S"],
      "pangram":"ALMONDS",                           // highlighted on reveal
      "words":  "ADDS ALMOND ALMONDS ..."            // space/comma separated
    }

  Keep this file ASCII-only: PowerShell 5.1 reads a BOM-less .ps1 as ANSI and
  would mangle any typographic characters. Anything non-ASCII (the middot and
  em dash in the hero CTA) is preserved FROM THE TEMPLATE by capturing it in a
  regex group rather than being typed here. Put accented / typographic text in
  the JSON, which is read as UTF-8.

  Usage:
    powershell -File "tools\New-Day.ps1" -DataFile "new_days.json" -Template "day42.html"
#>
param(
  [Parameter(Mandatory=$true)][string]$DataFile,
  [Parameter(Mandatory=$true)][string]$Template
)
$ErrorActionPreference = 'Stop'
$SL   = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Resolve-In($p) { if (Test-Path $p) { (Resolve-Path $p).Path } else { Join-Path $root $p } }

$tplPath = Resolve-In $Template
if (-not (Test-Path $tplPath)) { throw "template not found: $Template" }
$tpl = [System.IO.File]::ReadAllText($tplPath)

$json = [System.IO.File]::ReadAllText((Resolve-In $DataFile))
$days = $json | ConvertFrom-Json
if ($days -isnot [array]) { $days = @($days) }

function Esc-Js($s) { $s -replace '\\', '\\' -replace "'", "\'" }

foreach ($d in $days) {
  $t = $tpl

  # --- hero -----------------------------------------------------------------
  $t = [regex]::Replace($t, '(<h2 class="hero-theme">).*?(</h2>)', ('${1}' + $d.hero + '${2}'), $SL)
  $t = [regex]::Replace($t, '(<p class="hero-credit">).*?(</p>)',  ('${1}' + $d.credit + '${2}'), $SL)
  # keep the template's own middot / em dash / arrow by capturing them
  # \s* before the entity is consumed and one space re-emitted, so spacing is
  # identical whether or not the source already had it
  $t = [regex]::Replace($t, '(class="hero-cta">Day )\d+( . )[^&]*?\s*(&nbsp;)',
                        ('${1}' + $d.day + '${2}' + $d.date + ' ${3}'))

  # --- Today's Inspiration card --------------------------------------------
  $t = [regex]::Replace($t, '(<div class="theme-card">.*?<p>).*?(</p>)', ('${1}' + $d.card + '${2}'), $SL)
  $t = [regex]::Replace($t, '(<cite>).*?(</cite>)', ('${1}' + $d.cite + '${2}'), $SL)

  # --- the printed rule (does NOT auto-update from PUZZLE.center) -----------
  $t = [regex]::Replace($t, '(<strong>center letter )[A-Z](</strong>)', ('${1}' + $d.center + '${2}'))

  # --- PUZZLE config --------------------------------------------------------
  $outer = ($d.outer | ForEach-Object { "'$_'" }) -join ', '
  $t = [regex]::Replace($t, "day:\s*\d+,",            "day: $($d.day),")
  $t = [regex]::Replace($t, "date:\s*'[^']*',",       "date: '$(Esc-Js $d.date)',")
  $t = [regex]::Replace($t, "theme:\s*'[^']*',",      "theme: '$(Esc-Js $d.theme)',")
  $t = [regex]::Replace($t, "center:\s*'[A-Z]',",     "center: '$($d.center)',")
  $t = [regex]::Replace($t, "outer:\s*\[[^\]]*\],",   "outer: [$outer],")

  # --- pangram highlight on reveal -----------------------------------------
  $t = [regex]::Replace($t, 'if \(w === "[A-Z]+"\)', "if (w === `"$($d.pangram)`")")

  $outPath = Join-Path $root $d.file
  [System.IO.File]::WriteAllText($outPath, $t, $utf8)

  # --- word list + maxScore, then the baked-in ruler numbers ----------------
  & (Join-Path $PSScriptRoot 'Update-Words.ps1') -File $outPath -Words $d.words | Out-Null
  & (Join-Path $PSScriptRoot 'Set-Ruler.ps1')    -File $outPath | Out-Null

  $n = (($d.words -split '[,\s]+') | Where-Object { $_ -ne '' }).Count
  Write-Output ("built {0,-12} day {1,-3} {2,-18} centre {3}  {4} words" -f $d.file, $d.day, $d.date, $d.center, $n)
}
