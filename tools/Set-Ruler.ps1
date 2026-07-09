<#
  Set-Ruler.ps1  —  rewrite the hardcoded rank-tick placeholders so they match
  what the JS computes for this puzzle (0 / 10% / 25% / 42% / 58% / 75% / Queen).

  The site's JS recomputes the ruler at load anyway, but the baked-in HTML values
  are what show before JS runs (and in static previews). Run this on any draft
  after setting its word list so the placeholder is never stale.

  Usage:
    powershell -File "tools\Set-Ruler.ps1" -File "DRAFT_day12_ketubah.html"
#>
param([Parameter(Mandatory=$true)][string]$File)
$ErrorActionPreference = 'Stop'
$SL = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent
$path = if (Test-Path $File) { $File } else { Join-Path $root $File }
$utf8 = New-Object System.Text.UTF8Encoding($false)
$AZ = [System.MidpointRounding]::AwayFromZero   # match JS Math.round (.5 rounds up)

$t = [System.IO.File]::ReadAllText($path)
$center = [regex]::Match($t,"center:\s*'([A-Z])'").Groups[1].Value
$outer  = [regex]::Matches([regex]::Match($t,"outer:\s*\[(.*?)\]").Groups[1].Value,"'([A-Z])'") | ForEach-Object { $_.Groups[1].Value }
$allowed = @($center) + $outer | Sort-Object -Unique
$setBlock = [regex]::Match($t,'words: new Set\(\[(.*?)\]\)',$SL).Groups[1].Value
$words = [regex]::Matches($setBlock,"'([A-Z]+)'") | ForEach-Object { $_.Groups[1].Value }

# total score: 4-letter word = 1 pt, else length; +7 for a pangram (uses all 7 letters)
$score = 0
foreach ($w in $words) {
  $score += (&{ if ($w.Length -eq 4) { 1 } else { $w.Length } })
  if (($allowed | Where-Object { $w -notmatch $_ }).Count -eq 0) { $score += 7 }
}
function Thr($p) { [int][Math]::Round($score * $p, 0, $AZ) }

$star = [char]0x2B50
$ticks = @(
  @('Beginner', 0),
  @('Good',    (Thr 0.10)),
  @('Solid',   (Thr 0.25)),
  @('Great',   (Thr 0.42)),
  @('Amazing', (Thr 0.58)),
  @('Genius',  (Thr 0.75)),
  @(($star + ' Queen'), 'All!')
)
$nl = "`r`n"
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('<div class="rank-ticks" id="rankTicks">' + $nl)
foreach ($tk in $ticks) {
  [void]$sb.Append('      <span class="rank-tick">' + $tk[0] + '<br>' + $tk[1] + '</span>' + $nl)
}
[void]$sb.Append('    </div>')
$newBlock = $sb.ToString()

$old = [regex]::Match($t,'<div class="rank-ticks" id="rankTicks">.*?</div>',$SL).Value
if ([string]::IsNullOrEmpty($old)) { throw "rank-ticks block not found in $File" }
$t = $t.Replace($old, $newBlock)
[System.IO.File]::WriteAllText($path, $t, $utf8)
Write-Output ("$File ruler placeholders -> 0 / {0} / {1} / {2} / {3} / {4} / All!  (Queen={5})" -f (Thr 0.10),(Thr 0.25),(Thr 0.42),(Thr 0.58),(Thr 0.75),$score)
