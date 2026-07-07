<#
  Update-Words.ps1  —  replace a page's word list (and recompute maxScore).

  Pass the words as a comma/space/newline-separated string of UPPERCASE words.
  maxScore is auto-computed (4-letter word = 1 pt, else length, +7 per pangram)
  to match the site's JS, though the JS overwrites it at load anyway.

  Usage:
    powershell -File "tools\Update-Words.ps1" -File "day1.html" -Words "ARCH CARE CRUSADE ..."
#>
param(
  [Parameter(Mandatory=$true)][string]$File,
  [Parameter(Mandatory=$true)][string]$Words
)
$ErrorActionPreference = 'Stop'
$SL = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent
$path = if (Test-Path $File) { $File } else { Join-Path $root $File }
$utf8 = New-Object System.Text.UTF8Encoding($false)

$t = [System.IO.File]::ReadAllText($path)
$center = [regex]::Match($t,"center:\s*'([A-Z])'").Groups[1].Value
$outer  = [regex]::Matches([regex]::Match($t,"outer:\s*\[(.*?)\]").Groups[1].Value,"'([A-Z])'") | ForEach-Object { $_.Groups[1].Value }
$allowed = @($center) + $outer | Sort-Object -Unique

$list = ($Words -split '[,\s]+') | Where-Object { $_ -ne '' } | ForEach-Object { $_.Trim().ToUpper() } | Select-Object -Unique
# compute maxScore
$score = 0
foreach ($w in $list) {
  $score += (&{ if ($w.Length -eq 4) { 1 } else { $w.Length } })
  $isPan = ($allowed | Where-Object { $w -notmatch $_ }).Count -eq 0
  if ($isPan) { $score += 7 }
}

# wrap 8 words per line for readability
$lines = @(); for ($i=0; $i -lt $list.Count; $i += 8) {
  $chunk = $list[$i..([Math]::Min($i+7,$list.Count-1))]
  $lines += "  " + (($chunk | ForEach-Object { "'$_'" }) -join ',')
}
$newSet = "words: new Set([`r`n" + ($lines -join ",`r`n") + "`r`n]),"

$oldSet = [regex]::Match($t,'words: new Set\(\[.*?\]\),',$SL).Value
if ([string]::IsNullOrEmpty($oldSet)) { throw "word set not found in $File" }
$t = $t.Replace($oldSet, $newSet)
$oldMax = [regex]::Match($t,'maxScore: \d+,').Value
$t = $t.Replace($oldMax, "maxScore: $score,")
[System.IO.File]::WriteAllText($path, $t, $utf8)
Write-Output ("$File updated -> {0} words, maxScore={1} (center {2})" -f $list.Count, $score, $center)
