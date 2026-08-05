<#
  Remove-Banner.ps1  -  strip the "new home" banner from every published page.

  It went up on 29 June 2026 to move readers off Netlify onto GitHub Pages,
  and was always meant to be temporary. By August it was telling people to
  bookmark the site they were already reading.

  Removes all four pieces so nothing is left behind: the CSS comment, the
  .move-banner* rules, the markup, and the dismissBanner() script (with its
  sb_banner_dismissed localStorage key).

  Only touches published pages - day*.html, index.html, farewell.html.
  Untracked DRAFT_*.html files are left alone.

  Usage:
    powershell -File "tools\Remove-Banner.ps1" [-WhatIfOnly]
#>
param([switch]$WhatIfOnly)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object System.Text.UTF8Encoding($false)

$targets = @()
$targets += Get-ChildItem $root -Filter 'day*.html'
foreach ($extra in 'index.html','farewell.html') {
  $p = Join-Path $root $extra
  if (Test-Path $p) { $targets += Get-Item $p }
}

$patterns = @(
  '/\*[^*/]*NEW-HOME BANNER[^*/]*\*/\s*',                       # css comment
  '\.move-banner[^{]*\{[^}]*\}\s*',                             # all .move-banner* rules
  '<!--[^>]*NEW-HOME BANNER[^>]*-->\s*',                        # html comment
  '(?s)<div class="move-banner".*?</div>\s*',                   # the banner itself
  '(?s)<script>\s*function dismissBanner\(\).*?</script>\s*'    # its dismiss script
)

$changed = 0
foreach ($f in $targets) {
  $t = [System.IO.File]::ReadAllText($f.FullName)
  $before = $t
  foreach ($p in $patterns) { $t = [regex]::Replace($t, $p, '') }
  if ($t -ne $before) {
    if (-not $WhatIfOnly) { [System.IO.File]::WriteAllText($f.FullName, $t, $utf8) }
    $changed++
  }
  foreach ($leftover in 'move-banner','dismissBanner','sb_banner_dismissed') {
    if ($t -match [regex]::Escape($leftover)) { Write-Warning ("{0}: '{1}' still present" -f $f.Name, $leftover) }
  }
}
Write-Output ("banner removed from {0} of {1} pages{2}" -f $changed, $targets.Count, $(if ($WhatIfOnly) { ' (dry run)' } else { '' }))
