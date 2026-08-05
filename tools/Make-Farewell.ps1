<#
  Make-Farewell.ps1  -  build farewell.html, the closing note that goes live
  the day after the last puzzle, and repair the dead Substack link that sits
  in the footer of every page.

  The page is made by copying day50.html and cutting out the game, so it
  inherits the site's CSS, header, nav and footer exactly. Nothing is
  restyled.

  Also fixes: dinaraveh974319.substack.com returns 404 and appears twice in
  every footer. The live publication is theshofarbee.substack.com, which is
  what the nav button already uses.

  ASCII-only source; typographic characters go in as HTML entities.

  Usage:
    powershell -File "tools\Make-Farewell.ps1"
#>
$ErrorActionPreference = 'Stop'
$SL   = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object System.Text.UTF8Encoding($false)

$DEAD = 'dinaraveh974319.substack.com'
$LIVE = 'theshofarbee.substack.com'

# ---- 1. repair the dead footer link everywhere -----------------------------
$fixed = 0
foreach ($f in (Get-ChildItem $root -Filter 'day*.html') + (Get-ChildItem $root -Filter 'index.html')) {
  $t = [System.IO.File]::ReadAllText($f.FullName)
  if ($t -match [regex]::Escape($DEAD)) {
    $t = $t.Replace($DEAD, $LIVE)
    [System.IO.File]::WriteAllText($f.FullName, $t, $utf8)
    $fixed++
  }
}
Write-Output ("repaired the dead Substack link on {0} pages" -f $fixed)

# ---- 2. build farewell.html from day50.html --------------------------------
$src = Join-Path $root 'day50.html'
$t = [System.IO.File]::ReadAllText($src)

$hero = @'
    <h2 class="hero-theme">Fifty Puzzles, and a Pause</h2>
    <p class="hero-credit">Thank you for playing. We are stopping here for a while &mdash; and we would like to hear from you.</p>
'@

$body = @'
<section class="game-section" id="game">

  <div class="theme-card">
    <div class="theme-card-label">From Dina and Shalom</div>
    <p>That is fifty. We began at the end of June with a queen holding Jerusalem, and we have finished with six lines of silver from a hillside above the Hinnom valley.</p>
  </div>

  <div style="max-width:640px; margin:22px auto 0; line-height:1.7; color:#3a2e1e; font-size:0.95rem;">

    <p style="margin-bottom:16px;">We are stopping here for a while. Rosh Hashana comes in on the evening of Friday 11 September, and the weeks before it are not the weeks for setting word puzzles.</p>

    <p style="margin-bottom:16px;">Two things before we go.</p>

    <p style="margin-bottom:16px;"><strong>The first is a question, and we would genuinely like an answer: would you like us to come back after the chagim?</strong> If enough of you say yes, we will think seriously about starting again once the festivals are over. Tell us on Substack &mdash; we are reading.</p>

    <p style="margin-bottom:16px;"><strong>The second:</strong> all fifty are being made into a paperback. The same hives, with room to write beside each one and every answer gathered at the back. We will post the link here the moment it exists.</p>

    <p style="margin-bottom:22px;">Meanwhile nothing disappears. Every puzzle stays exactly where it is, and you can still play all fifty &mdash; the days are listed at the top of this page.</p>

    <p style="text-align:center; margin:28px 0 6px; font-family:'Playfair Display',serif; font-style:italic; font-size:1.15rem; color:#C9A84C;">Shana tova.</p>
    <p style="text-align:center; font-size:0.9rem; color:#8B7355;">Dina &amp; Shalom</p>

    <p style="text-align:center; margin-top:26px;">
      <a href="https://theshofarbee.substack.com/" target="_blank" class="nav-btn primary" style="display:inline-block; padding:9px 20px;">Tell us on Substack</a>
    </p>
  </div>

</section>
'@

# swap the hero lines and drop the whole game section
$t = [regex]::Replace($t, '(<div class="hero-inner">).*?(</div>)', ('${1}' + "`r`n" + $hero + '  ${2}'), $SL)
$t = [regex]::Replace($t, '<section class="game-section" id="game">.*?</section>', $body, $SL)

# the puzzle engine is gone, so keep only the nav call
$t = [regex]::Replace($t,
  '(<script src="schedule\.js"></script>\s*<script>).*?(</script>)',
  ('${1}' + "`r`n" + 'if (window.sbBuildNav) sbBuildNav(51);' + "`r`n" + '${2}'), $SL)

$out = Join-Path $root 'farewell.html'
[System.IO.File]::WriteAllText($out, $t, $utf8)
Write-Output ("wrote {0} ({1:N0} bytes)" -f $out, (Get-Item $out).Length)

# ---- 3. let a schedule entry carry its own nav label -----------------------
$sf = Join-Path $root 'schedule.js'
$s = [System.IO.File]::ReadAllText($sf)
if ($s -notmatch 's\.label') {
  $s = $s.Replace(
    "h += '<a href=`"' + s.file + '`" class=`"nav-btn' + (s.day === currentDay ? ' primary' : '') + '`">Day ' + s.day + '</a>';",
    "h += '<a href=`"' + s.file + '`" class=`"nav-btn' + (s.day === currentDay ? ' primary' : '') + '`">' + (s.label || ('Day ' + s.day)) + '</a>';")
  Write-Output 'sbBuildNav now honours an optional per-entry label'
}
if ($s -notmatch 'farewell\.html') {
  $s = $s.Replace(
    "  { day: 50, date: '2026-08-23', file: 'day50.html' }`r`n];",
    "  { day: 50, date: '2026-08-23', file: 'day50.html' },`r`n  { day: 51, date: '2026-08-24', file: 'farewell.html', label: 'Fifty ' + String.fromCharCode(38) + ' a Pause' }`r`n];")
  Write-Output 'scheduled farewell.html for 24 August 2026'
}
[System.IO.File]::WriteAllText($sf, $s, $utf8)
