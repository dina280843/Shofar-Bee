<#
  Promote-Day.ps1  —  publish a new Shofar Bee day in one step.

  What it does:
    1. Archives the current live index.html  ->  day{N-1}.html
    2. Copies the staged draft               ->  index.html   (new live day N)
    3. Rebuilds the nav on EVERY page:
         - index (day N)  : Substack + Day 1..N-1  (each "Day k")
         - day1..N-1      : Substack + Day 1..N-1 (current day = primary) + "Day N ->" index
    4. Restores any day page missing from disk from git first (guards the
       occasional "files vanished from the folder" glitch).

  Usage (run from anywhere):
    powershell -File "tools\Promote-Day.ps1" -Day 11 -Draft "DRAFT_day11_something.html"

  After it finishes: git add + commit + push, then verify live.
#>
param(
  [Parameter(Mandatory=$true)][int]$Day,
  [Parameter(Mandatory=$true)][string]$Draft
)
$ErrorActionPreference = 'Stop'
$SL   = [System.Text.RegularExpressions.RegexOptions]::Singleline
$root = Split-Path $PSScriptRoot -Parent          # the A_Shofar Bee folder
$nl   = "`r`n"; $rarr = [char]0x2192; $lhook = [char]0x21A9
$prev = $Day - 1
$utf8 = New-Object System.Text.UTF8Encoding($false)

function ReadF($p){ [System.IO.File]::ReadAllText($p) }
function WriteF($p,$t){ [System.IO.File]::WriteAllText($p,$t,$utf8) }

$indexPath = Join-Path $root 'index.html'
$draftPath = Join-Path $root $Draft

# sanity: index really is day N-1, draft really is day N
$curDay = [regex]::Match((ReadF $indexPath),'day:\s*(\d+),').Groups[1].Value
if ($curDay -ne "$prev") { throw "index.html is day $curDay, expected $prev before promoting Day $Day. Aborting." }
if (-not (Test-Path $draftPath)) { throw "Draft not found: $draftPath" }
$draftDay = [regex]::Match((ReadF $draftPath),'day:\s*(\d+),').Groups[1].Value
if ($draftDay -ne "$Day") { throw "Draft is day $draftDay, expected $Day. Aborting." }

# 1) archive current live day, 2) promote the draft
Copy-Item $indexPath (Join-Path $root ("day{0}.html" -f $prev)) -Force
Copy-Item $draftPath $indexPath -Force

# nav builders
function IndexNav {                                  # nav for the live page (day N)
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<nav class="header-nav">'+$nl)
  [void]$sb.Append('      <a href="https://theshofarbee.substack.com/" class="nav-btn" target="_blank">Substack</a>'+$nl)
  for ($k=1; $k -le $prev; $k++) { [void]$sb.Append('      <a href="day'+$k+'.html" class="nav-btn">Day '+$k+' '+$lhook+'</a>'+$nl) }
  [void]$sb.Append('    </nav>'); $sb.ToString()
}
function ArchiveNav([int]$D) {                        # nav for archive page day D
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<nav class="header-nav">'+$nl)
  [void]$sb.Append('      <a href="https://theshofarbee.substack.com/" class="nav-btn" target="_blank">Substack</a>'+$nl)
  for ($k=1; $k -le $prev; $k++) {
    $cls = if ($k -eq $D) { 'nav-btn primary' } else { 'nav-btn' }
    [void]$sb.Append('      <a href="day'+$k+'.html" class="'+$cls+'">Day '+$k+'</a>'+$nl)
  }
  [void]$sb.Append('      <a href="index.html" class="nav-btn">Day '+$Day+' '+$rarr+'</a>'+$nl)
  [void]$sb.Append('    </nav>'); $sb.ToString()
}
function SwapNav($p, $newNav) {
  $t = ReadF $p
  $old = [regex]::Match($t,'<nav class="header-nav">.*?</nav>',$SL).Value
  if ([string]::IsNullOrEmpty($old)) { throw "nav not found in $p" }
  WriteF $p ($t.Replace($old,$newNav))
}

# 3) live page nav
SwapNav $indexPath (IndexNav)
Write-Output "index.html = Day $Day (live); archived old day -> day$prev.html"

# 4) archive pages nav (restore from git if a file vanished)
for ($D=1; $D -le $prev; $D++) {
  $p = Join-Path $root ("day{0}.html" -f $D)
  if (-not (Test-Path $p)) {
    Write-Output "  day$D.html missing - restoring from git"
    & git -C $root checkout HEAD -- ("day{0}.html" -f $D) 2>&1 | Out-Null
  }
  SwapNav $p (ArchiveNav $D)
  Write-Output "  day$D.html nav -> Day $D primary, Day $Day -> index"
}
Write-Output "DONE. Next: git add index.html day1..$prev.html ; commit ; push ; verify live."
