<#
  Build-Book-PPT.ps1  -  lay the paperback out in PowerPoint instead of Word.

  Why PowerPoint:
   * a slide is an absolute canvas. Word's floating shapes are anchored to
     paragraphs, and getting one anchor per page wrong stacked whole pages on
     top of each other; there is no anchor to get wrong here.
   * each hive is grouped into a single object, so the shapes can be moved or
     resized as one in the PowerPoint UI.
   * PowerPoint's SaveAs(..., 32) PDF export works on this machine, while
     Word's PDF export hangs forever on anything non-trivial.

  One slide = one book page, 6 x 9 inches. Odd slides are rectos (gutter on
  the left), even slides are versos (gutter on the right), mirrored by hand
  since slides have no concept of facing pages.

  ASCII-only source: PowerShell 5.1 reads a BOM-less .ps1 as ANSI. Prose comes
  from the UTF-8 JSON; typographic characters are built from [char] codes.

  Usage:
    powershell -File "tools\Build-Book-PPT.ps1" -Sample
    powershell -File "tools\Build-Book-PPT.ps1"            # all 50 puzzles
#>
param(
  [switch]$Sample,
  [switch]$NoPdf,
  [string]$DataFile = 'tools\book_data.json',
  [string]$OutDir   = 'book'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$outDirFull = Join-Path $root $OutDir
if (-not (Test-Path $outDirFull)) { New-Item -ItemType Directory -Path $outDirFull | Out-Null }

$LogPath = Join-Path $outDirFull 'build-ppt.log'
Set-Content -Path $LogPath -Value '' -Encoding UTF8
function Log($m) {
  $line = ('[{0:HH:mm:ss}] {1}' -f (Get-Date), $m)
  Add-Content -Path $LogPath -Value $line -Encoding UTF8
  Write-Output $line
}

# ---- typography ------------------------------------------------------------
$FaceDisplay = 'Palatino Linotype'
$FaceBody    = 'Palatino Linotype'
$FaceLabel   = 'Corbel'
$FaceList    = 'Cambria'

function ToBgr($hex) {
  $h = $hex.TrimStart('#')
  [Convert]::ToInt32($h.Substring(0,2),16) + ([Convert]::ToInt32($h.Substring(2,2),16) * 256) + ([Convert]::ToInt32($h.Substring(4,2),16) * 65536)
}
$GOLD = ToBgr 'C9A84C'; $DARK = ToBgr '1E1A10'; $STONE = ToBgr '8B7355'
$PALE = ToBgr 'FDF6E3'; $BORDER = ToBgr 'E2D9C5'; $WHITE = ToBgr 'FFFFFF'

$MDASH = [char]0x2014; $MIDDOT = [char]0x00B7
$SITE  = 'https://dina280843.github.io/Shofar-Bee'

# ---- geometry, 6 x 9 in ----------------------------------------------------
$PW = 432.0; $PH = 648.0
$MT = 54.0
$MIN_ = 63.0; $MOUT = 45.0

# ---- PowerPoint constants --------------------------------------------------
$ppLayoutBlank = 12
$msoTextHoriz  = 1
$msoTrue = -1; $msoFalse = 0
$ppAlignLeft = 1; $ppAlignCenter = 2; $ppAlignRight = 3; $ppAlignJustify = 4
$ppAutoSizeNone = 0
$msoLineRoundDot = 3
$ppSaveAsPPTX = 24; $ppSaveAsPDF = 32

# ---- helpers ---------------------------------------------------------------
function New-Box($slide, $x, $y, $w, $h) {
  $s = $slide.Shapes.AddTextbox($msoTextHoriz, [single]$x, [single]$y, [single]$w, [single]$h)
  $s.Fill.Visible = $msoFalse
  $s.Line.Visible = $msoFalse
  $tf = $s.TextFrame
  $tf.MarginLeft = [single]0; $tf.MarginRight = [single]0
  $tf.MarginTop  = [single]0; $tf.MarginBottom = [single]0
  $tf.WordWrap = $msoTrue
  $tf.AutoSize = $ppAutoSizeNone
  $s
}

function Set-Box($shape, $text, $face, $size, $color, $align, [double]$lead, [bool]$bold, [bool]$italic, [double]$spacing) {
  $tr = $shape.TextFrame.TextRange
  $tr.Text = $text
  $tr.Font.Name = $face
  $tr.Font.Size = [single]$size
  $tr.Font.Color.RGB = $color
  $tr.Font.Bold   = $(if ($bold)   { $msoTrue } else { $msoFalse })
  $tr.Font.Italic = $(if ($italic) { $msoTrue } else { $msoFalse })
  $pf = $tr.ParagraphFormat
  $pf.Alignment = $align
  if ($lead -gt 0) { $pf.LineRuleWithin = $msoFalse; $pf.SpaceWithin = [single]$lead }
  $pf.SpaceBefore = [single]0; $pf.SpaceAfter = [single]0
  # letter spacing lives only on the newer TextFrame2 interface
  if ($spacing -ne 0) { $shape.TextFrame2.TextRange.Font.Spacing = [single]$spacing }
  $shape
}

function Add-Rule($slide, $x, $y, $w, $color, [double]$weight) {
  $s = $slide.Shapes.AddLine([single]$x, [single]$y, [single]($x + $w), [single]$y)
  $s.Line.ForeColor.RGB = $color
  $s.Line.Weight = [single]$weight
  $s
}

# flat-top hexagon, identical geometry to the site's hexPoints()
function Add-Hex($slide, [double]$cx, [double]$cy, [double]$R, $fill, $line) {
  $b = $slide.Shapes.BuildFreeform(1, [single]($cx + $R), [single]$cy)   # msoEditingCorner
  for ($i = 1; $i -lt 6; $i++) {
    $a = [Math]::PI / 180.0 * (60.0 * $i)
    $b.AddNodes(0, 1, [single]($cx + $R * [Math]::Cos($a)), [single]($cy + $R * [Math]::Sin($a)))
  }
  $b.AddNodes(0, 1, [single]($cx + $R), [single]$cy)
  $s = $b.ConvertToShape()
  if ($null -eq $fill) { $s.Fill.Visible = $msoFalse } else { $s.Fill.Visible = $msoTrue; $s.Fill.ForeColor.RGB = $fill }
  $s.Line.ForeColor.RGB = $line
  $s.Line.Weight = [single]1.0
  $s
}

# the whole hive, returned as ONE grouped shape
function Add-Hive($slide, [double]$cx, [double]$cy, [double]$R, $centre, $outer, [double]$fs) {
  $d = $R * [Math]::Sqrt(3)
  $angles = @(270, 330, 30, 90, 150, 210)
  $cells = @(@{x=$cx; y=$cy; ch=$centre; isC=$true})
  for ($i = 0; $i -lt 6; $i++) {
    $rad = $angles[$i] * [Math]::PI / 180.0
    $cells += @{ x = $cx + $d * [Math]::Cos($rad); y = $cy + $d * [Math]::Sin($rad); ch = $outer[$i]; isC = $false }
  }
  $names = @()
  foreach ($c in $cells) {
    $h = if ($c.isC) { Add-Hex $slide $c.x $c.y $R $GOLD $GOLD } else { Add-Hex $slide $c.x $c.y $R $PALE $GOLD }
    $names += $h.Name
  }
  foreach ($c in $cells) {
    $bw = $R * 1.5
    $t = New-Box $slide ($c.x - $bw/2) ($c.y - $fs*0.78) $bw ($fs*1.6)
    Set-Box $t $c.ch $FaceLabel $fs $(if ($c.isC) { $WHITE } else { $DARK }) $ppAlignCenter 0 $true $false 1.0 | Out-Null
    $names += $t.Name
  }
  $g = $slide.Shapes.Range([ref]$names).Group()
  $g.Name = 'Hive'
  $g
}

# make a whole text box a clickable hyperlink; survives the PDF export
function Add-Link($shape, $url) {
  $shape.TextFrame.TextRange.ActionSettings.Item(1).Hyperlink.Address = $url   # ppMouseClick
  $shape
}

# Split the answer section into printable columns.
# Returns an array of slides, each an array of 3 columns, each an array of
# lines. Nothing overflows: capacity is fixed and the packer respects it.
# A day heading is never left as the last line or two of a column.
function Build-AnswerSlides($data, [int]$rows) {
  $lines = @()
  foreach ($p in $data) {
    $lines += @{ k = 'head'; t = ('DAY ' + $p.day + '  ' + $MIDDOT + '  ' + $p.hero) }
    $lines += @{ k = 'meta'; t = ('centre ' + $p.center + '   Queen ' + $p.queen) }
    foreach ($w in $p.words) { $lines += @{ k = 'word'; t = $w } }
    $lines += @{ k = 'gap';  t = '' }
  }

  # A heading carries the day's title and nearly always wraps to two rendered
  # lines in a column this narrow, so it is budgeted as two. Counting it as
  # one let columns creep past the bottom margin.
  $slides = @(); $cols = @(); $col = @(); $cost = 0
  foreach ($ln in $lines) {
    $c = if ($ln.k -eq 'head') { 2 } else { 1 }
    # do not strand a heading at the foot of a column
    if ($ln.k -eq 'head' -and $cost -gt 0 -and ($cost + $c) -gt ($rows - 2)) {
      while ($cost -lt $rows) { $col += @{ k = 'gap'; t = '' }; $cost++ }
    }
    if (($cost + $c) -gt $rows) {
      $cols += ,$col; $col = @(); $cost = 0
      if ($cols.Count -eq 3) { $slides += ,$cols; $cols = @() }
    }
    $col += $ln; $cost += $c
  }
  if ($col.Count -gt 0) { $cols += ,$col }
  if ($cols.Count -gt 0) { $slides += ,$cols }
  ,$slides
}

function Paint-Answers($slide, $columns, [bool]$isFirst, $x0, $cw, [double]$lead, [int]$rows) {
  $title = if ($isFirst) { 'ANSWERS' } else { 'ANSWERS' + [char]0x2003 + 'continued' }
  Set-Box (New-Box $slide $x0 $MT $cw 24) $title $FaceDisplay 16 $DARK $ppAlignLeft 20 $true $false 1.5 | Out-Null
  Add-Rule $slide $x0 ($MT + 26) $cw $GOLD 1.0 | Out-Null

  $gap = 14.0
  $colW = ($cw - 2*$gap) / 3.0
  $top = $MT + 40
  $h = $rows * $lead + 2

  for ($c = 0; $c -lt $columns.Count; $c++) {
    $lines = $columns[$c]
    if (-not $lines -or $lines.Count -eq 0) { continue }
    $box = New-Box $slide ($x0 + $c*($colW+$gap)) $top $colW $h
    Set-Box $box (($lines | ForEach-Object { $_.t }) -join "`r") $FaceList 8.5 $DARK $ppAlignLeft $lead $false $false 0 | Out-Null
    $tr = $box.TextFrame.TextRange
    for ($i = 0; $i -lt $lines.Count; $i++) {
      if ($i + 1 -gt $tr.Paragraphs().Count) { break }
      $pr = $tr.Paragraphs($i + 1)
      switch ($lines[$i].k) {
        'head' { $pr.Font.Name = $FaceLabel; $pr.Font.Bold = $msoTrue; $pr.Font.Size = [single]8 }
        'meta' { $pr.Font.Italic = $msoTrue; $pr.Font.Color.RGB = $STONE; $pr.Font.Size = [single]7.5 }
      }
    }
  }
}

function Ranks($queen) {
  $AZ = [System.MidpointRounding]::AwayFromZero
  $f = @(0.10, 0.25, 0.42, 0.58, 0.75); $n = @('Good','Solid','Great','Amazing','Genius')
  $out = @('Beginner 0')
  for ($i = 0; $i -lt 5; $i++) { $out += ('{0} {1}' -f $n[$i], [int][Math]::Round($queen * $f[$i], 0, $AZ)) }
  $out + ('Queen ' + $queen)
}

# ---- page painters ---------------------------------------------------------
function Paint-Puzzle($slide, $p, $x0, $cw) {
  $y = $MT
  Set-Box (New-Box $slide $x0 $y $cw 16) ('DAY ' + $p.day) $FaceLabel 10 $STONE $ppAlignLeft 13 $true $false 2.4 | Out-Null
  Add-Rule $slide $x0 ($y + 19) $cw $BORDER 0.75 | Out-Null
  $y += 34

  Set-Box (New-Box $slide $x0 $y $cw 54) $p.hero $FaceDisplay 20 $DARK $ppAlignLeft 24 $true $false 0 | Out-Null
  $y += 60

  Set-Box (New-Box $slide $x0 $y $cw 60) $p.credit $FaceBody 10.5 $STONE $ppAlignLeft 15 $false $true 0 | Out-Null
  $y += 70

  Add-Hive $slide ($x0 + $cw/2) ($y + 104) 40 $p.center $p.outer 21 | Out-Null
  $y += 222

  Set-Box (New-Box $slide $x0 $y $cw 16) ('Every word must contain the centre letter ' + $p.center) $FaceLabel 9.5 $STONE $ppAlignCenter 13 $false $false 0.6 | Out-Null
  $y += 30

  # inspiration block: bar height follows the text, so short quotes do not
  # leave the citation stranded at the foot of the page
  $cardLead = 14.0
  $cardLines = [Math]::Max(2, [Math]::Ceiling($p.card.Length / 62.0))
  $cardH = $cardLines * $cardLead
  $blockH = 16 + $cardH + 18
  $bar = $slide.Shapes.AddLine([single]$x0, [single]$y, [single]$x0, [single]($y + $blockH))
  $bar.Line.ForeColor.RGB = $GOLD; $bar.Line.Weight = [single]3.0

  Set-Box (New-Box $slide ($x0 + 14) $y ($cw - 14) 13) 'INSPIRATION' $FaceLabel 8 $STONE $ppAlignLeft 11 $true $false 2.0 | Out-Null
  Set-Box (New-Box $slide ($x0 + 14) ($y + 16) ($cw - 14) $cardH) $p.card $FaceBody 10 $DARK $ppAlignLeft $cardLead $false $true 0 | Out-Null
  Set-Box (New-Box $slide ($x0 + 14) ($y + 16 + $cardH + 4) ($cw - 14) 13) $p.cite $FaceLabel 8.5 $STONE $ppAlignLeft 11 $false $false 0 | Out-Null
}

function Paint-Write($slide, $p, $x0, $cw) {
  Set-Box (New-Box $slide $x0 $MT $cw 16) ('DAY ' + $p.day + '  ' + $MIDDOT + '  WORDS FOUND') $FaceLabel 10 $STONE $ppAlignRight 13 $true $false 2.4 | Out-Null
  Add-Rule $slide $x0 ($MT + 19) $cw $BORDER 0.75 | Out-Null

  $gap = 12.0
  $colW = ($cw - 2*$gap) / 3.0
  $top = $MT + 40
  $rows = 34
  $lead = 13.0
  $names = @()
  for ($c = 0; $c -lt 3; $c++) {
    $cx = $x0 + $c * ($colW + $gap)
    for ($r = 0; $r -lt $rows; $r++) {
      $ln = $slide.Shapes.AddLine([single]$cx, [single]($top + $r*$lead), [single]($cx + $colW), [single]($top + $r*$lead))
      $ln.Line.ForeColor.RGB = $BORDER
      $ln.Line.Weight = [single]0.5
      $ln.Line.DashStyle = $msoLineRoundDot
      $names += $ln.Name
    }
  }
  $g = $slide.Shapes.Range([ref]$names).Group(); $g.Name = 'WritingGrid'

  $ly = $top + $rows*$lead + 16
  Add-Rule $slide $x0 $ly $cw $BORDER 0.75 | Out-Null
  Set-Box (New-Box $slide $x0 ($ly + 10) $cw 12) 'SCORE' $FaceLabel 8 $STONE $ppAlignLeft 11 $true $false 2.0 | Out-Null
  # 8pt with tight separators: at 8.5 the four-digit Queen scores wrapped
  Set-Box (New-Box $slide $x0 ($ly + 24) $cw 30) ((Ranks $p.queen) -join ('  ' + $MIDDOT + '  ')) $FaceLabel 8 $DARK $ppAlignLeft 12 $false $false 0 | Out-Null
  Set-Box (New-Box $slide $x0 ($ly + 56) $cw 12) 'A four-letter word scores 1; longer words score their length; the pangram adds 7.' $FaceBody 8 $STONE $ppAlignLeft 11 $false $true 0 | Out-Null
}

# ---- build -----------------------------------------------------------------
$dataPath = if ([System.IO.Path]::IsPathRooted($DataFile)) { $DataFile } else { Join-Path $root $DataFile }
$data = [System.IO.File]::ReadAllText($dataPath) | ConvertFrom-Json

# Answers are paginated up front so the plan knows how many slides they need.
# 46 lines at 10.5pt leaves the bottom margin clear on a 9in page.
$ansRows  = 46
$ansLead  = 10.5
$ansData   = if ($Sample) { $data | Where-Object { $_.day -le 2 } } else { $data }
$ansSlides = Build-AnswerSlides $ansData $ansRows

if ($Sample) { $plan = @('title','copyright','howto','puzzle:1','write:1','puzzle:50','write:50') }
else {
  $plan = @('title','copyright','howto','blank')
  foreach ($p in $data) { $plan += ('puzzle:' + $p.day); $plan += ('write:' + $p.day) }
}
for ($i = 0; $i -lt $ansSlides.Count; $i++) { $plan += ('answers:' + $i) }

Log ('starting PowerPoint, {0} slides planned' -f $plan.Count)
$ppt = New-Object -ComObject PowerPoint.Application
try {
  $ppt.Visible = $msoTrue           # PowerPoint refuses to run fully hidden
  $pres = $ppt.Presentations.Add($msoFalse)
  $pres.PageSetup.SlideWidth  = [single]$PW
  $pres.PageSetup.SlideHeight = [single]$PH

  for ($i = 0; $i -lt $plan.Count; $i++) {
    $spec = $plan[$i]
    $pageNo = $i + 1
    $isRecto = ($pageNo % 2 -eq 1)
    $x0 = if ($isRecto) { $MIN_ } else { $MOUT }
    $cw = $PW - $MIN_ - $MOUT
    $slide = $pres.Slides.Add($pageNo, $ppLayoutBlank)
    Log ('  slide {0}/{1}  {2}' -f $pageNo, $plan.Count, $spec)

    switch -Regex ($spec) {
      '^title$' {
        Set-Box (New-Box $slide $x0 186 $cw 42) 'THE SHOFAR BEE' $FaceDisplay 30 $DARK $ppAlignCenter 34 $true $false 3.5 | Out-Null
        Add-Rule $slide ($x0 + $cw*0.30) 244 ($cw*0.40) $GOLD 1.0 | Out-Null
        Set-Box (New-Box $slide $x0 260 $cw 24) 'Fifty Word Puzzles from Jerusalem' $FaceBody 14 $STONE $ppAlignCenter 18 $false $true 0 | Out-Null
        Add-Hive $slide ($x0 + $cw/2) 380 26 'B' @('S','H','O','F','A','R') 15 | Out-Null
        Set-Box (New-Box $slide $x0 500 $cw 20) ('Dina Raveh ' + [char]0x0026 + ' Shalom Brenner') $FaceBody 12.5 $DARK $ppAlignCenter 16 $false $false 1.2 | Out-Null
      }
      '^copyright$' {
        # the URL sits in its own box so it can carry a real hyperlink into
        # the PDF, rather than being dead printed text
        $above = @(
          'The Shofar Bee: Fifty Word Puzzles from Jerusalem', '',
          ('Copyright ' + [char]0x00A9 + ' 2026 Dina Raveh and Shalom Brenner'),
          'All rights reserved.', '',
          'The puzzles in this book first appeared, one a day, at')
        Set-Box (New-Box $slide $x0 406 $cw 84) ($above -join "`r") $FaceBody 9.5 $STONE $ppAlignLeft 14 $false $false 0 | Out-Null
        $link = New-Box $slide $x0 490 $cw 14
        Set-Box $link 'dina280843.github.io/Shofar-Bee' $FaceBody 9.5 $GOLD $ppAlignLeft 14 $true $false 0 | Out-Null
        Add-Link $link $SITE | Out-Null
        Set-Box (New-Box $slide $x0 506 $cw 42) (@('where they can still be played.', '', 'First edition, 2026') -join "`r") $FaceBody 9.5 $STONE $ppAlignLeft 14 $false $false 0 | Out-Null
      }
      '^howto$' {
        Set-Box (New-Box $slide $x0 $MT $cw 20) 'HOW TO PLAY' $FaceLabel 11 $STONE $ppAlignLeft 14 $true $false 2.2 | Out-Null
        Add-Rule $slide $x0 ($MT + 24) $cw $GOLD 1.0 | Out-Null
        $body = @(
          ('Every puzzle gives you seven letters arranged in a hive. Make as many words as you can, and every word must use the letter in the middle ' + $MDASH + ' the gold one. The other six are yours to use as often as you like, or not at all.'), '',
          'Words must be at least four letters long, and letters may repeat within a word. Proper nouns are normally out, though a handful of famous names stand as pangrams where the theme asks for one.', '',
          ('At least one word in each puzzle uses all seven letters. It is called a pangram, and it is worth an extra seven points. Some puzzles hide several. Every pangram in this book is tied to the theme of its page ' + $MDASH + ' the picture, the verse, the place. If you are stuck, read the passage again.'), '',
          'Scoring: a four-letter word is one point. Any longer word scores its own length. A pangram adds seven. The ladder at the foot of each writing page tells you how you are doing, and what a perfect score would be.', '',
          ('The puzzles ran daily through the summer of 2026 and are printed here in the order they appeared. All the answers are gathered at the back ' + $MDASH + ' no peeking.'))
        Set-Box (New-Box $slide $x0 ($MT + 42) $cw 440) ($body -join "`r") $FaceBody 11 $DARK $ppAlignJustify 16 $false $false 0 | Out-Null
      }
      '^blank$' { }
      '^puzzle:(\d+)$' { Paint-Puzzle $slide ($data | Where-Object { $_.day -eq [int]$Matches[1] }) $x0 $cw }
      '^write:(\d+)$'  { Paint-Write  $slide ($data | Where-Object { $_.day -eq [int]$Matches[1] }) $x0 $cw }
      '^answers:(\d+)$' {
        $idx = [int]$Matches[1]
        Paint-Answers $slide $ansSlides[$idx] ($idx -eq 0) $x0 $cw $ansLead $ansRows
      }
    }
  }

  $stem = if ($Sample) { 'ShofarBee_sample_interior_PPT' } else { 'ShofarBee_interior_PPT' }
  $pptx = Join-Path $outDirFull ($stem + '.pptx')
  # SaveAs fails silently if a stray POWERPNT still holds the previous file,
  # which once left a stale deck on disk that looked like a successful build
  foreach ($stale in @($pptx, (Join-Path $outDirFull ($stem + '.pdf')))) {
    if (Test-Path $stale) { Remove-Item $stale -Force -ErrorAction Stop }
  }
  Log 'saving .pptx'
  $pres.SaveAs($pptx, $ppSaveAsPPTX)
  if (-not (Test-Path $pptx)) { throw "SaveAs reported success but $pptx is not on disk" }
  if (-not $NoPdf) {
    $pdf = Join-Path $outDirFull ($stem + '.pdf')
    Log 'saving PDF'
    $pres.SaveAs($pdf, $ppSaveAsPDF)
    Log ('PDF -> ' + $pdf)
  }
  $pres.Close()
  Log ('done, {0} slides' -f $plan.Count)
}
finally {
  try { $ppt.Quit() } catch {}
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
}
