<#
  Build-Book.ps1  -  generate the paperback interior and the wraparound cover
  as print-ready PDFs, straight from book_data.json (run Export-PuzzleData.ps1
  first). Word COM does the PDF export.

  Every element is an absolutely positioned frame rather than flowing text:
  a puzzle page is a poster, not prose, so exact repeatable placement matters
  more than reflow. Only the answer section flows.

  ASCII-only source. PowerShell 5.1 reads a BOM-less .ps1 as ANSI, so any
  typographic character is built from [char] codes and all prose comes from
  the UTF-8 JSON.

  Usage:
    powershell -File "tools\Build-Book.ps1" -Sample
    powershell -File "tools\Build-Book.ps1" -CoverOnly -Pages 136
#>
param(
  [switch]$Sample,
  [switch]$CoverOnly,
  [switch]$NoPdf,                  # .docx only; see the note on Save-Pdf
  [string]$Only,                   # comma-separated plan entries (diagnostics).
                                   # A string, not string[]: an array parameter
                                   # spills extra tokens onto the next one.
  [int]$Pages = 136,               # interior page count; sets the spine width
  [string]$DataFile = 'tools\book_data.json',
  [string]$OutDir   = 'book'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$outDirFull = Join-Path $root $OutDir
if (-not (Test-Path $outDirFull)) { New-Item -ItemType Directory -Path $outDirFull | Out-Null }

# progress goes to a file as it happens, not just to stdout at exit, so a run
# that stalls can still be diagnosed from outside
$LogPath = Join-Path $outDirFull 'build.log'
Set-Content -Path $LogPath -Value '' -Encoding UTF8
function Log($m) {
  $line = ('[{0:HH:mm:ss}] {1}' -f (Get-Date), $m)
  Add-Content -Path $LogPath -Value $line -Encoding UTF8
  Write-Output $line
}

# ---- typography ------------------------------------------------------------
# Neither Playfair Display nor Inter is installed. Palatino Linotype is the
# closest full-family book face present; Corbel carries the small labels.
$FaceDisplay = 'Palatino Linotype'
$FaceBody    = 'Palatino Linotype'
$FaceLabel   = 'Corbel'
$FaceMono    = 'Cambria'          # answer lists: narrow, legible at 8.5pt

# ---- palette (from the site's :root) ---------------------------------------
function ToBgr($hex) {
  $h = $hex.TrimStart('#')
  $r = [Convert]::ToInt32($h.Substring(0,2),16)
  $g = [Convert]::ToInt32($h.Substring(2,2),16)
  $b = [Convert]::ToInt32($h.Substring(4,2),16)
  $r + ($g * 256) + ($b * 65536)
}
$GOLD   = ToBgr 'C9A84C'
$DARK   = ToBgr '1E1A10'
$STONE  = ToBgr '8B7355'
$PALE   = ToBgr 'FDF6E3'
$BORDER = ToBgr 'E2D9C5'
$INK    = ToBgr '1a1508'
$WHITE  = ToBgr 'FFFFFF'

# ---- typographic characters (kept out of the source as codes) --------------
$MDASH  = [char]0x2014
$MIDDOT = [char]0x00B7
$LSQUO  = [char]0x2018; $RSQUO = [char]0x2019
$LDQUO  = [char]0x201C; $RDQUO = [char]0x201D

# ---- page geometry, 6 x 9 inches -------------------------------------------
$PW = 432.0; $PH = 648.0
$MT = 54.0;  $MB = 54.0
$MIN_ = 63.0        # inside (gutter) margin
$MOUT = 45.0        # outside margin

# content x-range depends on which side of the spread the page falls
function ContentX([bool]$isRecto) {
  if ($isRecto) { @($MIN_, ($PW - $MOUT)) } else { @($MOUT, ($PW - $MIN_)) }
}

# ---- Word constants --------------------------------------------------------
$msoTextHoriz = 1; $msoFalse = 0; $msoTrue = -1
$wdRelPage = 1; $wdWrapNone = 3
$wdAlignLeft = 0; $wdAlignCenter = 1; $wdAlignRight = 2; $wdAlignJustify = 3
$wdLineSpaceExactly = 4
$wdBorderBottom = -3; $wdLineStyleDot = 2; $wdLineWidth050 = 4
$wdFormatDocx = 16; $wdFormatPDF = 17
$wdPageBreak = 7; $wdSectionBreakNextPage = 2

# Word's PDF generation hangs on this machine - BOTH ExportAsFixedFormat and
# SaveAs2(wdFormatPDF) never return, while SaveAs2 to .docx returns instantly.
# So the .docx is always written first and is the reliable artefact; the PDF
# attempt is optional and skippable with -NoPdf.
function Save-Pdf($doc, $docxPath, $pdfPath) {
  Log '  saving .docx master'
  $doc.SaveAs2($docxPath, $wdFormatDocx)
  if ($NoPdf) { Log '  (-NoPdf) skipping PDF'; return }
  Log '  saving PDF'
  $doc.SaveAs2($pdfPath, $wdFormatPDF)
}

# ---- small helpers ---------------------------------------------------------
# Shape.Left/Top are Single. PowerShell hands Word a Double and the COM call
# dies with "Specified cast is not valid", so every placement goes through here.
#
# LockAnchor matters more than it looks. Page-positioned shapes whose anchor
# is free to migrate make Word's PDF renderer repaginate, move the anchor,
# repaginate again - it never terminates, and PDF export hangs forever on any
# document of two or more pages. Pinning the anchor breaks the cycle.
function Place($s, $x, $y) {
  $s.LockAnchor = $true
  $s.Left = [single]$x
  $s.Top  = [single]$y
}

function New-Frame($doc, $anchor, $x, $y, $w, $h) {
  $s = $doc.Shapes.AddTextbox($msoTextHoriz, $x, $y, $w, $h, $anchor)
  $s.Line.Visible = $msoFalse
  $s.Fill.Visible = $msoFalse
  $s.RelativeHorizontalPosition = $wdRelPage
  $s.RelativeVerticalPosition   = $wdRelPage
  $s.WrapFormat.Type = $wdWrapNone
  Place $s $x $y
  $tf = $s.TextFrame
  $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
  $tf.WordWrap = $msoTrue
  $tf.AutoSize = $msoFalse
  $s
}

function Set-Text($shape, $text, $face, $size, $color, $align, [double]$lead, [bool]$bold, [bool]$italic, [double]$spacing) {
  $r = $shape.TextFrame.TextRange
  $r.Text = $text
  $r.Font.Name = $face
  # every one of these is Single in the Word object model; an uncast decimal
  # literal arrives as Double and throws "Specified cast is not valid"
  $r.Font.Size = [single]$size
  $r.Font.Color = $color
  if ($bold)   { $r.Font.Bold = $msoTrue }   else { $r.Font.Bold = $msoFalse }
  if ($italic) { $r.Font.Italic = $msoTrue } else { $r.Font.Italic = $msoFalse }
  if ($spacing -ne 0) { $r.Font.Spacing = [single]$spacing }
  $p = $r.ParagraphFormat
  $p.Alignment = $align
  $p.SpaceBefore = [single]0; $p.SpaceAfter = [single]0
  if ($lead -gt 0) { $p.LineSpacingRule = $wdLineSpaceExactly; $p.LineSpacing = [single]$lead }
}

function Add-Rule($doc, $anchor, $x, $y, $w, $color, [double]$weight) {
  $s = $doc.Shapes.AddLine($x, $y, ($x + $w), $y, $anchor)
  $s.RelativeHorizontalPosition = $wdRelPage
  $s.RelativeVerticalPosition   = $wdRelPage
  $s.WrapFormat.Type = $wdWrapNone
  Place $s $x $y
  $s.Line.ForeColor.RGB = $color
  $s.Line.Weight = [single]$weight
  $s
}

# flat-top hexagon: vertices at 0,60,...,300 degrees, matching the site's
# hexPoints() exactly, so print and screen agree
function Add-Hex($doc, $anchor, [double]$cx, [double]$cy, [double]$R, $fill, $line, [double]$lw) {
  # single[,], not double[,]: AddPolyline marshals as a SAFEARRAY of Single
  # and rejects doubles with "incorrect data type"
  $pts = New-Object 'single[,]' 7,2
  for ($i = 0; $i -lt 6; $i++) {
    $a = [Math]::PI / 180.0 * (60.0 * $i)
    $pts[$i,0] = [single]($cx + $R * [Math]::Cos($a))
    $pts[$i,1] = [single]($cy + $R * [Math]::Sin($a))
  }
  $pts[6,0] = $pts[0,0]; $pts[6,1] = $pts[0,1]
  $s = $doc.Shapes.AddPolyline($pts, $anchor)
  $s.RelativeHorizontalPosition = $wdRelPage
  $s.RelativeVerticalPosition   = $wdRelPage
  $s.WrapFormat.Type = $wdWrapNone
  Place $s ($cx - $R) ($cy - ($R * [Math]::Sqrt(3) / 2.0))
  if ($null -eq $fill) { $s.Fill.Visible = $msoFalse } else { $s.Fill.Visible = $msoTrue; $s.Fill.ForeColor.RGB = $fill }
  $s.Line.ForeColor.RGB = $line
  $s.Line.Weight = [single]$lw
  $s
}

# the whole 6-around-1 hive, centred on (cx,cy)
function Add-Hive($doc, $anchor, [double]$cx, [double]$cy, [double]$R, $centre, $outer, [double]$fontSize) {
  $d = $R * [Math]::Sqrt(3)
  $angles = @(270, 330, 30, 90, 150, 210)
  $cells = @(@{x=$cx; y=$cy; ch=$centre; isC=$true})
  for ($i = 0; $i -lt 6; $i++) {
    $rad = $angles[$i] * [Math]::PI / 180.0
    $cells += @{ x = $cx + $d * [Math]::Cos($rad); y = $cy + $d * [Math]::Sin($rad); ch = $outer[$i]; isC = $false }
  }
  foreach ($c in $cells) {
    if ($c.isC) { Add-Hex $doc $anchor $c.x $c.y $R $GOLD $GOLD 1.0 | Out-Null }
    else        { Add-Hex $doc $anchor $c.x $c.y $R $PALE $GOLD 1.0 | Out-Null }
  }
  # letters go on top, as separate frames, so they are never clipped
  foreach ($c in $cells) {
    $bw = $R * 1.4
    $t = New-Frame $doc $anchor ($c.x - $bw/2) ($c.y - $fontSize*0.72) $bw ($fontSize*1.5)
    $col = if ($c.isC) { $WHITE } else { $DARK }
    Set-Text $t $c.ch $FaceLabel $fontSize $col $wdAlignCenter 0 $true $false 1.0
  }
}

function Ranks($queen) {
  $AZ = [System.MidpointRounding]::AwayFromZero
  $f = @(0.10, 0.25, 0.42, 0.58, 0.75)
  $n = @('Good','Solid','Great','Amazing','Genius')
  $out = @("Beginner 0")
  for ($i = 0; $i -lt 5; $i++) { $out += ("{0} {1}" -f $n[$i], [int][Math]::Round($queen * $f[$i], 0, $AZ)) }
  $out += ("Queen " + $queen)
  $out
}

# ============================================================================
#  INTERIOR
# ============================================================================
function Build-Interior($data, $pdfPath) {
  Log 'interior: starting Word'
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  try {
    $doc = $word.Documents.Add()
    $ps = $doc.PageSetup
    $ps.PageWidth = $PW; $ps.PageHeight = $PH
    $ps.TopMargin = $MT; $ps.BottomMargin = $MB
    $ps.MirrorMargins = $msoTrue
    $ps.LeftMargin = $MIN_; $ps.RightMargin = $MOUT
    $ps.Gutter = 0

    # page plan: title, copyright, how-to, D1 puzzle, D1 write, D50 puzzle,
    # D50 write, answers. Puzzles land on even (left) pages by construction.
    $plan = @('title','copyright','howto','puzzle:1','write:1','puzzle:50','write:50','answers')
    if ($Only) { $keep = $Only -split ','; $plan = @($plan | Where-Object { $keep -contains $_ }) }

    # one tiny empty paragraph per page, to anchor that page's frames to
    $doc.Content.Text = ''
    for ($i = 1; $i -lt $plan.Count; $i++) {
      $doc.Content.InsertParagraphAfter()
      $doc.Paragraphs.Item($doc.Paragraphs.Count).Range.InsertBreak($wdSectionBreakNextPage)
    }
    foreach ($p in $doc.Paragraphs) { $p.Range.Font.Size = 1; $p.SpaceAfter = 0; $p.SpaceBefore = 0 }

    $anchors = @()
    for ($i = 1; $i -le $plan.Count; $i++) { $anchors += $doc.Paragraphs.Item($i).Range }

    for ($i = 0; $i -lt $plan.Count; $i++) {
      $spec = $plan[$i]
      $pageNo = $i + 1
      $isRecto = ($pageNo % 2 -eq 1)
      $xr = ContentX $isRecto
      $x0 = $xr[0]; $x1 = $xr[1]; $cw = $x1 - $x0
      $a = $anchors[$i]
      Log ("  page {0}/{1}  {2}" -f $pageNo, $plan.Count, $spec)

      switch -Regex ($spec) {

        '^title$' {
          $t = New-Frame $doc $a $x0 190 $cw 40
          Set-Text $t 'THE SHOFAR BEE' $FaceDisplay 30 $DARK $wdAlignCenter 34 $true $false 3.5
          Add-Rule $doc $a ($x0 + $cw*0.30) 244 ($cw*0.40) $GOLD 1.0 | Out-Null
          $t = New-Frame $doc $a $x0 262 $cw 24
          Set-Text $t 'Fifty Word Puzzles from Jerusalem' $FaceBody 14 $STONE $wdAlignCenter 18 $false $true 0
          Add-Hive $doc $a ($x0 + $cw/2) 380 26 'B' @('S','H','O','F','A','R') 15
          $t = New-Frame $doc $a $x0 500 $cw 20
          Set-Text $t ('Dina Raveh ' + [char]0x0026 + ' Shalom Brenner') $FaceBody 12.5 $DARK $wdAlignCenter 16 $false $false 1.2
        }

        '^copyright$' {
          $lines = @(
            'The Shofar Bee: Fifty Word Puzzles from Jerusalem',
            '',
            ('Copyright ' + [char]0x00A9 + ' 2026 Dina Raveh and Shalom Brenner'),
            'All rights reserved.',
            '',
            'The puzzles in this book first appeared, one a day, at',
            'dina280843.github.io/Shofar-Bee',
            'where they can still be played.',
            '',
            'First edition, 2026'
          )
          $t = New-Frame $doc $a $x0 420 $cw 150
          Set-Text $t ($lines -join "`r") $FaceBody 9.5 $STONE $wdAlignLeft 14 $false $false 0
        }

        '^howto$' {
          $t = New-Frame $doc $a $x0 $MT $cw 30
          Set-Text $t 'HOW TO PLAY' $FaceLabel 11 $STONE $wdAlignLeft 14 $true $false 2.2
          Add-Rule $doc $a $x0 ($MT + 22) $cw $GOLD 1.0 | Out-Null
          $body = @(
            ('Every puzzle gives you seven letters arranged in a hive. Make as many words as you can, and every word must use the letter in the middle ' + $MDASH + ' the gold one. The other six are yours to use as often as you like, or not at all.'),
            '',
            'Words must be at least four letters long. Proper nouns are out. Letters may repeat within a word.',
            '',
            ('One word in each puzzle uses all seven letters. It is called a pangram, and it is worth an extra seven points. Every pangram in this book is tied to the theme of its page ' + $MDASH + ' the picture, the verse, the place. If you are stuck, read the passage again.'),
            '',
            ('Scoring: a four-letter word is one point. Any longer word scores its own length. A pangram adds seven. The ladder at the foot of each writing page tells you how you are doing, and what a perfect score would be.'),
            '',
            ('The puzzles ran daily through the summer of 2026 and are printed here in the order they appeared. All the answers are gathered at the back ' + $MDASH + ' no peeking at the facing page.')
          )
          $t = New-Frame $doc $a $x0 ($MT + 40) $cw 420
          Set-Text $t ($body -join "`r") $FaceBody 11 $DARK $wdAlignJustify 16 $false $false 0
        }

        '^puzzle:(\d+)$' {
          $p = $data | Where-Object { $_.day -eq [int]$Matches[1] }
          $y = $MT

          $t = New-Frame $doc $a $x0 $y $cw 16
          Set-Text $t ('DAY ' + $p.day) $FaceLabel 10 $STONE $wdAlignLeft 13 $true $false 2.4
          Add-Rule $doc $a $x0 ($y + 17) $cw $BORDER 0.75 | Out-Null
          $y += 32

          $t = New-Frame $doc $a $x0 $y $cw 56
          Set-Text $t $p.hero $FaceDisplay 20 $DARK $wdAlignLeft 25 $true $false 0
          $y += 58

          $t = New-Frame $doc $a $x0 $y $cw 56
          Set-Text $t $p.credit $FaceBody 10.5 $STONE $wdAlignLeft 15 $false $true 0
          $y += 66

          Add-Hive $doc $a ($x0 + $cw/2) ($y + 104) 40 $p.center $p.outer 21
          $y += 224

          $t = New-Frame $doc $a $x0 $y $cw 16
          Set-Text $t ('Every word must contain the centre letter ' + $p.center) $FaceLabel 9.5 $STONE $wdAlignCenter 13 $false $false 0.6
          $y += 30

          # inspiration block, gold rule on the outer edge like the web card
          $barX = $x0
          $s = $doc.Shapes.AddLine($barX, $y, $barX, ($y + 120), $a)
          $s.RelativeHorizontalPosition = $wdRelPage; $s.RelativeVerticalPosition = $wdRelPage
          $s.WrapFormat.Type = $wdWrapNone; Place $s $barX $y
          $s.Line.ForeColor.RGB = $GOLD; $s.Line.Weight = [single]3.0

          $t = New-Frame $doc $a ($x0 + 14) $y ($cw - 14) 14
          Set-Text $t 'INSPIRATION' $FaceLabel 8 $STONE $wdAlignLeft 11 $true $false 2.0
          $t = New-Frame $doc $a ($x0 + 14) ($y + 16) ($cw - 14) 86
          Set-Text $t $p.card $FaceBody 10 $DARK $wdAlignLeft 14 $false $true 0
          $t = New-Frame $doc $a ($x0 + 14) ($y + 104) ($cw - 14) 14
          Set-Text $t $p.cite $FaceLabel 8.5 $STONE $wdAlignLeft 11 $false $false 0
        }

        '^write:(\d+)$' {
          $p = $data | Where-Object { $_.day -eq [int]$Matches[1] }

          $t = New-Frame $doc $a $x0 $MT $cw 16
          Set-Text $t ('DAY ' + $p.day + '  ' + $MIDDOT + '  WORDS FOUND') $FaceLabel 10 $STONE $wdAlignRight 13 $true $false 2.4
          Add-Rule $doc $a $x0 ($MT + 17) $cw $BORDER 0.75 | Out-Null

          # three columns of dotted rules, drawn as bottom-bordered paragraphs
          $gap = 12.0
          $colW = ($cw - 2*$gap) / 3.0
          $gridTop = $MT + 34
          $rows = 35
          $lead = 12.5
          for ($c = 0; $c -lt 3; $c++) {
            $cx = $x0 + $c * ($colW + $gap)
            $f = New-Frame $doc $a $cx $gridTop $colW ($rows * $lead + 4)
            $r = $f.TextFrame.TextRange
            $r.Text = (([string[]](1..$rows | ForEach-Object { '' })) -join "`r")
            $r.Font.Size = [single]9
            $r.Font.Name = $FaceBody
            $r.Font.Color = $DARK
            $pf = $r.ParagraphFormat
            $pf.LineSpacingRule = $wdLineSpaceExactly
            $pf.LineSpacing = [single]$lead
            $pf.SpaceAfter = [single]0; $pf.SpaceBefore = [single]0
            $b = $r.Borders.Item($wdBorderBottom)
            $b.LineStyle = $wdLineStyleDot
            $b.LineWidth = $wdLineWidth050
            $b.Color = $BORDER
          }

          # score ladder
          $ladderY = $gridTop + $rows * $lead + 22
          Add-Rule $doc $a $x0 $ladderY $cw $BORDER 0.75 | Out-Null
          $t = New-Frame $doc $a $x0 ($ladderY + 8) $cw 12
          Set-Text $t 'SCORE' $FaceLabel 8 $STONE $wdAlignLeft 11 $true $false 2.0
          $t = New-Frame $doc $a $x0 ($ladderY + 22) $cw 30
          Set-Text $t ((Ranks $p.queen) -join ('   ' + $MIDDOT + '   ')) $FaceLabel 8.5 $DARK $wdAlignLeft 13 $false $false 0
          $t = New-Frame $doc $a $x0 ($ladderY + 52) $cw 12
          Set-Text $t ('A four-letter word scores 1; longer words score their length; the pangram adds 7.') $FaceBody 8 $STONE $wdAlignLeft 11 $false $true 0
        }

        '^answers$' {
          $t = New-Frame $doc $a $x0 $MT $cw 22
          Set-Text $t 'ANSWERS' $FaceDisplay 16 $DARK $wdAlignLeft 20 $true $false 1.5
          Add-Rule $doc $a $x0 ($MT + 24) $cw $GOLD 1.0 | Out-Null

          $gap = 14.0
          $colW = ($cw - 2*$gap) / 3.0
          $top = $MT + 40
          $blocks = @()
          foreach ($dnum in 1,2) {
            $p = $data | Where-Object { $_.day -eq $dnum }
            $blocks += ('DAY ' + $p.day + '  ' + $MIDDOT + '  ' + $p.hero)
            $blocks += ('centre ' + $p.center + '   pangram ' + $p.pangram + '   Queen ' + $p.queen)
            $blocks += $p.words
            $blocks += ''
          }
          $per = [Math]::Ceiling($blocks.Count / 3.0)
          for ($c = 0; $c -lt 3; $c++) {
            $slice = $blocks[($c*$per)..([Math]::Min(($c+1)*$per-1, $blocks.Count-1))]
            if (-not $slice) { continue }
            $f = New-Frame $doc $a ($x0 + $c*($colW+$gap)) $top $colW 470
            Set-Text $f ($slice -join "`r") $FaceMono 8.5 $DARK $wdAlignLeft 10.5 $false $false 0
            # headers in gold
            for ($ln = 1; $ln -le $f.TextFrame.TextRange.Paragraphs.Count; $ln++) {
              $pr = $f.TextFrame.TextRange.Paragraphs.Item($ln).Range
              if ($pr.Text -match '^DAY \d+') { $pr.Font.Bold = $msoTrue; $pr.Font.Color = $DARK; $pr.Font.Name = $FaceLabel; $pr.Font.Size = [single]8.5 }
              elseif ($pr.Text -match '^centre ') { $pr.Font.Italic = $msoTrue; $pr.Font.Color = $STONE; $pr.Font.Size = [single]7.5 }
            }
          }
        }
      }
    }

    Log 'interior: exporting'
    Save-Pdf $doc ([IO.Path]::ChangeExtension($pdfPath, 'docx')) $pdfPath
    $doc.Close(0)
    Log ("interior -> {0} ({1} pages)" -f $pdfPath, $plan.Count)
  }
  finally {
    # close any still-open document first, or a failed run leaves Quit blocked
    # and strands an invisible WINWORD process. Bounded: if Close silently
    # fails this must not spin forever.
    for ($k = 0; $k -lt 5 -and $word.Documents.Count -gt 0; $k++) {
      try { $word.Documents.Item(1).Close(0) } catch { break }
    }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }
}

# ============================================================================
#  WRAPAROUND COVER
# ============================================================================
function Build-Cover($pdfPath, [int]$pageCount) {
  # KDP: white paper is 0.002252 inch per page; bleed 0.125 in on outer edges
  $bleed   = 9.0
  $spineW  = $pageCount * 0.002252 * 72.0
  $coverW  = 432.0
  $trimH   = 648.0
  $totalW  = $bleed + $coverW + $spineW + $coverW + $bleed
  $totalH  = $bleed + $trimH + $bleed
  $backX   = $bleed
  $spineX  = $bleed + $coverW
  $frontX  = $spineX + $spineW
  $safe    = 18.0    # 0.25 in inside the trim

  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  try {
    $doc = $word.Documents.Add()
    $ps = $doc.PageSetup
    $ps.PageWidth = $totalW; $ps.PageHeight = $totalH
    $ps.TopMargin = 0; $ps.BottomMargin = 0; $ps.LeftMargin = 0; $ps.RightMargin = 0
    $ps.MirrorMargins = $msoFalse
    $doc.Content.Text = ''
    $doc.Paragraphs.Item(1).Range.Font.Size = 1
    $a = $doc.Paragraphs.Item(1).Range

    # full-bleed dark ground
    $bg = $doc.Shapes.AddShape(1, 0, 0, $totalW, $totalH, $a)
    $bg.RelativeHorizontalPosition = $wdRelPage; $bg.RelativeVerticalPosition = $wdRelPage
    $bg.WrapFormat.Type = $wdWrapNone; Place $bg 0 0
    $bg.Fill.ForeColor.RGB = $INK; $bg.Line.Visible = $msoFalse

    # faint honeycomb in the corners, echoing the site banner
    foreach ($corner in @(@{ox=$backX+40; oy=$bleed+40}, @{ox=$frontX+300; oy=$totalH-190})) {
      for ($row = 0; $row -lt 3; $row++) {
        for ($col = 0; $col -lt 3; $col++) {
          $R = 26.0
          $d = $R * [Math]::Sqrt(3)
          $hx = $corner.ox + $col * ($R * 1.5)
          $hy = $corner.oy + $row * $d + (($col % 2) * ($d / 2))
          $h = Add-Hex $doc $a $hx $hy $R $null $GOLD 0.75
          $h.Line.Transparency = [single]0.82
        }
      }
    }

    # ---- front cover ----
    $fx = $frontX + $safe
    $fw = $coverW - 2*$safe
    $t = New-Frame $doc $a $fx 120 $fw 44
    Set-Text $t 'THE' $FaceLabel 14 $GOLD $wdAlignCenter 18 $false $false 6.0
    $t = New-Frame $doc $a $fx 148 $fw 60
    Set-Text $t 'SHOFAR BEE' $FaceDisplay 40 $PALE $wdAlignCenter 46 $true $false 2.0
    Add-Rule $doc $a ($fx + $fw*0.28) 224 ($fw*0.44) $GOLD 1.0 | Out-Null
    $t = New-Frame $doc $a $fx 240 $fw 24
    Set-Text $t 'Fifty Word Puzzles from Jerusalem' $FaceBody 14 $GOLD $wdAlignCenter 18 $false $true 0.4

    Add-Hive $doc $a ($frontX + $coverW/2) 400 34 'B' @('S','H','O','F','A','R') 19

    $t = New-Frame $doc $a $fx 545 $fw 22
    Set-Text $t ('Dina Raveh ' + [char]0x0026 + ' Shalom Brenner') $FaceBody 14 $PALE $wdAlignCenter 18 $false $false 1.4

    # ---- spine ---- (only legal above 100 pages; 136 qualifies)
    $sp = New-Frame $doc $a ($spineX - 150 + $spineW/2) ($totalH/2 - 9) 300 18
    Set-Text $sp ('THE SHOFAR BEE   ' + $MIDDOT + '   RAVEH ' + [char]0x0026 + ' BRENNER') $FaceLabel 9 $PALE $wdAlignCenter 11 $true $false 1.2
    $sp.Rotation = [single]90

    # ---- back cover ----
    $bx = $backX + $safe + 10
    $bw = $coverW - 2*$safe - 20
    $blurb = @(
      ('Every morning for fifty days, a hive of seven letters appeared online with a piece of Jerusalem attached to it ' + $MDASH + ' a scroll signed by a woman in Rome in 1767, the stones of a citadel rebuilt by every conqueror who took the city, a silver amulet carrying the priestly blessing four centuries before the Dead Sea Scrolls.'),
      '',
      ('Here are all fifty, with room to write. Make as many words as you can from the seven letters; every word must use the one in the middle. Somewhere in each puzzle is a word that uses all seven, and it always belongs to the story on the page.'),
      '',
      'No app, no timer, no streak to lose. A pencil will do.'
    )
    $t = New-Frame $doc $a $bx 90 $bw 300
    Set-Text $t ($blurb -join "`r") $FaceBody 11.5 $PALE $wdAlignLeft 17 $false $false 0

    Add-Rule $doc $a $bx 410 $bw $GOLD 0.75 | Out-Null
    $t = New-Frame $doc $a $bx 424 $bw 40
    Set-Text $t ('Play them online, free, at' + "`r" + 'dina280843.github.io/Shofar-Bee') $FaceLabel 11 $GOLD $wdAlignLeft 16 $false $false 0.5
    $t = New-Frame $doc $a $bx 476 ($bw - 170) 30
    Set-Text $t ('A word game from Jerusalem' + "`r" + ('by Dina Raveh ' + [char]0x0026 + ' Shalom Brenner')) $FaceBody 9 $STONE $wdAlignLeft 12 $false $true 0

    # KDP prints its barcode in a 2 x 1.2 in box at the foot of the back
    # cover; leave it empty. Drawn here only as a visible reminder.
    $bcW = 144.0; $bcH = 86.4
    $bcX = $spineX - $safe - $bcW; $bcY = $totalH - $bleed - $safe - $bcH
    $bc = $doc.Shapes.AddShape(1, $bcX, $bcY, $bcW, $bcH, $a)
    $bc.RelativeHorizontalPosition = $wdRelPage; $bc.RelativeVerticalPosition = $wdRelPage
    $bc.WrapFormat.Type = $wdWrapNone; Place $bc $bcX $bcY
    $bc.Fill.ForeColor.RGB = $WHITE; $bc.Fill.Transparency = [single]0.88
    $bc.Line.ForeColor.RGB = $STONE; $bc.Line.Weight = [single]0.5; $bc.Line.DashStyle = 4

    Log 'cover: exporting'
    Save-Pdf $doc ([IO.Path]::ChangeExtension($pdfPath, 'docx')) $pdfPath
    $doc.Close(0)
    Log ("cover -> {0}" -f $pdfPath)
    Log ("        {0:N4} x {1:N4} in canvas, spine {2:N4} in for {3} pages" -f ($totalW/72), ($totalH/72), ($spineW/72), $pageCount)
  }
  finally {
    # close any still-open document first, or a failed run leaves Quit blocked
    # and strands an invisible WINWORD process. Bounded: if Close silently
    # fails this must not spin forever.
    for ($k = 0; $k -lt 5 -and $word.Documents.Count -gt 0; $k++) {
      try { $word.Documents.Item(1).Close(0) } catch { break }
    }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }
}

# ---- run -------------------------------------------------------------------
$dataPath = if ([System.IO.Path]::IsPathRooted($DataFile)) { $DataFile } else { Join-Path $root $DataFile }
$data = [System.IO.File]::ReadAllText($dataPath) | ConvertFrom-Json

if (-not $CoverOnly) { Build-Interior $data (Join-Path $outDirFull 'ShofarBee_sample_interior.pdf') }
Build-Cover (Join-Path $outDirFull 'ShofarBee_cover_wraparound.pdf') $Pages
