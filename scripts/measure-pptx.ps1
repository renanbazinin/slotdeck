# Asks the locally installed PowerPoint how it lays out the text of every shape of a .pptx:
# the lines each text breaks into, and the height the text takes. Writes JSON.
# Dev-only helper for checking the text measurer against PowerPoint; needs Windows + PowerPoint.
#
#   powershell -File scripts/measure-pptx.ps1 -Path out/deck.pptx -Out out/deck.lines.json
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string]$Out
)

$ErrorActionPreference = 'Stop'

$full = (Resolve-Path -LiteralPath $Path).Path

function Read-Text($range) {
  $lines = @()
  $lineWidths = @()
  $count = $range.Lines().Count
  for ($i = 1; $i -le $count; $i++) {
    $line = $range.Lines($i, 1)
    $lines += , ($line.Text -replace "[\r\n\v]+$", '')
    # The width includes the space or paragraph mark that ends the line.
    $lineWidths += , [math]::Round($line.BoundWidth, 3)
  }
  $paragraphs = @()
  $total = $range.Paragraphs().Count
  for ($i = 1; $i -le $total; $i++) { $paragraphs += , $range.Paragraphs($i, 1).Lines().Count }
  return [ordered]@{
    lines       = $lines
    lineWidths  = $lineWidths
    paragraphs  = $paragraphs
    boundHeight = [math]::Round($range.BoundHeight, 2)
    boundWidth  = [math]::Round($range.BoundWidth, 2)
    fontSize    = $range.Font.Size
    font        = $range.Font.Name
  }
}

function Read-Shape($shape, $list) {
  if ($shape.Type -eq 6) {
    foreach ($child in $shape.GroupItems) { Read-Shape $child $list }
    return
  }
  if ($shape.HasTable -eq -1) {
    $table = $shape.Table
    $cells = @()
    for ($r = 1; $r -le $table.Rows.Count; $r++) {
      for ($c = 1; $c -le $table.Columns.Count; $c++) {
        $cell = $table.Cell($r, $c).Shape
        $entry = Read-Text $cell.TextFrame.TextRange
        $entry.row = $r
        $entry.column = $c
        $cells += , $entry
      }
    }
    $rows = @()
    for ($r = 1; $r -le $table.Rows.Count; $r++) { $rows += , [math]::Round($table.Rows.Item($r).Height, 2) }
    $list.Add([ordered]@{
        id = $shape.Id; name = $shape.Name; kind = 'table'
        top = [math]::Round($shape.Top, 2); height = [math]::Round($shape.Height, 2); width = [math]::Round($shape.Width, 2)
        rowHeights = $rows; cells = $cells
      })
    return
  }
  if ($shape.HasTextFrame -ne -1) { return }
  if ($shape.TextFrame.HasText -ne -1) { return }
  $entry = Read-Text $shape.TextFrame.TextRange
  $entry.id = $shape.Id
  $entry.name = $shape.Name
  $entry.kind = 'text'
  $entry.top = [math]::Round($shape.Top, 2)
  $entry.height = [math]::Round($shape.Height, 2)
  $entry.width = [math]::Round($shape.Width, 2)
  $entry.autoSize = $shape.TextFrame2.AutoSize
  $entry.wordWrap = $shape.TextFrame.WordWrap
  $list.Add($entry)
}

# PowerPoint is single-instance: if the user already has it open we attach to
# their session, so we must not quit it when we are done.
$wasRunning = @(Get-Process | Where-Object { $_.ProcessName -eq 'POWERPNT' }).Count -gt 0

$app = New-Object -ComObject PowerPoint.Application
try {
  # Open(FileName, ReadOnly, Untitled, WithWindow) - MsoTriState: -1 true, 0 false
  $pres = $app.Presentations.Open($full, -1, 0, 0)
  try {
    $slides = @()
    foreach ($slide in $pres.Slides) {
      $list = New-Object System.Collections.Generic.List[object]
      foreach ($shape in $slide.Shapes) { Read-Shape $shape $list }
      $slides += , [ordered]@{ slide = $slide.SlideIndex; shapes = $list.ToArray() }
    }
    $json = [ordered]@{ slideWidth = $pres.PageSetup.SlideWidth; slideHeight = $pres.PageSetup.SlideHeight; slides = $slides } | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($Out), $json, (New-Object System.Text.UTF8Encoding $false))
    Write-Output ("measured {0} slides to {1}" -f $pres.Slides.Count, $Out)
  }
  finally {
    $pres.Close()
  }
}
finally {
  if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
