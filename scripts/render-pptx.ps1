# Renders every slide of a .pptx to PNG using the locally installed PowerPoint.
# Dev-only helper for looking at generated decks; needs Windows + PowerPoint.
#
#   powershell -File scripts/render-pptx.ps1 -Path out/deck.pptx -OutDir out/png
param(
  [Parameter(Mandatory = $true)][string]$Path,
  [Parameter(Mandatory = $true)][string]$OutDir,
  [int]$Width = 1280
)

$ErrorActionPreference = 'Stop'

$full = (Resolve-Path -LiteralPath $Path).Path
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$out = (Resolve-Path -LiteralPath $OutDir).Path

# PowerPoint is single-instance: if the user already has it open we attach to
# their session, so we must not quit it when we are done.
$wasRunning = @(Get-Process | Where-Object { $_.ProcessName -eq 'POWERPNT' }).Count -gt 0

$app = New-Object -ComObject PowerPoint.Application
try {
  # Open(FileName, ReadOnly, Untitled, WithWindow) - MsoTriState: -1 true, 0 false
  $pres = $app.Presentations.Open($full, -1, 0, 0)
  try {
    $height = [int][math]::Round($Width * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
    $count = $pres.Slides.Count
    for ($i = 1; $i -le $count; $i++) {
      $file = Join-Path $out ("slide-{0:D2}.png" -f $i)
      $pres.Slides.Item($i).Export($file, 'PNG', $Width, $height)
    }
    Write-Output ("rendered {0} slides to {1}" -f $count, $out)
  }
  finally {
    $pres.Close()
  }
}
finally {
  if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
