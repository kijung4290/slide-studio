param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$presentation = $null
$powerPoint = $null
$previousSecurity = $null
$hadPowerPoint = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
try {
  $powerPoint = New-Object -ComObject PowerPoint.Application
  $previousSecurity = $powerPoint.AutomationSecurity
  $powerPoint.AutomationSecurity = 3
  $presentation = $powerPoint.Presentations.Open($SourcePath, -1, 0, 0)
  $width = [double]$presentation.PageSetup.SlideWidth
  $height = [double]$presentation.PageSetup.SlideHeight
  if ($presentation.Slides.Count -gt 200) { throw '200 slides maximum.' }
  $imageWidth = 1600
  $imageHeight = [int][Math]::Round($imageWidth * $height / $width)
  $slides = @()
  foreach ($slide in $presentation.Slides) {
    $number = [int]$slide.SlideIndex
    $imageName = 'slide-' + $number + '.png'
    $slide.Export((Join-Path $OutputPath $imageName), 'PNG', $imageWidth, $imageHeight)
    $title = ''
    try { if ($slide.Shapes.HasTitle) { $title = [string]$slide.Shapes.Title.TextFrame.TextRange.Text } } catch {}
    if ([string]::IsNullOrWhiteSpace($title)) { $title = 'Slide ' + $number }
    $notes = @()
    foreach ($shape in $slide.NotesPage.Shapes) {
      try {
        if ($shape.Type -eq 14 -and $shape.PlaceholderFormat.Type -eq 2 -and $shape.HasTextFrame -and $shape.TextFrame.HasText) {
          $notes += [string]$shape.TextFrame.TextRange.Text
        }
      } catch {}
    }
    $slides += @{ title=$title.Trim(); image=$imageName; notes=(($notes -join "`n`n") -replace "`r", "`n"); width=$imageWidth; height=$imageHeight }
  }
  $result = @{ slides=$slides; width=$imageWidth; height=$imageHeight }
  $json = ConvertTo-Json -InputObject $result -Depth 8 -Compress
  [IO.File]::WriteAllText((Join-Path $OutputPath 'result.json'), $json, (New-Object Text.UTF8Encoding($false)))
} catch {
  [Console]::Error.WriteLine($_.Exception.Message)
  exit 1
} finally {
  if ($null -ne $presentation) { try { $presentation.Close() } catch {} }
  if ($null -ne $powerPoint) {
    if ($null -ne $previousSecurity) { try { $powerPoint.AutomationSecurity = $previousSecurity } catch {} }
    if (-not $hadPowerPoint) { try { if ($powerPoint.Presentations.Count -eq 0) { $powerPoint.Quit() } } catch {} }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
