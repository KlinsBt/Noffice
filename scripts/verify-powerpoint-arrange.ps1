$ErrorActionPreference = 'Stop'
$inputPath = Join-Path (Get-Location) 'tests/fixtures/arrange-oracle-cases.json'
$outputPath = Join-Path (Get-Location) 'tests/fixtures/native-powerpoint-arrange.json'
$cases = Get-Content -LiteralPath $inputPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app = $null; $presentation = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  $alerts = $app.DisplayAlerts; $security = $app.AutomationSecurity
  $app.DisplayAlerts = 1; $app.AutomationSecurity = 3
  $presentation = $app.Presentations.Add(0)
  $results = @()
  foreach ($case in $cases) {
    $presentation.PageSetup.SlideWidth = $case.width
    $presentation.PageSetup.SlideHeight = $case.height
    $slide = $presentation.Slides.Add(1, 12)
    $names = @(); $i = 0
    foreach ($shape in $case.shapes) {
      $i++
      $object = $slide.Shapes.AddShape(1, $shape.x, $shape.y, $shape.w, $shape.h)
      $object.Name = "Case$i"; $object.Rotation = $shape.rotation
      $names += $object.Name
    }
    $range = $slide.Shapes.Range([object[]]$names)
    $relative = 0; if ($case.toSlide) { $relative = -1 }
    $align = @{left=0; center=1; right=2; top=3; middle=4; bottom=5}
    if ($align.ContainsKey($case.command)) { $range.Align($align[$case.command], $relative) }
    else { $direction = 0; if ($case.command -eq 'vertical') { $direction = 1 }; $range.Distribute($direction, $relative) }
    $positions = @()
    foreach ($name in $names) {
      $object = $slide.Shapes.Item($name)
      $positions += [ordered]@{x=[double]$object.Left; y=[double]$object.Top; w=[double]$object.Width; h=[double]$object.Height; rotation=[double]$object.Rotation}
    }
    $results += [ordered]@{name=$case.name; positions=$positions}
    $slide.Delete()
  }
  $report = [ordered]@{application='Microsoft PowerPoint'; version=[string]$app.Version; inputSha256=(Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant(); cases=$results}
  [IO.File]::WriteAllText($outputPath, ($report | ConvertTo-Json -Depth 8) + "`n", (New-Object Text.UTF8Encoding($false)))
  Write-Output ("Recorded {0} native PowerPoint arrangement cases." -f $results.Count)
} finally {
  if ($presentation) { try { $presentation.Close() } catch { Write-Warning $_ } }
  if ($app) {
    try {
      $app.DisplayAlerts = $alerts; $app.AutomationSecurity = $security
      if ($app.Presentations.Count -eq 0) { $app.Quit() }
    } catch { Write-Warning $_ }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  }
}
