$ErrorActionPreference = 'Stop'
$app = $null; $presentation = $null
$path = [IO.Path]::GetFullPath((Join-Path (Get-Location) '.local/pptx-validation/arranged.pptx'))
$oraclePath = Join-Path (Get-Location) 'tests/fixtures/native-powerpoint-arrange.json'
$oracle = Get-Content -LiteralPath $oraclePath -Raw -Encoding UTF8 | ConvertFrom-Json
$bottom = ($oracle.cases | Where-Object { $_.name -eq 'rotated-bottom-slide' }).positions
$center = ($oracle.cases | Where-Object { $_.name -eq 'rotated-center-slide' }).positions
try {
  $app = New-Object -ComObject PowerPoint.Application
  $alerts = $app.DisplayAlerts; $security = $app.AutomationSecurity
  $app.DisplayAlerts = 1; $app.AutomationSecurity = 3
  $presentation = $app.Presentations.Open($path, -1, 0, 0)
  if ($presentation.Slides.Count -ne 2) { throw 'Arrangement changed the slide count.' }
  $slide = $presentation.Slides.Item(1)
  if ($slide.Shapes.Count -ne 3) { throw 'Arrangement changed the object count.' }
  $names = @('First object', 'Second object', 'Third object')
  $positions = @()
  for ($i = 0; $i -lt 3; $i++) {
    $shape = $slide.Shapes.Item($i + 1)
    $expected = @{ x=$center[$i].x + 7.5; y=$bottom[$i].y + 1; w=$center[$i].w; h=$center[$i].h; rotation=$center[$i].rotation }
    $actual = [ordered]@{x=[double]$shape.Left; y=[double]$shape.Top; w=[double]$shape.Width; h=[double]$shape.Height; rotation=[double]$shape.Rotation}
    foreach ($key in $expected.Keys) { if ([Math]::Abs($actual[$key] - $expected[$key]) -gt 0.01) { throw "Object $i has incorrect native $key geometry." } }
    if ($shape.TextFrame.TextRange.Text -ne $names[$i]) { throw 'Object text or order changed.' }
    $positions += $actual
  }
  $pdf = [IO.Path]::ChangeExtension($path, '.pdf')
  $presentation.SaveAs($pdf, 32)
  [ordered]@{application='Microsoft PowerPoint'; version=[string]$app.Version; inputSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant(); oracleSha256=(Get-FileHash -LiteralPath $oraclePath -Algorithm SHA256).Hash.ToLowerInvariant(); positions=$positions; pdf=$pdf} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath .local/pptx-validation/arrange-report.json -Encoding UTF8
  Write-Output 'Native PowerPoint verified all arranged positions, sizes, rotations and text; PDF rendered.'
} finally {
  if ($presentation) { try { $presentation.Close() } catch { Write-Warning $_ } }
  if ($app) {
    try { $app.DisplayAlerts=$alerts; $app.AutomationSecurity=$security; if ($app.Presentations.Count -eq 0) { $app.Quit() } } catch { Write-Warning $_ }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  }
}
