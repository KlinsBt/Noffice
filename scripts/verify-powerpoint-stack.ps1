$ErrorActionPreference = 'Stop'
$inputPath = Join-Path (Get-Location) 'tests/fixtures/stack-oracle-cases.json'
$outputPath = Join-Path (Get-Location) 'tests/fixtures/native-powerpoint-stack.json'
$cases = Get-Content -LiteralPath $inputPath -Raw | ConvertFrom-Json
$app = $null; $presentation = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  $alerts = $app.DisplayAlerts; $security = $app.AutomationSecurity
  $app.DisplayAlerts = 1; $app.AutomationSecurity = 3
  $presentation = $app.Presentations.Add(0)
  $results = @()
  foreach ($case in $cases) {
    $slide = $presentation.Slides.Add(1, 12)
    foreach ($name in $case.order) {
      $shape = $slide.Shapes.AddShape(1, 100, 100, 200, 100)
      $shape.Name = $name
    }
    $command = @{ front=0; back=1; forward=2; backward=3 }[$case.command]
    $slide.Shapes.Range([object[]]$case.selected).ZOrder($command)
    $order = @(for ($i=1; $i -le $slide.Shapes.Count; $i++) { $slide.Shapes.Item($i).Name })
    $results += [ordered]@{ name=$case.name; order=$order }
    $slide.Delete()
  }
  $report = [ordered]@{application='Microsoft PowerPoint'; version=[string]$app.Version; inputSha256=(Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant(); cases=$results}
  [IO.File]::WriteAllText($outputPath, ($report | ConvertTo-Json -Depth 8) + "`n", [Text.UTF8Encoding]::new($false))
  Write-Output "Recorded $($results.Count) native PowerPoint stacking cases."
} finally {
  if ($presentation) { try { $presentation.Close() } catch { Write-Warning $_ } }
  if ($app) {
    try { $app.DisplayAlerts=$alerts; $app.AutomationSecurity=$security; if ($app.Presentations.Count -eq 0) { $app.Quit() } } catch { Write-Warning $_ }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  }
}
