$ErrorActionPreference = 'Stop'
$app = $null
$presentation = $null
try {
  $app = New-Object -ComObject PowerPoint.Application
  $alerts = $app.DisplayAlerts
  $security = $app.AutomationSecurity
  $app.DisplayAlerts = 1
  $app.AutomationSecurity = 3
  $path = [IO.Path]::GetFullPath((Join-Path (Get-Location) '.local/pptx-validation/flipped.pptx'))
  $presentation = $app.Presentations.Open($path, -1, 0, 0)
  if ($presentation.Slides.Count -ne 1) { throw 'Image editing changed slide count.' }
  $pictures = @($presentation.Slides.Item(1).Shapes | Where-Object { $_.Type -eq 13 })
  if ($pictures.Count -ne 1) { throw 'Expected exactly one picture in the browser export.' }
  $picture = $pictures[0]
  if ($picture.HorizontalFlip -ne 0 -or $picture.VerticalFlip -ne -1) { throw 'Native PowerPoint reported incorrect image flip flags.' }
  if ([Math]::Abs($picture.Rotation - 30) -gt 0.01) { throw 'Native PowerPoint reported incorrect image rotation.' }
  $pdf = [IO.Path]::ChangeExtension($path, '.pdf')
  $presentation.SaveAs($pdf, 32)
  [ordered]@{
    application = 'Microsoft PowerPoint'; version = $app.Version
    inputSha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
    horizontalFlip = $picture.HorizontalFlip; verticalFlip = $picture.VerticalFlip
    rotation = $picture.Rotation; pdf = $pdf
  } | ConvertTo-Json | Set-Content -LiteralPath .local/pptx-validation/flips-report.json -Encoding UTF8
  Write-Output 'Native PowerPoint verified image flips and rotation; PDF rendered.'
} finally {
  if ($null -ne $presentation) { try { $presentation.Close() } catch { Write-Warning $_ } }
  if ($null -ne $app) {
    try {
      $app.DisplayAlerts = $alerts
      $app.AutomationSecurity = $security
      if ($app.Presentations.Count -eq 0) { $app.Quit() }
    } catch { Write-Warning $_ }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  }
}
