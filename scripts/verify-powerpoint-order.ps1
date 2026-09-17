$ErrorActionPreference='Stop'
$app=$null; $presentation=$null
try {
  $app=New-Object -ComObject PowerPoint.Application
  $alerts=$app.DisplayAlerts; $security=$app.AutomationSecurity
  $app.DisplayAlerts=1; $app.AutomationSecurity=3
  $path=[IO.Path]::GetFullPath((Join-Path (Get-Location) '.local/pptx-validation/reordered.pptx'))
  $presentation=$app.Presentations.Open($path,-1,0,0)
  if($presentation.Slides.Count -ne 2){throw 'Reordering changed slide count.'}
  $chart=$false
  foreach($shape in $presentation.Slides.Item(1).Shapes){if($shape.HasChart -eq -1){$chart=$true}}
  if(!$chart){throw 'The chart is not on the first slide after reordering.'}
  $text=@()
  foreach($shape in $presentation.Slides.Item(2).Shapes){if($shape.HasTextFrame -eq -1){$text+=$shape.TextFrame.TextRange.Text}}
  if(($text -join "`n") -notlike '*Original title*'){throw 'The title is not on the second slide.'}
  $pdf=[IO.Path]::ChangeExtension($path,'.pdf');$presentation.SaveAs($pdf,32)
  [ordered]@{application='Microsoft PowerPoint';version=$app.Version;inputSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();slides=$presentation.Slides.Count;firstSlideHasChart=$chart;secondSlideText=$text;pdf=$pdf}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath .local/pptx-validation/order-report.json -Encoding UTF8
  Write-Output 'Native PowerPoint verified reordered slide identities and chart relationships; PDF rendered.'
} finally {
  if($null -ne $presentation){try{$presentation.Close()}catch{Write-Warning $_}}
  if($null -ne $app){try{$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($app.Presentations.Count -eq 0){$app.Quit()}}catch{Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
