$ErrorActionPreference = 'Stop'
$app=$null; $presentation=$null
$path=[IO.Path]::GetFullPath((Join-Path (Get-Location) '.local/pptx-validation/stacked.pptx'))
try {
  $app=New-Object -ComObject PowerPoint.Application
  $alerts=$app.DisplayAlerts; $security=$app.AutomationSecurity
  $app.DisplayAlerts=1; $app.AutomationSecurity=3
  $presentation=$app.Presentations.Open($path,-1,0,0)
  if($presentation.Slides.Count -ne 2){throw 'Layering changed the slide count.'}
  $slide=$presentation.Slides.Item(1)
  if($slide.Shapes.Count -ne 5){throw 'Layering lost an original object or the inserted text.'}
  $expected=@('Inserted layer','image','chart','Top text','Bottom text edited')
  $actual=@()
  for($i=1;$i -le 5;$i++){
    $shape=$slide.Shapes.Item($i)
    $label=if($shape.HasChart -eq -1){'chart'}elseif($shape.Type -eq 13){'image'}else{[string]$shape.TextFrame.TextRange.Text}
    if($label -ne $expected[$i-1] -or $shape.ZOrderPosition -ne $i){throw "Unexpected native layer $i : $label"}
    $actual += [ordered]@{name=[string]$shape.Name;text=$label;position=[int]$shape.ZOrderPosition;left=[double]$shape.Left;top=[double]$shape.Top;width=[double]$shape.Width;height=[double]$shape.Height}
  }
  if($presentation.Slides.Item(2).Shapes.Item(1).TextFrame.TextRange.Text -ne 'Untouched slide'){throw 'The second slide changed.'}
  $pdf=[IO.Path]::ChangeExtension($path,'.pdf');$presentation.SaveAs($pdf,32)
  [ordered]@{application='Microsoft PowerPoint';version=[string]$app.Version;inputSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();layers=$actual;pdf=$pdf} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath .local/pptx-validation/stack-report.json -Encoding UTF8
  Write-Output 'Native PowerPoint verified all five layers, retained image/chart/text and rendered the browser export.'
} finally {
  if($presentation){try{$presentation.Close()}catch{Write-Warning $_}}
  if($app){try{$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($app.Presentations.Count -eq 0){$app.Quit()}}catch{Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
