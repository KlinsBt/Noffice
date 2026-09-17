$ErrorActionPreference='Stop'
$app=$null;$presentation=$null
function HexColor($color){$rgb=[int]$color.RGB;return '#{0:x2}{1:x2}{2:x2}' -f ($rgb-band255),(($rgb-shr8)-band255),(($rgb-shr16)-band255)}
try {
 $app=New-Object -ComObject PowerPoint.Application
 $alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity;$app.DisplayAlerts=1;$app.AutomationSecurity=3
 $results=@()
 foreach($stage in @('source','edited')){
  $path=[IO.Path]::GetFullPath((Join-Path (Get-Location) ".local/pptx-validation/paints-$stage.pptx"))
  $presentation=$app.Presentations.Open($path,-1,0,0)
  $slide=$presentation.Slides.Item(1)
  $expected=if($stage -eq 'source'){'#ddeeaa'}else{'#223344'}
  $opacity=if($stage -eq 'source'){0.5}else{0.75}
  if((HexColor $slide.Background.Fill.ForeColor) -ne '#bbddee'){throw 'Native inherited background mismatch.'}
  $text=$slide.Shapes.Item('Filled text')
  if((HexColor $text.Fill.ForeColor) -ne $expected -or [Math]::Abs(1-$text.Fill.Transparency-$opacity) -gt 0.00001){throw 'Native text-box paint mismatch.'}
  if($slide.Shapes.Item('No fill').Fill.Visible -ne 0){throw 'No-fill shape gained a fill.'}
  if((HexColor $slide.Shapes.Item('Theme fill').Fill.ForeColor) -ne '#345678'){throw 'Theme color changed.'}
  if($presentation.Slides.Count -ne 2 -or $presentation.Slides.Item(2).Shapes.Item(1).TextFrame.TextRange.Text -ne 'Untouched slide'){throw 'An unrelated slide changed.'}
  $pdf=[IO.Path]::ChangeExtension($path,'.pdf');$presentation.SaveAs($pdf,32)
  $results += [ordered]@{stage=$stage;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();background='#bbddee';fill=$expected;opacity=$opacity;theme='#345678';pdf=$pdf}
  $presentation.Close();$presentation=$null
 }
 [ordered]@{application='Microsoft PowerPoint';version=[string]$app.Version;cases=$results} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath .local/pptx-validation/paints-report.json -Encoding UTF8
 Write-Output 'Native PowerPoint verified source and edited fills, opacity, inherited background, no-fill and theme colors; PDFs rendered.'
} finally {
 if($presentation){try{$presentation.Close()}catch{Write-Warning $_}}
 if($app){try{$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($app.Presentations.Count -eq 0){$app.Quit()}}catch{Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
