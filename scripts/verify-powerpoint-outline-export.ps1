$ErrorActionPreference='Stop'
$app=$null;$presentation=$null
function HexColor($color){$rgb=[int]$color.RGB;return '#{0:x2}{1:x2}{2:x2}' -f ($rgb-band255),(($rgb-shr8)-band255),(($rgb-shr16)-band255)}
try {
 $app=New-Object -ComObject PowerPoint.Application
 $alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity;$app.DisplayAlerts=1;$app.AutomationSecurity=3
 $results=@()
 foreach($stage in @('source','edited')){
  $path=[IO.Path]::GetFullPath((Join-Path (Get-Location) ".local/pptx-validation/outlines-$stage.pptx"))
  $presentation=$app.Presentations.Open($path,-1,0,0)
  $slide=$presentation.Slides.Item(1);$shape=$slide.Shapes.Item('No fill')
  $expected=if($stage -eq 'source'){'#ff0000'}else{'#224466'}
  $opacity=if($stage -eq 'source'){1}else{0.6}
  $width=if($stage -eq 'source'){2}else{6}
  $dash=if($stage -eq 'source'){1}else{5}
  if($shape.Line.Visible -eq 0 -or (HexColor $shape.Line.ForeColor) -ne $expected){throw 'Native outline color/visibility mismatch.'}
  if([Math]::Abs(1-$shape.Line.Transparency-$opacity) -gt 0.00001 -or [Math]::Abs($shape.Line.Weight-$width) -gt 0.00001){throw 'Native outline width/transparency mismatch.'}
  if($shape.Line.DashStyle -ne $dash){throw "Native dash mismatch: $($shape.Line.DashStyle) expected $dash"}
  if($shape.Fill.Visible -ne 0){throw 'No-fill shape gained a fill.'}
  if((HexColor $slide.Background.Fill.ForeColor) -ne '#bbddee'){throw 'Inherited background changed.'}
  if($presentation.Slides.Count -ne 2 -or $presentation.Slides.Item(2).Shapes.Item(1).TextFrame.TextRange.Text -ne 'Untouched slide'){throw 'An unrelated slide changed.'}
  $pdf=[IO.Path]::ChangeExtension($path,'.pdf');$presentation.SaveAs($pdf,32)
  $results += [ordered]@{stage=$stage;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();color=$expected;opacity=$opacity;widthPoints=$width;dashStyle=$dash;pdf=$pdf}
  $presentation.Close();$presentation=$null
 }
 [ordered]@{application='Microsoft PowerPoint';version=[string]$app.Version;cases=$results} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath .local/pptx-validation/outlines-report.json -Encoding UTF8
 Write-Output 'Native PowerPoint verified source/edited outlines, opacity, width and dash style; PDF rendering passed.'
} finally {
 if($presentation){try{$presentation.Close()}catch{Write-Warning $_}}
 if($app){try{$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($app.Presentations.Count -eq 0){$app.Quit()}}catch{Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
