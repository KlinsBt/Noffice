$ErrorActionPreference='Stop'
$app=$null;$presentation=$null
$inputPath=Join-Path (Get-Location) 'tests/fixtures/paint-oracle-cases.json'
$path=[IO.Path]::GetFullPath((Join-Path (Get-Location) '.local/pptx-validation/paint-oracle.pptx'))
$cases=Get-Content -LiteralPath $inputPath -Raw | ConvertFrom-Json
try {
 $app=New-Object -ComObject PowerPoint.Application
 $alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity;$app.DisplayAlerts=1;$app.AutomationSecurity=3
 $presentation=$app.Presentations.Open($path,-1,0,0)
 $results=@()
 for($i=0;$i -lt $cases.Count;$i++){
  $fill=$presentation.Slides.Item($i+1).Shapes.Item(1).Fill
  $rgb=[int]$fill.ForeColor.RGB
  $color='#{0:x2}{1:x2}{2:x2}' -f ($rgb-band255),(($rgb-shr8)-band255),(($rgb-shr16)-band255)
  $results += [ordered]@{name=$cases[$i].name;color=$color;opacity=1-[double]$fill.Transparency}
 }
 [ordered]@{application='Microsoft PowerPoint';version=[string]$app.Version;inputSha256=(Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant();presentationSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();cases=$results} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath tests/fixtures/native-powerpoint-paints.json -Encoding UTF8
 Write-Output "Recorded $($results.Count) native PowerPoint fill colors and opacity values."
} finally {
 if($presentation){try{$presentation.Close()}catch{Write-Warning $_}}
 if($app){try{$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($app.Presentations.Count -eq 0){$app.Quit()}}catch{Write-Warning $_};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
