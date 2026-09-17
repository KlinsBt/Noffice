$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/excel-dependencies'))
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Hash($path){(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Invoke-DependencyCall([scriptblock]$Action){
 for($attempt=0;$attempt -lt 40;$attempt++){
  try{return (& $Action)}catch{
   $errorObject=$_.Exception;$busy=$false
   while($errorObject){if($errorObject.HResult -in @(-2147418111,-2147417846,-2146777998)){$busy=$true};$errorObject=$errorObject.InnerException}
   if(!$busy -or $attempt -eq 39){throw}
   if($ownedProcess){[ExcelTestWindow]::DismissReminder($ownedProcess)}
   Start-Sleep -Milliseconds 250
  }
 }
}
function Snapshot($book){
 $data=Invoke-DependencyCall {$book.Worksheets.Item('Data')};$summary=Invoke-DependencyCall {$book.Worksheets.Item('Summary')}
 return @{data=@((Invoke-DependencyCall {$data.Range('A1').Value2}),(Invoke-DependencyCall {$data.Range('B1').Value2}));summary=@((Invoke-DependencyCall {$summary.Range('A1').Value2}),(Invoke-DependencyCall {$summary.Range('B1').Value2}))}
}
$app=$null;$book=$null;$links=$null;$security=$null;[uint32]$ownedProcess=0
try{
 $receipt=Get-Content (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
 $source=[string](Join-Path $root 'source.xlsx');$output=[string](Join-Path $root 'browser.xlsx')
 if(!$receipt.passed -or (Hash $source) -ne $receipt.sourceHash -or (Hash $output) -ne $receipt.exportHash){throw 'Stale browser evidence'}
 $app=New-Object -ComObject Excel.Application
 [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr](Invoke-DependencyCall {$app.Hwnd}),[ref]$ownedProcess)
 if((Invoke-DependencyCall {$app.Workbooks.Count}) -ne 0){throw 'Native dependency oracle requires an idle owned instance.'}
 $alerts=Invoke-DependencyCall {$app.DisplayAlerts};$security=Invoke-DependencyCall {$app.AutomationSecurity};$links=Invoke-DependencyCall {$app.AskToUpdateLinks}
 Invoke-DependencyCall {$app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.AskToUpdateLinks=$false}
 $baseline=Get-Content (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
 $exeHash=Hash (Join-Path ([string](Invoke-DependencyCall {$app.Path})) 'EXCEL.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'EXCEL.EXE').sha256){throw 'Excel baseline changed'}
 [ExcelTestWindow]::DismissReminder($ownedProcess)
 $book=Invoke-DependencyCall {$app.Workbooks.Open($source,0,$false)}
 Invoke-DependencyCall {$app.CalculateFullRebuild()};$initial=Snapshot $book
 Invoke-DependencyCall {$book.Worksheets.Item('Data').Range('A1').Value2=[double]5}
 Invoke-DependencyCall {$app.CalculateFullRebuild()};$expected=Snapshot $book
 Invoke-DependencyCall {$book.SaveAs([string](Join-Path $root 'expected.xlsx'),51)};Invoke-DependencyCall {$book.Close($false)};$book=$null
 $book=Invoke-DependencyCall {$app.Workbooks.Open($output,0,$true)};$cached=Snapshot $book
 Invoke-DependencyCall {$app.CalculateFullRebuild()};$actual=Snapshot $book
 foreach($field in @('data','summary')){
  if(($actual.$field -join ',') -ne ($expected.$field -join ',') -or ($cached.$field -join ',') -ne ($expected.$field -join ',')){throw "Native dependency result differs: $field"}
  if(($actual.$field -join ',') -ne ($receipt.expected.$field -join ',')){throw "Browser/native dependency result differs: $field"}
 }
 @{passed=$true;sourceHash=Hash $source;exportHash=Hash $output;expectedHash=Hash (Join-Path $root 'expected.xlsx');initial=$initial;expected=$expected;actual=$actual;executableHash=$exeHash;scriptHash=Hash $PSCommandPath}|ConvertTo-Json -Depth 8|Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output 'Native dependency chain: input 2 -> 5; intermediate 6 -> 15; dependent 7 -> 16; unrelated result 4 unchanged.'
}finally{
 try{if($book){Invoke-DependencyCall {$book.Close($false)}}}finally{if($app){try{if($null -ne $security){Invoke-DependencyCall {$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($null -ne $links){$app.AskToUpdateLinks=$links};$app.Quit()}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}
