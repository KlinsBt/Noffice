$ErrorActionPreference='Stop'
$app=$null;$book=$null
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Invoke-ValidationCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++) {
  try{return (& $Action)}catch{
   if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
   [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
  }
 }
}
try {
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0
 [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $results=@()
 foreach($stage in @('source','edited')) {
  $path=[IO.Path]::GetFullPath((Join-Path (Get-Location) ".local/xlsx-validation/validation-$stage.xlsx"))
  $book=Invoke-ValidationCall {$app.Workbooks.Open($path,0,$true)}
  if(!$book){throw 'Excel returned no workbook.'}
  $sheet=Invoke-ValidationCall {$book.Worksheets.Item('Entry')}
  $expectedType=if($stage -eq 'source'){1}else{2}
  if((Invoke-ValidationCall {$sheet.Range('A2').Validation.Type}) -ne $expectedType){throw 'Validation type mismatch.'}
  if((Invoke-ValidationCall {$sheet.Range('A3').Validation.Type}) -ne 1){throw 'Surviving relative rule was lost.'}
  Invoke-ValidationCall {$sheet.Activate();$sheet.Range('A3').Select()}
  $relative=Invoke-ValidationCall {[string]$sheet.Range('A3').Validation.Formula1}
  if($relative.TrimStart('=') -ne 'B3'){throw "Relative formula mismatch: $relative"}
  if((Invoke-ValidationCall {$sheet.Range('D9000').Validation.Type}) -ne 3){throw 'Large-range list rule was lost.'}
  if((Invoke-ValidationCall {$sheet.Range('A2').Comment.Text()}) -ne 'Preserve this note'){throw 'Note changed.'}
  if(!(Invoke-ValidationCall {$book.Worksheets.Item('Protected').ProtectContents})){throw 'Protection changed.'}
  $oldValue=Invoke-ValidationCall {$sheet.Range('A2').Value2}
  $expectedValue=if($stage -eq 'source'){5}else{0.5}
  if($oldValue -ne $expectedValue -or (Invoke-ValidationCall {$sheet.Range('E1').Value2}) -ne (2*$expectedValue)){throw 'Edited/cached values mismatch.'}
  if($stage -eq 'edited') {
   if((Invoke-ValidationCall {$sheet.Range('A2').Validation.InputMessage}) -ne 'Enter a fraction'){throw 'Input message changed.'}
   if((Invoke-ValidationCall {$sheet.Range('A2').Validation.ErrorMessage}) -ne 'Between zero and one'){throw 'Stop alert message changed.'}
   if(!(Invoke-ValidationCall {$sheet.Range('A2').Validation.ShowError})){throw 'Stop alert was disabled.'}
   $removed=try{Invoke-ValidationCall {$sheet.Range('D4').Validation.Type}}catch{$null}
   if($null -ne $removed){throw 'Cleared validation remains.'}
  }
  $cases=@()
  foreach($value in @(0,0.5,1,2,'text')) {
   Invoke-ValidationCall {if($value -is [string]){$sheet.Range('A2').Value2=[string]$value}else{$sheet.Range('A2').Value2=[double]$value}}
   $valid=Invoke-ValidationCall {[bool]$sheet.Range('A2').Validation.Value}
   $expected=if($stage -eq 'source'){$value -eq 2}else{$value -isnot [string] -and $value -ge 0 -and $value -le 1}
   if($valid -ne $expected){throw "Native validation mismatch for $stage input $value"}
   $cases += [ordered]@{value=$value;valid=$valid}
  }
  Invoke-ValidationCall {$sheet.Range('A2').Value2=$oldValue;$app.Calculate()}
  $oldList=Invoke-ValidationCall {$sheet.Range('D2').Value2}
  $choices=if($stage -eq 'source'){@('Red','Blue')}else{@('Blue','Red')}
  foreach($value in $choices) {
   Invoke-ValidationCall {$sheet.Range('D2').Value2=[string]$value}
   $valid=Invoke-ValidationCall {[bool]$sheet.Range('D2').Validation.Value}
   if($valid -ne ($value -eq $choices[0])){throw 'Native list validation mismatch.'}
   $cases += [ordered]@{value=$value;valid=$valid}
  }
  Invoke-ValidationCall {if($null -eq $oldList){[void]$sheet.Range('D2').ClearContents()}else{$sheet.Range('D2').Value2=[string]$oldList}}
  $pdf=[IO.Path]::ChangeExtension($path,'.pdf');Invoke-ValidationCall {$sheet.ExportAsFixedFormat(0,$pdf)}
  $results += [ordered]@{stage=$stage;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();type=$expectedType;survivingFormula=$relative;cases=$cases;pdf=$pdf}
  Invoke-ValidationCall {$book.Close($false)};$book=$null
 }
 [ordered]@{application='Microsoft Excel';version=[string]$app.Version;build=[string]$app.Build;results=$results}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath .local/xlsx-validation/validation-report.json -Encoding UTF8
 Write-Output 'Native Excel verified validation metadata, surviving relative/large-range rules, 14 input decisions and source/edited PDF rendering.'
} finally {
 if($book){try{Invoke-ValidationCall {$book.Close($false)}}catch{Write-Warning $_}}
 if($app){try{Invoke-ValidationCall {$app.Quit()}}catch{Write-Warning $_}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
}
