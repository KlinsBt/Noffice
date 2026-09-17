$ErrorActionPreference = 'Stop'
$app=$null;$book=$null
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Invoke-LinkCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++) {
  try {return (& $Action)} catch {
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
  $path=[IO.Path]::GetFullPath((Join-Path (Get-Location) ".local/xlsx-validation/links-$stage.xlsx"))
  $book=Invoke-LinkCall {$app.Workbooks.Open($path,0,$true)}
  if(!$book){throw 'Excel returned no workbook.'}
  $sheet=Invoke-LinkCall {$book.Worksheets.Item('Links')}
  $url=if($stage -eq 'source'){'https://example.com/original'}else{'https://example.com/edited?x=1&y=2'}
  $label=if($stage -eq 'source'){'Original label'}else{'Edited label'}
  $tip=if($stage -eq 'source'){'Original tip'}else{'Edited tip'}
  if((Invoke-LinkCall {$sheet.Range('A1').Hyperlinks.Item(1).Address}) -ne $url){throw 'Address mismatch.'}
  if((Invoke-LinkCall {$sheet.Range('A1').Value2}) -ne $label){throw 'Label mismatch.'}
  if((Invoke-LinkCall {$sheet.Range('A1').Hyperlinks.Item(1).ScreenTip}) -ne $tip){throw 'ScreenTip mismatch.'}
  if((Invoke-LinkCall {$sheet.Range('A2').Value2}) -isnot [double]){throw 'Numeric hyperlink cell became text.'}
  if((Invoke-LinkCall {$sheet.Range('A2').Value2}) -ne 42 -or (Invoke-LinkCall {$sheet.Range('A3').Value2}) -ne 84 -or (Invoke-LinkCall {$sheet.Range('A3').Formula}) -ne '=A2*2'){throw 'Formula/number changed.'}
  if((Invoke-LinkCall {$sheet.Range('B1').Hyperlinks.Item(1).SubAddress}) -ne "'O''Brien Data'!`$B`$3"){throw 'Internal location mismatch.'}
  foreach($ref in @('A4','C2')) {if((Invoke-LinkCall {$sheet.Range($ref).Hyperlinks.Item(1).Address}) -ne 'https://example.com/original'){throw "Unedited shared/range link changed at $ref"}}
  if($stage -eq 'edited') {
   if((Invoke-LinkCall {$sheet.Range('A2').Hyperlinks.Item(1).Address}) -ne 'mailto:hello@example.com'){throw 'Email link mismatch.'}
   if((Invoke-LinkCall {$sheet.Range('A3').Hyperlinks.Item(1).SubAddress}) -ne 'A2'){throw 'Formula-cell link mismatch.'}
   if((Invoke-LinkCall {$sheet.Range('C1').Hyperlinks.Count}) -ne 0){throw 'Removed range-cell link remains.'}
  }
  if(!(Invoke-LinkCall {$sheet.Range('A1').Font.Bold}) -or (Invoke-LinkCall {$sheet.Range('A1').Font.Size}) -ne 14){throw 'Existing font changed.'}
  if((Invoke-LinkCall {$sheet.Range('A1').Comment.Text()}) -ne 'Retained note'){throw 'Existing note changed.'}
  if(!(Invoke-LinkCall {$book.Worksheets.Item('Protected').ProtectContents})){throw 'Worksheet protection changed.'}
  $pdf=[IO.Path]::ChangeExtension($path,'.pdf');Invoke-LinkCall {$sheet.ExportAsFixedFormat(0,$pdf)}
  $results += [ordered]@{stage=$stage;sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant();address=$url;label=$label;screenTip=$tip;numericValue=42;formula='=A2*2';formulaResult=84;pdf=$pdf}
  Invoke-LinkCall {$book.Close($false)};$book=$null
 }
 [ordered]@{application='Microsoft Excel';version=[string]$app.Version;build=[string]$app.Build;cases=$results}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath .local/xlsx-validation/links-report.json -Encoding UTF8
 Write-Output 'Native Excel verified actual source/edited hyperlinks, numeric/formula types, retained links, notes and protection; PDF rendering passed.'
} finally {
 if($book){try{Invoke-LinkCall {$book.Close($false)}}catch{Write-Warning $_}}
 if($app){try{Invoke-LinkCall {$app.Quit()}}catch{Write-Warning $_}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
}
