$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
$app=$null;$source=$null;$actual=$null
function Invoke-TableCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++) {
  try {return (& $Action)} catch {
   if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
   [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
  }
 }
}
function Read-Tables($book) {
 $ws=Invoke-TableCall {$book.Worksheets.Item('Sales data')};$tables=@();$cells=@()
 foreach($t in (Invoke-TableCall {$ws.ListObjects})) {
  $columns=@();foreach($c in $t.ListColumns){$columns += [string]$c.Name}
  $tables += [ordered]@{name=[string]$t.Name;ref=[string]$t.Range.Address();columns=$columns;style=[string]$t.TableStyle.Name;rows=[bool]$t.ShowTableStyleRowStripes;cols=[bool]$t.ShowTableStyleColumnStripes;first=[bool]$t.ShowTableStyleFirstColumn;last=[bool]$t.ShowTableStyleLastColumn;totals=[bool]$t.ShowTotals;headers=[bool]$t.ShowHeaders}
 }
 for($r=1;$r -le 6;$r++){for($c=1;$c -le 7;$c++){
  $cell=Invoke-TableCall {$ws.Cells.Item($r,$c)};$format=Invoke-TableCall {$cell.DisplayFormat}
  $cells += [ordered]@{ref=[string]$cell.Address();value=$cell.Value2;formula=[string]$cell.Formula;format=[string]$cell.NumberFormat;fill=[int]$format.Interior.Color;color=[int]$format.Font.Color;bold=[bool]$format.Font.Bold}
 }}
 return [ordered]@{tables=$tables;cells=$cells;note=[string]$ws.Range('F1').Comment.Text()}
}
function Rgb-Css([int]$color) {return "rgb($($color -band 255), $(($color -shr 8) -band 255), $(($color -shr 16) -band 255))"}
try {
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0
 [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $root=Join-Path (Get-Location) '.local/xlsx-tables';$sourcePath=Join-Path $root 'source.xlsx'
 $source=Invoke-TableCall {$app.Workbooks.Open($sourcePath,0,$true)}
 $ws=Invoke-TableCall {$source.Worksheets.Item('Sales data')};$table=Invoke-TableCall {$ws.ListObjects.Item('Sales')}
 $results=@();$mismatches=@();$browserMismatches=@();$colors=Get-Content -Raw (Join-Path $root 'browser-colors.json') | ConvertFrom-Json
 foreach($stage in @('resized','style-2','style-3','style-4','style-5','style-6','style-7','bands','none','shrunk','created')) {
  switch -Regex ($stage) {
   '^resized$' {Invoke-TableCall {$table.Resize($ws.Range('A1:C6'))}}
   '^style-(\d)$' {$styleNumber=[int]$Matches[1];Invoke-TableCall {$table.TableStyle="TableStyleMedium$styleNumber"}}
   '^bands$' {Invoke-TableCall {$table.ShowTableStyleRowStripes=$false;$table.ShowTableStyleColumnStripes=$true;$table.ShowTableStyleFirstColumn=$true;$table.ShowTableStyleLastColumn=$true}}
   '^none$' {Invoke-TableCall {$table.TableStyle=''}}
   '^shrunk$' {Invoke-TableCall {$table.Resize($ws.Range('A1:C4'))};Invoke-TableCall {$table.TableStyle='TableStyleMedium7';$table.ShowTableStyleRowStripes=$true;$table.ShowTableStyleColumnStripes=$false;$table.ShowTableStyleFirstColumn=$true;$table.ShowTableStyleLastColumn=$false}}
   '^created$' {
    $supplies=Invoke-TableCall {$ws.ListObjects.Add(1,$ws.Range('F3:G6'),[Type]::Missing,1)}
    Invoke-TableCall {$supplies.Name='Supplies';$supplies.TableStyle='TableStyleMedium4';$supplies.ShowTableStyleRowStripes=$true;$supplies.ShowTableStyleLastColumn=$true;$ws.Range('E4').Formula='=SUM(Supplies[Amount])'}
   }
  }
  Invoke-TableCall {$app.CalculateFull()};$expected=Read-Tables $source
  $path=Join-Path $root "$stage.xlsx";$actual=Invoke-TableCall {$app.Workbooks.Open($path,0,$true)}
  Invoke-TableCall {$app.CalculateFull()};$observed=Read-Tables $actual
  if(($expected.tables | ConvertTo-Json -Depth 8 -Compress) -cne ($observed.tables | ConvertTo-Json -Depth 8 -Compress)) {$mismatches += @{stage=$stage;part='tables';expected=$expected.tables;actual=$observed.tables}}
  for($i=0;$i -lt $expected.cells.Count;$i++) {
   foreach($key in @('value','formula','format','fill','color','bold')) {
    if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]) {$mismatches += @{stage=$stage;ref=$expected.cells[$i].ref;part=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}}
   }
  }
  if($expected.note -cne $observed.note){$mismatches += @{stage=$stage;part='note'}}
  if($stage -match '^style-(\d)$' -or $stage -in @('bands','none')) {
   $styleKey=if($stage -match '^style-(\d)$'){$Matches[1]}else{$stage}
   foreach($browser in @($colors | Where-Object {[string]$_.style -eq $styleKey})) {
    $f=Invoke-TableCall {$actual.Worksheets.Item('Sales data').Range($browser.ref).DisplayFormat}
    $native=@{fill=(Rgb-Css ([int]$f.Interior.Color));color=(Rgb-Css ([int]$f.Font.Color));bold=[bool]$f.Font.Bold}
    foreach($key in @('fill','color','bold')) {if([string]$browser.$key -cne [string]$native[$key]) {$browserMismatches += @{stage=$stage;ref=$browser.ref;part=$key;browser=$browser.$key;native=$native[$key]}}}
   }
  }
  $results += [ordered]@{stage=$stage;sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash;expected=$expected;actual=$observed}
  Invoke-TableCall {$actual.Close($false)};$actual=$null
 }
 $newPath=Join-Path $root 'new.xlsx';$actual=Invoke-TableCall {$app.Workbooks.Open($newPath,0,$true)}
 Invoke-TableCall {$app.CalculateFull()};$newSheet=Invoke-TableCall {$actual.Worksheets.Item('Sheet 1')};$newTable=Invoke-TableCall {$newSheet.ListObjects.Item('Orders')}
 if([string]$newTable.Range.Address() -cne '$A$1:$C$3' -or [string]$newSheet.Range('B1').Value2 -cne 'Amount2' -or [string]$newSheet.Range('C1').Value2 -cne 'Column3' -or [double]$newSheet.Range('D1').Value2 -ne 2){$mismatches += @{stage='new';part='table/header/formula'}}
 foreach($browser in (Get-Content -Raw (Join-Path $root 'new-browser-colors.json') | ConvertFrom-Json)) {
  $f=Invoke-TableCall {$newSheet.Range($browser.ref).DisplayFormat}
  $native=@{fill=(Rgb-Css ([int]$f.Interior.Color));color=(Rgb-Css ([int]$f.Font.Color));bold=[bool]$f.Font.Bold}
  foreach($key in @('fill','color','bold')) {if([string]$browser.$key -cne [string]$native[$key]) {$browserMismatches += @{stage='new';ref=$browser.ref;part=$key;browser=$browser.$key;native=$native[$key]}}}
 }
 Invoke-TableCall {$actual.Close($false)};$actual=$null
 $report=[ordered]@{excelVersion=[string]$app.Version;excelBuild=[string]$app.Build;sourceSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $sourcePath).Hash;newWorkbookSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $newPath).Hash;stages=$results;cellSnapshots=462;browserStyleCells=33;mismatches=$mismatches;browserMismatches=$browserMismatches}
 $report | ConvertTo-Json -Depth 16 | Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output "Native Excel: $($mismatches.Count) package mismatches; $($browserMismatches.Count) browser style mismatches."
 if($mismatches.Count -or $browserMismatches.Count){throw 'Table comparison failed; see native-report.json.'}
} finally {if($actual){Invoke-TableCall {$actual.Close($false)}};if($source){Invoke-TableCall {$source.Close($false)}};if($app){Invoke-TableCall {$app.Quit()}}}
