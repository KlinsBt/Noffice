$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
$app=$null;$source=$null;$actual=$null
function Invoke-RenameCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++) {
  try {return (& $Action)} catch {
   if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
   [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
  }
 }
}
function Read-RenameBook($book) {
 $ws=Invoke-RenameCall {$book.Worksheets.Item('Sales data')};$cells=@();$names=@()
 foreach($name in $book.Names){$names += @{name=[string]$name.Name;formula=[string]$name.RefersTo}}
 for($r=1;$r -le 6;$r++){for($c=1;$c -le 8;$c++){
  $cell=Invoke-RenameCall {$ws.Cells.Item($r,$c)}
  $cells += [ordered]@{ref=[string]$cell.Address();value=$cell.Value2;formula=[string]$cell.Formula;format=[string]$cell.NumberFormat}
 }}
 $table=Invoke-RenameCall {$ws.ListObjects.Item(1)};$columns=@();foreach($c in $table.ListColumns){$columns += [string]$c.Name}
 return [ordered]@{name=[string]$table.Name;columns=$columns;cells=$cells;names=$names;validation=[string]$ws.Range('K2').Validation.Formula1;conditional=[string]$ws.Range('J2').FormatConditions.Item(1).Formula1;note=[string]$ws.Range('F1').Comment.Text()}
}
try {
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0
 [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $root=Join-Path (Get-Location) '.local/xlsx-table-rename';$sourcePath=Join-Path $root 'source.xlsx'
 $source=Invoke-RenameCall {$app.Workbooks.Open($sourcePath,0,$true)}
 $ws=Invoke-RenameCall {$source.Worksheets.Item('Sales data')};$table=Invoke-RenameCall {$ws.ListObjects.Item('Sales')}
 Invoke-RenameCall {$ws.Range('B2').Value2=4}
 $results=@();$mismatches=@()
 foreach($stage in @('renamed','special','escaped')) {
  switch($stage) {
   'renamed' {Invoke-RenameCall {$table.Name='Orders';$table.ListColumns.Item(2).Name='Units sold'}}
   'special' {Invoke-RenameCall {$table.Name='Revenue';$table.ListColumns.Item(2).Name='Net, total'}}
   'escaped' {Invoke-RenameCall {$table.ListColumns.Item(2).Name="#[Q]@'"}}
  }
  Invoke-RenameCall {$app.CalculateFull()};$expected=Read-RenameBook $source
  $path=Join-Path $root "$stage.xlsx";$actual=Invoke-RenameCall {$app.Workbooks.Open($path,0,$true)}
  Invoke-RenameCall {$app.CalculateFull()};$observed=Read-RenameBook $actual
  foreach($key in @('name','columns','names','validation','conditional','note')) {
   if(($expected[$key] | ConvertTo-Json -Depth 8 -Compress) -cne ($observed[$key] | ConvertTo-Json -Depth 8 -Compress)) {$mismatches += @{stage=$stage;part=$key;expected=$expected[$key];actual=$observed[$key]}}
  }
  for($i=0;$i -lt $expected.cells.Count;$i++) {foreach($key in @('value','formula','format')) {
   if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]) {$mismatches += @{stage=$stage;ref=$expected.cells[$i].ref;part=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}}
  }}
  $results += @{stage=$stage;sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash;expected=$expected;actual=$observed}
  Invoke-RenameCall {$actual.Close($false)};$actual=$null
 }
 $report=@{excelVersion=[string]$app.Version;excelBuild=[string]$app.Build;sourceSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $sourcePath).Hash;stages=$results;cellSnapshots=144;mismatches=$mismatches}
 $report | ConvertTo-Json -Depth 16 | Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output "Native Excel table renaming: $($mismatches.Count) mismatches across 144 cell snapshots and table/name/rule metadata."
 if($mismatches.Count){throw 'Table rename comparison failed; see native-report.json.'}
} finally {if($actual){Invoke-RenameCall {$actual.Close($false)}};if($source){Invoke-RenameCall {$source.Close($false)}};if($app){Invoke-RenameCall {$app.Quit()}}}
