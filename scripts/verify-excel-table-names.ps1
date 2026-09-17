$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
$app=$null;$source=$null;$actual=$null
function Invoke-NameCall([scriptblock]$Action){for($retry=0;$retry -lt 20;$retry++){try{return (& $Action)}catch{if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw};[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}}}
function Read-NameBook($book){
 $ws=Invoke-NameCall {$book.Worksheets.Item('Sales data')};$cells=@()
 for($r=1;$r -le 6;$r++){for($c=1;$c -le 8;$c++){$cell=Invoke-NameCall {$ws.Cells.Item($r,$c)};$cells += @{ref=[string]$cell.Address();value=$cell.Value2;formula=[string]$cell.Formula}}}
 return @{name=[string]$ws.ListObjects.Item(1).DisplayName;cells=$cells;definedName=[string]$book.Names.Item('Units').RefersTo}
}
try{
 $app=New-Object -ComObject Excel.Application;$app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $root=Join-Path (Get-Location) '.local/xlsx-table-names';$path=Join-Path $root 'source.xlsx';$names=Get-Content -Raw -Encoding UTF8 (Join-Path $root 'names.json')|ConvertFrom-Json
 $source=Invoke-NameCall {$app.Workbooks.Open($path,0,$true)};$table=Invoke-NameCall {$source.Worksheets.Item('Sales data').ListObjects.Item(1)};$results=@();$mismatches=@()
 for($stage=0;$stage -lt $names.Count;$stage++){
  $name=$names[$stage];Invoke-NameCall {$table.DisplayName=$name;$table.Name=$name;$app.CalculateFull()};$expected=Read-NameBook $source
  $output=Join-Path $root "stage-$stage.xlsx";$actual=Invoke-NameCall {$app.Workbooks.Open($output,0,$true)};Invoke-NameCall {$app.CalculateFull()};$observed=Read-NameBook $actual
  foreach($key in @('name','definedName')){if([string]$expected[$key] -cne [string]$observed[$key]){$mismatches += @{stage=$stage;part=$key;expected=$expected[$key];actual=$observed[$key]}}}
  for($i=0;$i -lt $expected.cells.Count;$i++){foreach($key in @('formula','value')){if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]){$mismatches += @{stage=$stage;ref=$expected.cells[$i].ref;part=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}}}}
  $results += @{stage=$stage;sha256=(Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash;expected=$expected;actual=$observed};Invoke-NameCall {$actual.Close($false)};$actual=$null
 }
 @{excelVersion=[string]$app.Version;excelBuild=[string]$app.Build;sourceSha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash;cellSnapshots=192;stages=$results;mismatches=$mismatches}|ConvertTo-Json -Depth 16|Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output "Native table names: $($mismatches.Count) mismatches across 192 cell snapshots and names."
 if($mismatches.Count){throw 'Table-name comparison failed; see native-report.json.'}
}finally{if($actual){Invoke-NameCall {$actual.Close($false)}};if($source){Invoke-NameCall {$source.Close($false)}};if($app){Invoke-NameCall {$app.Quit()}}}
