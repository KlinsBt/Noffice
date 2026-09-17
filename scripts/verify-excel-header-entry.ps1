$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
$app=$null;$source=$null;$actual=$null;$scratch=$null
function Invoke-HeaderCall([scriptblock]$Action) {
 for($retry=0;$retry -lt 20;$retry++) {try {return (& $Action)} catch {
  if($_.Exception.HResult -notin @(-2147418111,-2147417846,-2146777998) -or $retry -eq 19){throw}
  [ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250
 }}
}
function Read-HeaderBook($book) {
 $ws=Invoke-HeaderCall {$book.Worksheets.Item('Sales data')};$cells=@();$names=@();$columns=@()
 foreach($n in $book.Names){$names += @{name=[string]$n.Name;formula=[string]$n.RefersTo}}
 foreach($c in $ws.ListObjects.Item(1).ListColumns){$columns += [string]$c.Name}
 for($r=1;$r -le 6;$r++){for($c=1;$c -le 8;$c++){
  $cell=Invoke-HeaderCall {$ws.Cells.Item($r,$c)}
  $cells += @{ref=[string]$cell.Address();value=$cell.Value2;formula=[string]$cell.Formula;format=[string]$cell.NumberFormat}
 }}
 return @{columns=$columns;names=$names;cells=$cells;validation=[string]$ws.Range('K2').Validation.Formula1;conditional=[string]$ws.Range('J2').FormatConditions.Item(1).Formula1;note=[string]$ws.Range('F1').Comment.Text()}
}
function Copy-Headers([string[]]$headings,[bool]$includeData=$false) {
 $script:scratch=Invoke-HeaderCall {$app.Workbooks.Add()};$sw=Invoke-HeaderCall {$scratch.Worksheets.Item(1)}
 for($i=0;$i -lt $headings.Count;$i++){Invoke-HeaderCall {$sw.Cells.Item(1,$i+1).Value2=$headings[$i]}}
 if($includeData){Invoke-HeaderCall {$sw.Range('A2').Value2='North';$sw.Range('B2').Value2=4;$sw.Range('C2').Value2=40}}
 $ref=if($includeData){'A1:C2'}else{'A1:C1'}
 [void](Invoke-HeaderCall {$sw.Range($ref).Copy($ws.Range('A1'))})
 Invoke-HeaderCall {$scratch.Close($false)};$script:scratch=$null
}
try {
 $app=New-Object -ComObject Excel.Application
 $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.EnableEvents=$false
 [uint32]$testProcessId=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$testProcessId)
 for($startup=0;$startup -lt 8;$startup++){[ExcelTestWindow]::DismissReminder($testProcessId);Start-Sleep -Milliseconds 250}
 $root=Join-Path (Get-Location) '.local/xlsx-header-entry';$sourcePath=Join-Path $root 'source.xlsx'
 $source=Invoke-HeaderCall {$app.Workbooks.Open($sourcePath,0,$true)};$ws=Invoke-HeaderCall {$source.Worksheets.Item('Sales data')}
 $results=@();$mismatches=@();$localizations=@()
 foreach($stage in @('typed','duplicate','paste','swap','blank','clear','escaped','long','literal')) {
  switch($stage){
   'typed' {Invoke-HeaderCall {$ws.Range('B1').Value2='Units sold'}}
   'duplicate' {Invoke-HeaderCall {$ws.Range('B1').Value2='Price'}}
   'paste' {Copy-Headers @('Price','Price','Price2') $true}
   'swap' {Copy-Headers @('Price3','Price','Price2')}
   'blank' {Invoke-HeaderCall {$ws.Range('B1').ClearContents()}}
   'clear' {Invoke-HeaderCall {$ws.Range('A1:C1').ClearContents()}}
   'escaped' {Invoke-HeaderCall {$ws.Range('B1').Value2="Net, [#@']"}}
   'long' {Invoke-HeaderCall {$ws.Range('B1').Value2=('X'*260)}}
   'literal' {Invoke-HeaderCall {$ws.Range('B1').Value2="'=Heading"}}
  }
  # The application UI is English; this workstation's Office UI is German.
  # Record and translate only automatically generated blank headings, then let Excel repair references.
  if($stage -in @('blank','clear')) {foreach($c in $ws.ListObjects.Item(1).ListColumns){
   $label=[string]$c.Name
   if($label -match '^Spalte(\d+)$'){$english='Column'+$Matches[1];$localizations += @{stage=$stage;from=$label;to=$english};Invoke-HeaderCall {$c.Name=$english}}
  }}
  Invoke-HeaderCall {$app.CalculateFull()};$expected=Read-HeaderBook $source
  $path=Join-Path $root "$stage.xlsx";$actual=Invoke-HeaderCall {$app.Workbooks.Open($path,0,$true)}
  Invoke-HeaderCall {$app.CalculateFull()};$observed=Read-HeaderBook $actual
  foreach($key in @('columns','names','validation','conditional','note')){
   if(($expected[$key]|ConvertTo-Json -Depth 8 -Compress) -cne ($observed[$key]|ConvertTo-Json -Depth 8 -Compress)){$mismatches += @{stage=$stage;part=$key;expected=$expected[$key];actual=$observed[$key]}}
  }
  for($i=0;$i -lt $expected.cells.Count;$i++){foreach($key in @('value','formula','format')){
   if([string]$expected.cells[$i][$key] -cne [string]$observed.cells[$i][$key]){$mismatches += @{stage=$stage;ref=$expected.cells[$i].ref;part=$key;expected=$expected.cells[$i][$key];actual=$observed.cells[$i][$key]}}
  }}
  $results += @{stage=$stage;sha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash;expected=$expected;actual=$observed}
  Invoke-HeaderCall {$actual.Close($false)};$actual=$null
 }
 @{excelVersion=[string]$app.Version;excelBuild=[string]$app.Build;sourceSha256=(Get-FileHash -Algorithm SHA256 -LiteralPath $sourcePath).Hash;stages=$results;localizations=$localizations;cellSnapshots=432;mismatches=$mismatches} | ConvertTo-Json -Depth 16 | Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
 Write-Output "Native Excel header entry: $($mismatches.Count) mismatches across 432 cell snapshots; $($localizations.Count) generated headings translated from German to English."
 if($mismatches.Count){throw 'Header entry comparison failed; see native-report.json.'}
} finally {if($scratch){Invoke-HeaderCall {$scratch.Close($false)}};if($actual){Invoke-HeaderCall {$actual.Close($false)}};if($source){Invoke-HeaderCall {$source.Close($false)}};if($app){Invoke-HeaderCall {$app.Quit()}}}
