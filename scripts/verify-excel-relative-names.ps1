param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/excel-relative-names'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Hash($p) {(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}
function WriteJson($path,$value) {[IO.File]::WriteAllText($path,($value|ConvertTo-Json -Depth 16),(New-Object Text.UTF8Encoding($false)))}
function Snapshot($book) {
 $rows=@()
 foreach($sheetName in @('Calls','Cost data')) {
  $sheet=$book.Worksheets.Item($sheetName)
  foreach($row in @(2,4)){foreach($col in @('D','E','F','G','H','I','J','K','L','M','N')) {
   $cell=$sheet.Range("$col$row");$value=$cell.Value2
   $rows += [ordered]@{sheet=$sheetName;address="$col$row";formula=[string]$cell.Formula;value=$value;valueType=$value.GetType().FullName;text=[string]$cell.Text}
  }}
 }
 foreach($address in @('A1','A6','A7','A8','D3')) {
  $cell=$book.Worksheets.Item('Calls').Range($address);$value=$cell.Value2
  $rows += [ordered]@{sheet='Calls';address=$address;formula=[string]$cell.Formula;value=$value;valueType=$(if($null -eq $value){'blank'}else{$value.GetType().FullName});text=[string]$cell.Text}
 }
 return ,$rows
}
function Names($book) {
 $rows=@();foreach($n in $book.Names){$rows += [ordered]@{name=[string]$n.Name;a1=[string]$n.RefersTo;r1c1=[string]$n.RefersToR1C1}}
 return ,$rows
}
function Stage($book,$stage) {
 if($stage -eq 'source'){return}
 $data=$book.Worksheets.Item('Data');$calls=$book.Worksheets.Item('Calls')
 if($stage -eq 'edit'){$data.Range('A3').Value2=[double]500}
 if($stage -in @('delete','caller')){[void]$data.Range('A3').ClearContents()}
 if($stage -eq 'caller'){$calls.Range('D2').Formula='=1/RowNext';$calls.Range('D3').Formula='=RowNext'}
}
$app=$null;$book=$null;$failure=$null;$keeper=$null
try {
 $app=New-Object -ComObject Excel.Application
 if($app.Workbooks.Count -ne 0){throw 'Native oracle requires an idle Excel instance.'}
 $alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity;$links=$app.AskToUpdateLinks
 $app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.AskToUpdateLinks=$false
 $baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json')|ConvertFrom-Json
 $native=$baseline.applications|Where-Object executable -eq 'EXCEL.EXE'
 if((Hash $native.path) -ne $native.sha256){throw 'Excel baseline changed.'}
 [uint32]$owned=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$owned);[ExcelTestWindow]::DismissReminder($owned)
 Write-Output "Owned Excel process $owned"
 $source=Join-Path $root 'tests/fixtures/excel-relative-names.xlsx'
 if(!$Compare) {
  $book=$app.Workbooks.Add();while($book.Worksheets.Count -lt 3){[void]$book.Worksheets.Add()}
  $book.Worksheets.Item(1).Name='Data';$book.Worksheets.Item(2).Name='Calls';$book.Worksheets.Item(3).Name='Cost data'
  $data=$book.Worksheets.Item('Data')
  for($r=1;$r -le 12;$r++){for($c=1;$c -le 20;$c++){$data.Cells.Item($r,$c).Value2=[double]($r*100+$c)}}
  [void]$data.Activate();[void]$data.Range('D5').Select()
  [void]$book.Names.Add('Forward','=Data!B2')
  [void]$book.Names.Add('RowNext','=Data!$A2')
  [void]$book.Names.Add('ColNext','=Data!B$1')
  [void]$book.Names.Add('Area','=Data!$A2:$B3')
  [void]$book.Names.Add('Behind','=Data!XFD1048576')
  [void]$book.Names.Add('AliasForward','=Forward*2')
  [void]$book.Names.Add('Previous','=Calls!A1048576')
  [void]$data.Range('H9').Select()
  [void]$book.Names.Add('OtherAnchor','=Data!B2')
  [void]$book.Worksheets.Item('Cost data').Names.Add("'Cost data'!Forward",'=Data!$B2')
  $anchors=@()
  foreach($active in @('A1','D5','H9')){[void]$data.Range($active).Select();$anchors+=@{activeSheet='Data';activeCell=$active;names=(Names $book)}}
  foreach($sheetName in @('Calls','Cost data')) {
   $sheet=$book.Worksheets.Item($sheetName)
   foreach($r in @(2,4)) {
    $formulas=@('=RowNext','=ColNext','=Forward','=OtherAnchor','=Behind','=AliasForward','=SUM(Area)',"='Cost data'!Forward",'=ROW(RowNext)','=COLUMN(ColNext)','=IF(RowNext>0,RowNext,999)')
    for($i=0;$i -lt $formulas.Count;$i++){$sheet.Cells.Item($r,$i+4).Formula=$formulas[$i]}
   }
  }
  $calls=$book.Worksheets.Item('Calls');$calls.Range('A1').Formula='=Previous';$calls.Range('A6').Value2=[double]1
  $calls.Range('A7').Formula='=Previous+1';$calls.Range('A8').Formula='=Previous+1'
  [void]$calls.Activate();[void]$calls.Range('J8').Select()
  $app.CalculateFullRebuild();$book.SaveAs((Join-Path $directory 'source.xlsx'),51);$book.Close($false);$book=$null
  # Sanitize only our authored package after Save (Excel overwrites creator during Save).
  & python (Join-Path $PSScriptRoot 'prepare-excel-relative-names.py')
  if($LASTEXITCODE -ne 0){throw 'Fixture preparation failed.'}
  Copy-Item -LiteralPath (Join-Path $directory 'source.xlsx') -Destination $source
  $stages=@()
  foreach($stage in @('source','edit','delete','caller')) {
   $book=$app.Workbooks.Open($source,0,$false);Stage $book $stage;$app.CalculateFullRebuild()
   [void]$book.Worksheets.Item('Data').Activate();[void]$book.Worksheets.Item('Data').Range('A1').Select()
   $stages+=@{stage=$stage;cells=(Snapshot $book);names=(Names $book)}
   $book.SaveAs((Join-Path $directory "expected-$stage.xlsx"),51);$book.Close($false);$book=$null
  }
  $report=@{sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;version=$native.version;executableSha256=$native.sha256;anchors=$anchors;stages=$stages}
  WriteJson (Join-Path $root 'tests/fixtures/native-excel-relative-names.json') $report
  Write-Output 'Authored native relative names at two definition cells; captured three selection displays and four recalculated stages.'
 } else {
  # Retain an owned unsaved workbook while switching files. This keeps the
  # automation instance alive when Excel closes its last document window.
  $keeper=$app.Workbooks.Add()
  $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json')|ConvertFrom-Json
  if(!$browser.passed -or (Hash $source) -ne $browser.sourceSha256){throw 'Stale browser source.'}
  $reports=@()
  foreach($stage in @('edit','delete','caller')) {
   Write-Output "Comparing relative-name stage $stage"
   $output=Join-Path $directory "$stage.xlsx"
   if((Hash $output) -ne $browser.exports.$stage){throw 'Stale browser export.'}
   $book=$app.Workbooks.Open($source,0,$false);Stage $book $stage;$app.CalculateFullRebuild()
   [void]$book.Worksheets.Item('Data').Activate();[void]$book.Worksheets.Item('Data').Range('A1').Select()
   $expected=Snapshot $book;$expectedNames=Names $book
   $expectedPath=Join-Path $directory "expected-$stage.xlsx";$book.SaveAs($expectedPath,51);$book.Close($false);$book=$null
   $book=$app.Workbooks.Open($output,0,$true);$app.CalculateFullRebuild()
   [void]$book.Worksheets.Item('Data').Activate();[void]$book.Worksheets.Item('Data').Range('A1').Select()
   $actual=Snapshot $book;$actualNames=Names $book
   WriteJson (Join-Path $directory "comparison-$stage.json") @{expected=$expected;actual=$actual}
   if(($expected|ConvertTo-Json -Compress) -cne ($actual|ConvertTo-Json -Compress)){throw "Native cells differ: $stage"}
   if(($expectedNames|ConvertTo-Json -Compress) -cne ($actualNames|ConvertTo-Json -Compress)){throw "Native names differ: $stage"}
   $reports+=@{stage=$stage;exportSha256=Hash $output;expectedSha256=Hash $expectedPath;cells=$actual;names=$actualNames}
   $book.Close($false);$book=$null
  }
  WriteJson (Join-Path $directory 'native-report.json') @{passed=$true;sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;executableSha256=$native.sha256;stages=$reports}
  Write-Output 'All 147 native cell snapshots and nine name definitions match independent edits in three actual browser exports.'
 }
} catch {
 $failure=$_
 Write-Output ("Native operation failed: " + $_.ToString() + " at " + $_.ScriptStackTrace)
 throw
} finally {
 $cleanup=@()
 if($book){try{$book.Close($false)}catch{$cleanup+= $_}}
 if($keeper){try{$keeper.Close($false)}catch{$cleanup+= $_}}
 if($app){
  try{if($null -ne $security){$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;$app.AskToUpdateLinks=$links;if($app.Workbooks.Count -eq 0){$app.Quit()}}}catch{$cleanup+= $_}
  try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+= $_}
 }
 if($cleanup.Count){if($failure){Write-Warning ($cleanup|Out-String)}else{throw $cleanup[0]}}
}
