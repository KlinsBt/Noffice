param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/excel-scoped-names'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Hash($p) {(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}
function Snapshot($book) {
 $rows=@()
 foreach($sheet in $book.Worksheets) {
  foreach($address in @('A1','B1','C1','D1','E1','F1')) {
   $cell=$sheet.Range($address)
   $raw=[string]$cell.Formula
   # Native filenames with hyphens are quoted; both forms refer to OOXML's same [0] workbook.
   $formula=$raw.Replace([string]$book.Name,'[0]').Replace("'[0]'!",'[0]!')
   $value=$cell.Value2;$valueType=if($null -eq $value){'blank'}else{$value.GetType().FullName}
   $rows+= [ordered]@{sheet=[string]$sheet.Name;address=$address;rawFormula=$raw;formula=$formula;value=$value;valueType=$valueType}
  }
 }
 return ,$rows
}
function Names($book) {
 $result=@();foreach($name in $book.Names){$result += [ordered]@{name=[string]$name.Name;formula=[string]$name.RefersTo}}
 return ,$result
}
$app=$null;$book=$null
try {
 $app=New-Object -ComObject Excel.Application
 if($app.Workbooks.Count -ne 0){throw 'Native oracle requires an idle Excel instance.'}
 $alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity;$links=$app.AskToUpdateLinks
 $app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.AskToUpdateLinks=$false
 $baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json') | ConvertFrom-Json
 $native=$baseline.applications | Where-Object executable -eq 'EXCEL.EXE'
 if((Hash $native.path) -ne $native.sha256){throw 'Excel baseline changed.'}
 [uint32]$owned=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$owned)
 [ExcelTestWindow]::DismissReminder($owned)
 if(!$Compare) {
  $book=$app.Workbooks.Add()
  while($book.Worksheets.Count -lt 3){[void]$book.Worksheets.Add()}
  $book.Worksheets.Item(1).Name='Data';$book.Worksheets.Item(2).Name='Cost data';$book.Worksheets.Item(3).Name='Summary'
  $book.Worksheets.Item(1).Range('A1').Value2=[double]2
  $book.Worksheets.Item(2).Range('A1').Value2=[double]7
  $book.Worksheets.Item(3).Range('A1').Value2=[double]11
  [void]$book.Names.Add('Rate','=Data!$A$1')
  [void]$book.Names.Add('AliasRate','=Rate*3')
  [void]$book.Worksheets.Item(2).Names.Add("'Cost data'!Rate","='Cost data'!`$A`$1")
  [void]$book.Worksheets.Item(1).Names.Add('Data!Rate','=Summary!$A$1')
  $source=Join-Path $directory 'source.xlsx'
  $book.SaveAs($source,51)
  foreach($sheet in $book.Worksheets) {
   $sheet.Range('B1').Formula='=Rate'
   $sheet.Range('C1').Formula='=AliasRate'
   $sheet.Range('D1').Formula="='Cost data'!Rate"
   $sheet.Range('E1').Formula='=source.xlsx!Rate'
   $sheet.Range('F1').Formula='=SUM(Rate,AliasRate)'
  }
  foreach($key in @('Author','Last Author','Company','Manager')) {try{$book.BuiltinDocumentProperties.Item($key).Value=''}catch{}}
  $app.CalculateFullRebuild();$book.Save();$book.Close($false);$book=$null
  $book=$app.Workbooks.Open($source,0,$true)
  $names=@();foreach($name in $book.Names){$names+=@{name=[string]$name.Name;refersTo=[string]$name.RefersTo}}
  $report=[ordered]@{passed=$true;sourceSha256=Hash $source;version=$native.version;scriptSha256=Hash $PSCommandPath;names=$names;cells=(Snapshot $book)}
  [IO.File]::WriteAllText((Join-Path $directory 'source-report.json'),($report|ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
  Write-Output 'Authored/reopened native scoped names, aliases and workbook qualification.'
 } else {
  $source=Join-Path $root 'tests/fixtures/excel-scoped-names.xlsx'
  $output=Join-Path $directory 'browser.xlsx'
  $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json')|ConvertFrom-Json
  if(!$browser.passed -or (Hash $source) -ne $browser.sourceSha256 -or (Hash $output) -ne $browser.exportSha256){throw 'Stale browser evidence.'}
  $stages=@()
  foreach($stage in @('edit','delete')) {
   $output=Join-Path $directory $(if($stage -eq 'edit'){'browser.xlsx'}else{'deleted.xlsx'})
   $boundHash=if($stage -eq 'edit'){$browser.exportSha256}else{$browser.deleteSha256}
   if((Hash $output) -ne $boundHash){throw 'Stale stage export.'}
   $book=$app.Workbooks.Open($source,0,$false)
   if($stage -eq 'edit'){$book.Worksheets.Item('Data').Range('A1').Value2=[double]5}else{[void]$book.Worksheets.Item('Data').Range('A1').ClearContents()}
   $app.CalculateFullRebuild();$expected=Snapshot $book;$expectedNames=Names $book
   $expectedPath=Join-Path $directory "expected-$stage.xlsx"
   $book.SaveAs($expectedPath,51);$book.Close($false);$book=$null
   $book=$app.Workbooks.Open($output,0,$true);$app.CalculateFullRebuild();$actual=Snapshot $book;$actualNames=Names $book
   [IO.File]::WriteAllText((Join-Path $directory "comparison-$stage.json"),(@{expected=$expected;actual=$actual}|ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
   for($i=0;$i -lt $expected.Count;$i++){if($expected[$i].value -ne $actual[$i].value -or $expected[$i].valueType -ne $actual[$i].valueType -or $expected[$i].formula -cne $actual[$i].formula){throw "Native mismatch: $stage $($expected[$i].sheet)!$($expected[$i].address)"}}
   if(($expectedNames|ConvertTo-Json -Compress) -cne ($actualNames|ConvertTo-Json -Compress)){throw 'Native name definitions differ.'}
   $stages+=@{stage=$stage;exportSha256=Hash $output;expectedSha256=Hash $expectedPath;expected=$expected;actual=$actual;names=$actualNames}
   $book.Close($false);$book=$null
  }
  $report=@{passed=$true;sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;stages=$stages}
  [IO.File]::WriteAllText((Join-Path $directory 'native-report.json'),($report|ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
  Write-Output 'All 36 scoped-name cell snapshots and four name definitions match independent native edit/delete stages.'
 }
} finally {
 if($book){$book.Close($false)}
 if($app){if($null -ne $security){$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;$app.AskToUpdateLinks=$links;if($app.Workbooks.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
