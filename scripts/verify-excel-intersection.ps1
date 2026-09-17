param([switch]$Compare)
# Authored scalar-reference contract; compare real browser exports with native edits.
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/excel-intersection'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
function WriteJson($name,$value){[IO.File]::WriteAllText((Join-Path $directory $name),($value|ConvertTo-Json -Depth 16),(New-Object Text.UTF8Encoding($false)))}
$formulas=[ordered]@{
 C1='=Data!A1:A8';C2='=Data!A1:A5';C3='=Data!A1:A5+1';C4='=SQRT(Data!A1:A5)';C6='=Data!A1:A8';C7='=Data!A1:A8';C8='=Data!A1:A8';C9='=Data!A1:A5'
 B10='=Data!A1:C1';F10='=Data!A1:C1';B2='=Data!A1:B4';E2='=Data!A1:B4';D2='=SUM(Data!A1:A5)';E3='=SUM(Data!A1:A5+1)'
 D3='=INDEX(Data!A1:A5,0)';D4='=INDIRECT("Data!A1:A5")';D5='=ColumnValues';E4='=Data!A:A';E10='=Data!1:1'
 D6='=ISNUMBER(Data!A1:A8)';E6='=ISTEXT(Data!A1:A8)';F6='=ISBLANK(Data!A1:A8)'
 D7='=ISNUMBER(Data!A1:A8)';E7='=ISTEXT(Data!A1:A8)';F7='=ISBLANK(Data!A1:A8)'
 G2='=IFERROR(Data!A1:A8,99)';G8='=IFERROR(Data!A1:A8,99)';H2='=ISERROR(Data!A1:A8)';H8='=ISERROR(Data!A1:A8)'
 I7='=PMT(Data!A6:A7,1,100)';I2='=NextColumn';I6='=NextColumn';I3='=IFERROR(RelativeArea,0)';D8='=ISNUMBER(Data!A1:A8)';E8='=ISTEXT(Data!A1:A8)';F8='=ISBLANK(Data!A1:A8)'
 F2='=SUM(IF(TRUE,Data!A1:A5,0))';F3='=SUM(IFERROR(Data!A1:A5,0))';G3='=AVERAGE(Data!A1:A5)'
}
function Snapshot($book){
 $rows=@();$sheet=$book.Worksheets.Item('Calls')
 foreach($key in $formulas.Keys){
  $cell=$sheet.Range($key);$value=$cell.Value2;$error=$app.WorksheetFunction.IsError($cell)
  if($error){$code=[int]([long]$value -band 65535);$value=@{2000='#NULL!';2007='#DIV/0!';2015='#VALUE!';2023='#REF!';2029='#NAME?';2036='#NUM!';2042='#N/A'}[$code];if($null -eq $value){throw "Unknown native error $code"}}
  $rows+=,[ordered]@{ref=$key;formula=[string]$cell.Formula;value=$value;kind=$(if($error){'error'}else{'value'})}
 };return ,$rows
}
function SanitizeFixture($path){
 # Excel overwrites the author properties when saving. Sanitize the closed owned
 # package before capturing its reference hash; do not alter Office user settings.
 Add-Type -AssemblyName System.IO.Compression
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $zip=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
 try{
  foreach($part in @('docProps/core.xml','docProps/app.xml')){
   $entry=$zip.GetEntry($part);if(!$entry){continue}
   $reader=New-Object IO.StreamReader($entry.Open())
   try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
   foreach($node in $xml.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy" or local-name()="Company" or local-name()="Manager"]')){$node.InnerText='Noffice'}
   $entry.Delete();$replacement=$zip.CreateEntry($part)
   $writer=New-Object IO.StreamWriter($replacement.Open(),(New-Object Text.UTF8Encoding($false)))
   try{$writer.Write($xml.OuterXml)}finally{$writer.Dispose()}
  }
 }finally{$zip.Dispose()}
}
$app=$null;$book=$null
try{
 $app=New-Object -ComObject Excel.Application
 if($app.Workbooks.Count -ne 0){throw 'This probe requires an idle owned Excel instance.'}
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.AskToUpdateLinks
 $app.AutomationSecurity=3;$app.DisplayAlerts=$false;$app.AskToUpdateLinks=$false
 $baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json')|ConvertFrom-Json
 $native=$baseline.applications|Where-Object executable -eq 'EXCEL.EXE'
 if((Hash (Join-Path ([string]$app.Path) 'EXCEL.EXE')) -ne $native.sha256){throw 'The running Excel baseline changed.'}
 [uint32]$owned=0;[void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$owned);[ExcelTestWindow]::DismissReminder($owned)
 if(!$Compare){
 $book=$app.Workbooks.Add();while($book.Worksheets.Count -lt 2){[void]$book.Worksheets.Add()}
 $book.Worksheets.Item(1).Name='Data';$book.Worksheets.Item(2).Name='Calls'
 $data=$book.Worksheets.Item('Data');$calls=$book.Worksheets.Item('Calls')
 foreach($row in 1..5){$data.Range("A$row").Value2=[double]($row*$row)}
 $data.Range('A7').Formula='=""';$data.Range('A8').Formula='=1/0';$data.Range('B1').Value2=[double]6;$data.Range('C1').Value2=[double]7;$data.Range('B2').Formula='=FALSE()'
 [void]$book.Names.Add('ColumnValues','=Data!$A$1:$A$5')
 [void]$data.Activate();[void]$data.Range('A1').Select()
 [void]$book.Names.Add('NextColumn','=Data!$A1:$A5');[void]$book.Names.Add('RelativeArea','=Data!$A2:$B3')
 foreach($key in $formulas.Keys){$calls.Range($key).Formula=$formulas[$key]}
 $app.CalculateFullRebuild();$initial=Snapshot $book
 $source=Join-Path $directory 'source.xlsx';$book.SaveAs($source,51)
 $stages=@()
 foreach($stage in @('edit','delete','error')){
  if($stage -eq 'edit'){$data.Range('A2').Value2=[double]64}
  elseif($stage -eq 'delete'){[void]$data.Range('A2').ClearContents()}
  else{$data.Range('A2').Formula='=1/0'}
  $app.CalculateFullRebuild();$stages+=,[ordered]@{stage=$stage;cells=(Snapshot $book)}
  $book.SaveCopyAs((Join-Path $directory "native-$stage.xlsx"))
 }
 $book.Close($false);$book=$null
 SanitizeFixture $source
 $book=$app.Workbooks.Open($source,0,$true);$app.CalculateFullRebuild();$reopened=Snapshot $book
 if(($initial|ConvertTo-Json -Compress) -cne ($reopened|ConvertTo-Json -Compress)){throw 'Native saved-file scalar results changed on reopening.'}
 WriteJson 'reference.json' ([ordered]@{scope='Bounded legacy scalar intersection; full reference/array/table semantics remain open';scriptSha256=Hash $PSCommandPath;sourceSha256=Hash $source;executableSha256=$native.sha256;locale=$app.LanguageSettings.LanguageID(2);initial=$initial;stages=$stages;reopened=$reopened})
 $book.Close($false);$book=$null
  Copy-Item -LiteralPath $source -Destination (Join-Path $root 'tests/fixtures/excel-intersection.xlsx')
  Copy-Item -LiteralPath (Join-Path $directory 'reference.json') -Destination (Join-Path $root 'tests/fixtures/native-excel-intersection.json')
 }else{
  $fixture=Join-Path $root 'tests/fixtures/excel-intersection.xlsx'
  $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json')|ConvertFrom-Json
  if(!$browser.passed -or (Hash $fixture) -ne $browser.sourceSha256){throw 'Stale browser source evidence.'}
  $reports=@();$keeper=$app.Workbooks.Add()
  foreach($stage in @('edit','delete','error')){
   $output=Join-Path $directory "$stage.xlsx"
   if((Hash $output) -ne $browser.exports.$stage){throw 'Stale browser export evidence.'}
   $book=$app.Workbooks.Open($fixture,0,$false);$data=$book.Worksheets.Item('Data')
   if($stage -eq 'edit'){$data.Range('A2').Value2=[double]64}elseif($stage -eq 'delete'){[void]$data.Range('A2').ClearContents()}else{$data.Range('A2').Formula='=1/0'}
   $app.CalculateFullRebuild();$expected=Snapshot $book
   $expectedPath=Join-Path $directory "expected-$stage.xlsx";$book.SaveAs($expectedPath,51);$book.Close($false);$book=$null
   $book=$app.Workbooks.Open($output,0,$true);$app.CalculateFullRebuild();$actual=Snapshot $book
   WriteJson "comparison-$stage.json" @{expected=$expected;actual=$actual}
   if(($expected|ConvertTo-Json -Compress) -cne ($actual|ConvertTo-Json -Compress)){throw "Native scalar/reference results differ: $stage"}
   $reports+=,[ordered]@{stage=$stage;exportSha256=Hash $output;expectedSha256=Hash $expectedPath;cells=$actual}
   $book.Close($false);$book=$null
  }
  $keeper.Close($false);$keeper=$null
  WriteJson 'native-report.json' ([ordered]@{passed=$true;sourceSha256=Hash $fixture;scriptSha256=Hash $PSCommandPath;executableSha256=$native.sha256;stages=$reports})
  Write-Output "All $($formulas.Count*3) native scalar/reference snapshots match three actual browser exports."
 }
 if(!$Compare){Write-Output "Recorded $($initial.Count) native scalar/range cases, three edit stages and unchanged-source reopen."}
}finally{
 if($book){$book.Close($false)}
 if($keeper){$keeper.Close($false)}
 if($app){if($null -ne $security){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.AskToUpdateLinks=$links;if($app.Workbooks.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
