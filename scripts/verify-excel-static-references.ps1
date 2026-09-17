param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/excel-static-references'))
[void][IO.Directory]::CreateDirectory($root)
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Hash($path){(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Call([scriptblock]$Action){
 for($attempt=0;$attempt -lt 40;$attempt++){
  try{return (& $Action)}catch{
   $e=$_.Exception;$busy=$false
   while($e){if($e.HResult -in @(-2147418111,-2147417846,-2146777998)){$busy=$true};$e=$e.InnerException}
   if(!$busy -or $attempt -eq 39){throw}
   if($ownedProcess){[ExcelTestWindow]::DismissReminder($ownedProcess)}
   Start-Sleep -Milliseconds 250
  }
 }
}
function Snapshot($book){
 if(!$book){throw 'Excel returned no workbook; refusing to compare missing native values.'}
 $result=@()
 foreach($sheetName in @('Data','Cost data','Summary')){
  $sheet=Call {$book.Worksheets.Item($sheetName)}
  $addresses=if($sheetName -eq 'Summary'){@((1..16|ForEach-Object{"A$_"})+(1..6|ForEach-Object{"B$_"}))}elseif($sheetName -eq 'Data'){@('A1','A2','B1','B2','D1')}else{@('A1')}
  foreach($address in $addresses){
   $observed=$false
   for($read=0;$read -lt 40;$read++){
    $cell=Call {$sheet.Range($address)}
    if($cell){
     $value=Call {$cell.Value2};$formula=Call {$cell.Formula}
     # A populated Excel cell cannot have a missing Formula property. Do not
     # cast a transient null native read to an empty string and certify it.
     if($null -eq $value -or ![string]::IsNullOrEmpty([string]$formula)){$observed=$true;break}
    }
    if($ownedProcess){[ExcelTestWindow]::DismissReminder($ownedProcess)}
    Start-Sleep -Milliseconds 250
   }
   if(!$observed){throw "Missing native cell properties: $sheetName!$address"}
   $result += [ordered]@{sheet=$sheetName;address=$address;formula=[string]$formula;value=$value;valueType=$(if($null -eq $value){'blank'}else{$value.GetType().FullName})}
  }
 }
 return ,$result
}
function Edit($book,$stage){
 $data=Call {$book.Worksheets.Item('Data')}
 if($stage -eq 'edit'){Call {$data.Range('A1').Value2=[double]5}}
 elseif($stage -eq 'delete'){Call {$data.Range('A1').ClearContents()}|Out-Null}
 elseif($stage -eq 'branch'){Call {$data.Range('B2').Value2=[double]9}}
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
$app=$null;$book=$null;$keeper=$null;$security=$null;$links=$null;[uint32]$ownedProcess=0
try{
 $app=New-Object -ComObject Excel.Application
 [void][ExcelTestWindow]::GetWindowThreadProcessId([IntPtr](Call {$app.Hwnd}),[ref]$ownedProcess)
 if((Call {$app.Workbooks.Count}) -ne 0){throw 'Native oracle requires an idle owned instance.'}
 $alerts=Call {$app.DisplayAlerts};$security=Call {$app.AutomationSecurity};$links=Call {$app.AskToUpdateLinks}
 Call {$app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=3;$app.AskToUpdateLinks=$false}
 $baseline=Get-Content (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
 $exeHash=Hash (Join-Path ([string](Call {$app.Path})) 'EXCEL.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'EXCEL.EXE').sha256){throw 'Excel baseline changed'}
 [ExcelTestWindow]::DismissReminder($ownedProcess)
 # Keep the owned automation session alive while each comparison workbook closes.
 # Excel 2016 can otherwise return a null workbook on a subsequent Open.
 $keeper=Call {$app.Workbooks.Add()}
 if(!$keeper){throw 'Excel did not create the owned session keeper.'}
 $source=Join-Path $root 'source.xlsx'
 if(!$Compare){
  $book=Call {$app.Workbooks.Add()}
  while((Call {$book.Worksheets.Count}) -lt 3){Call {$book.Worksheets.Add()}|Out-Null}
  Call {$book.Worksheets.Item(1).Name='Data';$book.Worksheets.Item(2).Name='Cost data';$book.Worksheets.Item(3).Name='Summary'}
  $data=Call {$book.Worksheets.Item('Data')};$summary=Call {$book.Worksheets.Item('Summary')}
  Call {$data.Range('A1').Value2=[double]2;$data.Range('A2').Value2=[double]4;$data.Range('B1').Value2=[double]3;$data.Range('B2').Value2=[double]5;$data.Range('D1').Formula='=IF(A1>0,A2,B2)';$book.Worksheets.Item('Cost data').Range('A1').Value2=[double]7}
  $formulas=@('=Data!$A$1','=Data!$A1+Data!A$2',"='Cost data'!`$A`$1",'=SUM(Data!$A1:B$2)','=IF(Data!A1>0,Data!A2,Data!B2)','=IF(Data!A1>0,Data!A1,1/0)','=SUM(Data!A:A)','=SUM(Data!1:2)','=IF(FALSE,Data!B2,42)', '=A1+A4','=SUM(Data!$1:$2)','=IF(TRUE,Data!C1,1)','=ISBLANK(Data!$C$1)','=COUNTBLANK(Data!$C$1)','=IF(TRUE,"",Data!A1)','=IF(TRUE,Summary!A15,1)')
  for($i=0;$i -lt $formulas.Count;$i++){Call {$summary.Range("A$($i+1)").Formula=$formulas[$i]}}
  $blankChecks=@('=ISNUMBER(Data!C1)','=ISTEXT(Data!C1)','=PRODUCT(Data!C1,2)','=ISBLANK(A15)','=COUNTBLANK(A15)','=ISNUMBER(Data!A1)')
  for($i=0;$i -lt $blankChecks.Count;$i++){Call {$summary.Range("B$($i+1)").Formula=$blankChecks[$i]}}
  foreach($key in @('Author','Last Author','Company','Manager')){try{$book.BuiltinDocumentProperties.Item($key).Value='Noffice'}catch{}}
  Call {$app.CalculateFullRebuild();$book.SaveAs($source,51);$book.Close($false)};$book=$null
  SanitizeFixture $source
  $book=Call {$app.Workbooks.Open($source,0,$true)};Call {$app.CalculateFullRebuild()};$initial=Snapshot $book
  Call {$book.Worksheets.Item('Data').Activate()}
  $precedents=[string](Call {$book.Worksheets.Item('Data').Range('D1').DirectPrecedents.Address($false,$false)})
  Call {$book.Close($false)};$book=$null
  $stages=@()
  foreach($stage in @('edit','delete','branch')){
   $book=Call {$app.Workbooks.Open($source,0,$true)};Edit $book $stage;Call {$app.CalculateFullRebuild()}
   $stages+=@{stage=$stage;cells=Snapshot $book};Call {$book.Close($false)};$book=$null
  }
  @{passed=$true;sourceHash=Hash $source;initial=$initial;stages=$stages;conditionalPrecedents=$precedents;executableHash=$exeHash;scriptHash=Hash $PSCommandPath}|ConvertTo-Json -Depth 12|Set-Content -Encoding UTF8 (Join-Path $root 'reference.json')
  Write-Output "Native reference captured: 28 cells, 3 stages; conditional direct precedents $precedents."
 }else{
  $receipt=Get-Content (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
  if(!$receipt.passed -or (Hash $source) -ne $receipt.sourceHash){throw 'Stale browser source evidence'}
  $comparisons=@()
  foreach($stage in @('edit','delete','branch')){
   $output=Join-Path $root "browser-$stage.xlsx"
   if((Hash $output) -ne ($receipt.exports|Where-Object stage -eq $stage).sha256){throw 'Stale browser export'}
   $book=Call {$app.Workbooks.Open($source,0,$false)};Edit $book $stage;Call {$app.CalculateFullRebuild()};$expected=Snapshot $book
   $expectedPath=Join-Path $root "expected-$stage.xlsx"
   Call {$book.SaveAs($expectedPath,51);$book.Close($false)};$book=$null
   $book=Call {$app.Workbooks.Open($output,0,$true)};Call {$app.CalculateFullRebuild()};$actual=Snapshot $book
   @{expected=$expected;actual=$actual}|ConvertTo-Json -Depth 10|Set-Content -Encoding UTF8 (Join-Path $root "comparison-$stage.json")
   for($i=0;$i -lt $expected.Count;$i++){
    foreach($field in @('sheet','address','formula','value','valueType')){if($expected[$i].$field -cne $actual[$i].$field){throw "Native mismatch: $stage $($expected[$i].sheet)!$($expected[$i].address) $field"}}
   }
   $comparisons+=@{stage=$stage;exportHash=Hash $output;expectedHash=Hash $expectedPath;actual=$actual}
   Call {$book.Close($false)};$book=$null
  }
  @{passed=$true;sourceHash=Hash $source;comparisons=$comparisons;executableHash=$exeHash;scriptHash=Hash $PSCommandPath}|ConvertTo-Json -Depth 12|Set-Content -Encoding UTF8 (Join-Path $root 'native-report.json')
  Write-Output 'Native static-reference comparison: 84 cell snapshots match values/types/formulas.'
 }
}finally{
 try{if($book){Call {$book.Close($false)}}}finally{try{if($keeper){Call {$keeper.Close($false)}}}finally{if($app){try{if($null -ne $security){Call {$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if($null -ne $links){$app.AskToUpdateLinks=$links};if($app.Workbooks.Count -eq 0){$app.Quit()}}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}}
}
