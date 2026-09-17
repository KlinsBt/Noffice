$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/powerpoint-child-bounds'
[void][IO.Directory]::CreateDirectory($directory)
. (Join-Path $PSScriptRoot 'excel-test-window.ps1')
function Retry([scriptblock]$action){for($attempt=0;$attempt -lt 24;$attempt++){try{return (& $action)}catch{if($_.Exception.ToString() -notmatch 'RPC_E_CALL_REJECTED|0x80010001|0x8001010A' -or $attempt -eq 23){throw};if($script:officeProcess){[ExcelTestWindow]::DismissReminder($script:officeProcess)};Start-Sleep -Milliseconds 250}}}
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Transforms($path){
 $zip=[IO.Compression.ZipFile]::OpenRead($path);$result=@()
 try {foreach($i in 1..4){
  $reader=New-Object IO.StreamReader($zip.GetEntry("ppt/slides/slide$i.xml").Open())
  try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
  $frames=@();foreach($group in $xml.SelectNodes("//*[local-name()='grpSp']")){
   $id=$group.SelectSingleNode("./*[local-name()='nvGrpSpPr']/*[local-name()='cNvPr']").GetAttribute('id')
   $frame=[ordered]@{id=$id}
   foreach($node in $group.SelectSingleNode("./*[local-name()='grpSpPr']/*[local-name()='xfrm']").ChildNodes){foreach($attr in $node.Attributes){$frame["$($node.LocalName).$($attr.LocalName)"]=[double]$attr.Value}}
   $frames+=,$frame
  };$result+=,$frames
 };return ,$result}finally{$zip.Dispose()}
}
function Hash($p){(Get-FileHash -LiteralPath $p).Hash.ToLowerInvariant()}
function Frame($shape){return Retry { [ordered]@{id=[int]$shape.Id;name=[string]$shape.Name;x=[double]$shape.Left;y=[double]$shape.Top;w=[double]$shape.Width;h=[double]$shape.Height;rotation=[double]$shape.Rotation} }}
function Shapes($shapes,$parents){
 $result=@()
 for($shapeIndex=1;$shapeIndex -le $shapes.Count;$shapeIndex++){
  $shape=$shapes.Item($shapeIndex)
  $frame=Frame $shape;$frame['parents']=$parents;$frame['group']=($shape.Type -eq 6);$result+=,$frame
  if($shape.Type -eq 6){$result+=Shapes $shape.GroupItems @($parents+[int]$shape.Id)}
 }
 return $result
}
function Snapshot($deck){
 $result=@()
 foreach($slide in $deck.Slides){$result+=,@(Shapes $slide.Shapes @())}
 return ,$result
}
function SelectBounds($deck){
 return Retry {
 $result=@();$window=$deck.Windows.Item(1)
 foreach($slide in $deck.Slides){
  Retry {$window.View.GotoSlide($slide.SlideIndex)}
  $group=$slide.Shapes.Item(1);Retry {$group.Select(-1)}
  $result+=,(Retry {
   $selection=$window.Selection.ShapeRange
   [ordered]@{slide=[int]$slide.SlideIndex;count=[int]$selection.Count;x=[double]$selection.Left;y=[double]$selection.Top;w=[double]$selection.Width;h=[double]$selection.Height;rotation=[double]$selection.Rotation}
  })
 }
 return ,$result
 }
}
function FindLeaf($shapes,$id){foreach($shape in $shapes){if($shape.Type -eq 6){$found=FindLeaf $shape.GroupItems $id;if($found){return $found}}elseif($shape.Id -eq $id){return $shape}}}
function Ungroup($deck){foreach($slide in $deck.Slides){$more=$true;while($more){$more=$false;for($i=$slide.Shapes.Count;$i -ge 1;$i--){if($slide.Shapes.Item($i).Type -eq 6){[void]$slide.Shapes.Item($i).Ungroup();$more=$true}}}}}
function WriteJson($name,$value){[IO.File]::WriteAllText((Join-Path $directory $name),($value|ConvertTo-Json -Depth 24),(New-Object Text.UTF8Encoding($false)))}

function CompareFrames($expected,$actual){
 if($expected.Count -ne 4 -or $actual.Count -ne 4){throw 'Expected all four slides.'}
 $maximum=0
 for($i=0;$i -lt 4;$i++){
  if($expected[$i].Count -ne $actual[$i].Count -or $actual[$i].Count -lt 3){throw 'Shape count changed.'}
  for($j=0;$j -lt $expected[$i].Count;$j++){
   $p=$expected[$i][$j];$q=$actual[$i][$j]
   if($p.id -le 0 -or $p.id -ne $q.id -or $p.name -cne $q.name -or ($p.parents|ConvertTo-Json -Compress) -cne ($q.parents|ConvertTo-Json -Compress)){throw 'Native shape identity changed.'}
   if($p.w -le 0 -or $q.w -le 0 -or $p.h -le 0 -or $q.h -le 0){throw 'Invalid native dimensions.'}
   foreach($key in @('x','y','w','h','rotation')){
    $delta=[Math]::Abs($p.$key-$q.$key);$maximum=[Math]::Max($maximum,$delta)
    if($delta -gt 0.001){throw "Native frame mismatch slide $i shape $j $key ($delta)"}
   }
  }
 };return $maximum
}
function OpenOwned($path){
 $d=Retry {$app.Presentations.Open($path,0,0,0)}
 if(!$d -or $d.Slides.Count -ne 4){throw 'Native presentation is unavailable.'}
 $w=$d.NewWindow();Retry {$w.WindowState=2};return $d
}
$app=$null;$deck=$null
try{
 $baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json')|ConvertFrom-Json
 $native=$baseline.applications|Where-Object executable -eq 'POWERPNT.EXE'
 $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json')|ConvertFrom-Json
 $uiRoot=Join-Path $root '.local/powerpoint-child-ui'
 $ui=Get-Content -Raw (Join-Path $uiRoot 'ui-edit-report.json')|ConvertFrom-Json
 $recorded=Get-Content -Raw (Join-Path $root 'tests/fixtures/native-powerpoint-child-ui.json')|ConvertFrom-Json
 $source=Join-Path $root 'tests/fixtures/powerpoint-child-bounds.pptx'
 if(!$browser.passed -or $browser.sourceSha256 -ne (Hash $source) -or $ui.sourceSha256 -ne (Hash $source) -or $ui.scriptSha256 -ne (Hash (Join-Path $root 'scripts/verify-powerpoint-child-ui.ps1')) -or $ui.exportSha256 -ne (Hash (Join-Path $uiRoot 'ui-edited.pptx')) -or $ui.pdfSha256 -ne (Hash (Join-Path $uiRoot 'ui-edited.pdf')) -or $ui.executableSha256 -ne $native.sha256){throw 'Stale native UI/browser evidence.'}
 if($ui.stages.Count -ne 14 -or $recorded.stages.Count -ne 14){throw 'Incomplete native UI history.'}
 for($i=0;$i -lt 14;$i++){
  if($ui.stages[$i].stage -cne $recorded.stages[$i].stage){throw 'Native UI stage changed.'}
  foreach($key in @('x','y','w','h','rotation')){if([Math]::Abs($ui.stages[$i].frame.$key-$recorded.stages[$i].frame.$key) -gt 0.001){throw 'Native UI oracle changed.'}}
 }
 $app=New-Object -ComObject PowerPoint.Application
 if($app.Presentations.Count -ne 0){throw 'The native test requires an idle owned PowerPoint instance.'}
 $processes=@(Get-CimInstance Win32_Process -Filter "name = 'POWERPNT.EXE'")
 if($processes.Count -ne 1){throw 'Cannot uniquely identify the owned PowerPoint process.'}
 $script:officeProcess=[uint32]$processes[0].ProcessId
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$app.AutomationSecurity=3;$app.DisplayAlerts=1
 if((Hash (Join-Path ([string]$app.Path) 'POWERPNT.EXE')) -ne $native.sha256){throw 'The native baseline changed.'}
 $results=@()
 foreach($stage in @('edits','ui')){
  $actualPath=Join-Path $directory "$stage.pptx"
  if((Hash $actualPath) -ne $browser."${stage}Sha256"){throw 'Stale exported file.'}
  if($stage -eq 'ui'){$deck=OpenOwned (Join-Path $uiRoot 'ui-edited.pptx')}
  else{
   $deck=OpenOwned $source
   foreach($index in 1..4){
    $leaf=Retry {FindLeaf $deck.Slides.Item($index).Shapes 3};$leaf.LockAspectRatio=0
    foreach($property in @('Left','Top','Width','Height')){
     $delta=if($property -in @('Left','Width')){12}else{9}
     $leaf.$property=[double]$leaf.$property+$delta
    }
   }
  }
  $expectedPath=Join-Path $directory "$stage-native.pptx"
  $deck.SaveAs($expectedPath,24);$deck.Close();$deck=$null
  $snapshots=@();$hashes=@()
  foreach($pair in @(@{name='native';path=$expectedPath},@{name='browser';path=$actualPath})){
   $deck=OpenOwned $pair.path
   $before=Snapshot $deck;$selection=SelectBounds $deck
   $pdf=Join-Path $directory "$stage-$($pair.name).pdf";$deck.SaveAs($pdf,32)
   Ungroup $deck;$effective=Snapshot $deck;$deck.Saved=-1;$deck.Close();$deck=$null
   $deck=OpenOwned $pair.path
   $leaf=FindLeaf $deck.Slides.Item(3).Shapes 3;$leaf.Left=[double]$leaf.Left+6
   $reedit=Join-Path $directory "$stage-$($pair.name)-reedit.pptx"
   $deck.SaveAs($reedit,24);$deck.Close();$deck=$null
   $deck=OpenOwned $reedit
   $after=Snapshot $deck;$reeditPdf=Join-Path $directory "$stage-$($pair.name)-reedit.pdf"
   $deck.SaveAs($reeditPdf,32);$deck.Saved=-1;$deck.Close();$deck=$null
   $snapshots+=,@{before=$before;selection=$selection;effective=$effective;after=$after;transforms=(Transforms $pair.path)}
   $hashes+=,@{pptxSha256=Hash $pair.path;pdfSha256=Hash $pdf;reeditSha256=Hash $reedit;reeditPdfSha256=Hash $reeditPdf}
  }
  $maximum=0
  foreach($kind in @('before','effective','after')){$maximum=[Math]::Max($maximum,(CompareFrames $snapshots[0].$kind $snapshots[1].$kind))}
  for($i=0;$i -lt 4;$i++){
   if($snapshots[1].selection[$i].count -ne 1){throw 'Native selection count changed.'}
   foreach($key in @('x','y','w','h','rotation')){if([Math]::Abs($snapshots[0].selection[$i].$key-$snapshots[1].selection[$i].$key) -gt 0.001){throw 'Native selection frame differs.'}}
   for($j=0;$j -lt $snapshots[0].transforms[$i].Count;$j++){
    $p=$snapshots[0].transforms[$i][$j];$q=$snapshots[1].transforms[$i][$j]
    if($p.id -ne $q.id){throw 'Group XML identity changed.'}
    foreach($key in $p.Keys){if($key -ne 'id' -and [Math]::Abs($p[$key]-$q[$key]) -gt 12.7){throw "Group XML frame differs $stage $i $j $key"}}
   }
  }
  $results+=,@{stage=$stage;maximumFrameDifferencePt=$maximum;snapshots=$snapshots;hashes=$hashes}
 }
 WriteJson 'native-report.json' @{passed=$true;sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;executableSha256=$native.sha256;uiReceiptSha256=Hash (Join-Path $uiRoot 'ui-edit-report.json');browserReceiptSha256=Hash (Join-Path $directory 'browser-report.json');stages=$results;tolerancePt=0.001}
 Write-Output 'Both browser exports and independent native re-edits match native leaf, group, selection and XML frames.'
}finally{
 if($deck){Retry {$deck.Saved=-1;$deck.Close()}}
 if($app){if($null -ne $security){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;if($app.Presentations.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
