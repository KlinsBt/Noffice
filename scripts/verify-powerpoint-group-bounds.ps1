param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/powerpoint-group-bounds'
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
$app=$null;$deck=$null
try{
 $baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json')|ConvertFrom-Json
 $native=$baseline.applications|Where-Object executable -eq 'POWERPNT.EXE'
 $app=New-Object -ComObject PowerPoint.Application
 if($app.Presentations.Count -ne 0){throw 'The native test requires an idle owned PowerPoint instance.'}
 $processes=@(Get-CimInstance Win32_Process -Filter "name = 'POWERPNT.EXE'")
 if($processes.Count -ne 1){throw 'Cannot uniquely identify the owned PowerPoint process.'}
 $script:officeProcess=[uint32]$processes[0].ProcessId
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$app.AutomationSecurity=3;$app.DisplayAlerts=1
 if((Hash (Join-Path ([string]$app.Path) 'POWERPNT.EXE')) -ne $native.sha256){throw 'The native baseline changed.'}
 $source=Join-Path $root 'tests/fixtures/powerpoint-groups.pptx'
 $deck=$app.Presentations.Open($source,0,0,0)
 if(!$deck){throw 'PowerPoint did not return the authored test presentation.'}
 $window=$deck.NewWindow();Retry {$window.WindowState=2}
 $initial=Snapshot $deck;$initialSelection=SelectBounds $deck;$steps=@()
 foreach($index in 1..4){
  $leaf=FindLeaf $deck.Slides.Item($index).Shapes 2;$leaf.LockAspectRatio=0
  foreach($property in @('Left','Top','Width','Height')){
   $delta=if($property -in @('Left','Width')){12}else{9}
   $leaf.$property=[double]$leaf.$property+$delta
   $steps+=,[ordered]@{slide=$index;property=$property;shapes=@(Shapes $deck.Slides.Item($index).Shapes @())}
  }
 }
 $edited=Snapshot $deck;$editedSelection=SelectBounds $deck
 $expected=Join-Path $directory 'native-expected.pptx';$deck.SaveAs($expected,24);$deck.Close();$deck=$null
 $deck=$app.Presentations.Open($expected,-1,0,0);$window=$deck.NewWindow();Retry {$window.WindowState=2}
 $reopened=Snapshot $deck;$selection=SelectBounds $deck
 $deck.SaveAs((Join-Path $directory 'native-expected.pdf'),32)
 Ungroup $deck;$effective=Snapshot $deck;$deck.Saved=-1;$deck.Close();$deck=$null
 WriteJson 'reference.json' ([ordered]@{sourceSha256=Hash $source;scriptSha256=Hash $PSCommandPath;executableSha256=$native.sha256;initial=$initial;initialSelection=$initialSelection;steps=$steps;edited=$edited;editedSelection=$editedSelection;reopened=$reopened;selection=$selection;effective=$effective;transforms=(Transforms $expected);expectedSha256=Hash $expected})
 if($Compare){
  $export=Join-Path $directory 'edited.pptx'
  $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json')|ConvertFrom-Json
  if(!$browser.passed -or (Hash $source) -ne $browser.sourceSha256 -or (Hash $export) -ne $browser.exportSha256){throw 'Stale browser evidence.'}
  $deck=$app.Presentations.Open($export,-1,0,0);$window=$deck.NewWindow();Retry {$window.WindowState=2}
  $actual=Snapshot $deck;$actualSelection=SelectBounds $deck
  $deck.SaveAs((Join-Path $directory 'browser-export.pdf'),32)
  Ungroup $deck;$actualEffective=Snapshot $deck;$deck.Saved=-1;$deck.Close();$deck=$null
  $maximum=0
  foreach($pair in @(@{expected=$reopened;actual=$actual},@{expected=$effective;actual=$actualEffective})){
   for($i=0;$i -lt 4;$i++){
    if($pair.expected[$i].Count -ne $pair.actual[$i].Count){throw 'Shape count changed.'}
    for($j=0;$j -lt $pair.expected[$i].Count;$j++){
     $p=$pair.expected[$i][$j];$q=$pair.actual[$i][$j]
     if($p.id -ne $q.id -or $p.name -cne $q.name -or ($p.parents|ConvertTo-Json -Compress) -cne ($q.parents|ConvertTo-Json -Compress)){throw 'Native shape identity changed.'}
     foreach($key in @('x','y','w','h','rotation')){
      $difference=[Math]::Abs($p.$key-$q.$key);$maximum=[Math]::Max($maximum,$difference)
      if($difference -gt 0.001){throw "Native frame differs: slide $i shape $j $key ($difference)"}
     }
    }
   }
  }
  for($i=0;$i -lt 4;$i++){
   if($actualSelection[$i].count -ne 1){throw 'The native group selection is not singular.'}
   foreach($key in @('x','y','w','h','rotation')){if([Math]::Abs($selection[$i].$key-$actualSelection[$i].$key) -gt 0.001){throw "Selection bounds differ: slide $i $key"}}
  }
  $expectedTransforms=Transforms $expected;$actualTransforms=Transforms $export
  for($i=0;$i -lt 4;$i++){
   if($actualTransforms[$i].Count -ne $expectedTransforms[$i].Count){throw 'Nested XML group count changed.'}
   for($j=0;$j -lt $expectedTransforms[$i].Count;$j++){
    $p=$expectedTransforms[$i][$j];$q=$actualTransforms[$i][$j]
    if($p.id -ne $q.id){throw 'Nested XML group identity changed.'}
    foreach($key in $p.Keys){if($key -ne 'id' -and [Math]::Abs($p[$key]-$q[$key]) -gt 12.7){throw "Nested group transform differs: slide $i group $j $key"}}
   }
  }
  WriteJson 'native-report.json' ([ordered]@{passed=$true;sourceSha256=$browser.sourceSha256;exportSha256=$browser.exportSha256;expectedSha256=Hash $expected;scriptSha256=Hash $PSCommandPath;executableSha256=$native.sha256;actual=$actual;effective=$actualEffective;selection=$actualSelection;transforms=$actualTransforms;tolerancePt=0.001;maximumFrameDifferencePt=$maximum})
  Write-Output 'All native group selections, grouped/effective leaves and eight nested XML group frames match.'
 }
 Write-Output 'Recorded native group and selection bounds before/after all 16 child edits and after reopening.'
}catch{Write-Output $_.ToString();Write-Output $_.ScriptStackTrace;throw}finally{
 if($deck){Retry {$deck.Saved=-1;$deck.Close()}}
 if($app){if($null -ne $security){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;if($app.Presentations.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
