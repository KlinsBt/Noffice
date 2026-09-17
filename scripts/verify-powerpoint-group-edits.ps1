param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/powerpoint-group-edits'
[void][IO.Directory]::CreateDirectory($directory)
$baseline=Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json') | ConvertFrom-Json
$native=$baseline.applications | Where-Object executable -eq 'POWERPNT.EXE'
if ((Get-FileHash -LiteralPath $native.path).Hash.ToLowerInvariant() -ne $native.sha256) {throw 'Native PowerPoint baseline changed.'}
function Leaves($shapes) {
  $result=@()
  foreach ($shape in $shapes) {
    if ($shape.Type -eq 6) {$result+=Leaves $shape.GroupItems}
    else {$result += [ordered]@{name=[string]$shape.Name;id=[int]$shape.Id;x=[double]$shape.Left;y=[double]$shape.Top;w=[double]$shape.Width;h=[double]$shape.Height;rotation=[double]$shape.Rotation}}
  }
  return $result
}
function Snapshot($deck) {
  $result=@()
  foreach ($slide in $deck.Slides) {$result+= ,@(Leaves $slide.Shapes)}
  return ,$result
}
function FindLeaf($shapes,$id) {
  foreach ($shape in $shapes) {
    if ($shape.Type -eq 6) {$found=FindLeaf $shape.GroupItems $id;if ($found) {return $found}}
    elseif ($shape.Id -eq $id) {return $shape}
  }
}
function Ungroup($deck) {
  foreach ($slide in $deck.Slides) {
    $remaining=$true
    while ($remaining) {
      $remaining=$false
      for ($i=$slide.Shapes.Count;$i -ge 1;$i--) {
        if ($slide.Shapes.Item($i).Type -eq 6) {[void]$slide.Shapes.Item($i).Ungroup();$remaining=$true}
      }
    }
  }
}
function WriteJson($name,$value) {[IO.File]::WriteAllText((Join-Path $directory $name),($value | ConvertTo-Json -Depth 20),(New-Object Text.UTF8Encoding($false)))}
$app=$null;$deck=$null
try {
  $app=New-Object -ComObject PowerPoint.Application
  if ($app.Presentations.Count -ne 0) {throw 'PowerPoint has open presentations; this oracle needs an idle application.'}
  $alerts=$app.DisplayAlerts;$security=$app.AutomationSecurity
  $app.DisplayAlerts=1;$app.AutomationSecurity=3
  $runningBinary=Join-Path ([string]$app.Path) 'POWERPNT.EXE'
  if((Get-FileHash -LiteralPath $runningBinary).Hash.ToLowerInvariant() -ne $native.sha256){throw 'The running PowerPoint application differs from the pinned baseline.'}
  $source=Join-Path $root 'tests/fixtures/powerpoint-groups.pptx'
  $deck=$app.Presentations.Open($source,0,0,0)
  $initial=Snapshot $deck
  foreach($index in @(3,4)) {
    $leaf=FindLeaf $deck.Slides.Item($index).Shapes 2
    if(!$leaf -or $leaf.Name -ne 'Nested A'){throw 'Authored group child is missing.'}
    $leaf.LockAspectRatio=0
    $leaf.Left=$leaf.Left+12;$leaf.Top=$leaf.Top+9
    $leaf.Width=$leaf.Width+12;$leaf.Height=$leaf.Height+9
  }
  $expectedPath=Join-Path $directory 'native-expected.pptx'
  $deck.SaveAs($expectedPath,24);$deck.Close();$deck=$null
  $deck=$app.Presentations.Open($expectedPath,-1,0,0)
  $expected=Snapshot $deck
  $deck.SaveAs((Join-Path $directory 'native-expected.pdf'),32)
  Ungroup $deck
  $effective=Snapshot $deck
  $deck.Close();$deck=$null
  if(!$Compare){
    WriteJson 'reference.json' ([ordered]@{sourceSha256=(Get-FileHash $source).Hash.ToLowerInvariant();scriptSha256=(Get-FileHash $PSCommandPath).Hash.ToLowerInvariant();executableSha256=$native.sha256;initial=$initial;expected=$expected;effective=$effective})
    Write-Output 'Native rotated/reflected child edits captured: slides 3 and 4, child 2, move and resize.'
  }else{
    $export=Join-Path $directory 'edited.pptx'
    $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json') | ConvertFrom-Json
    if(!$browser.passed -or (Get-FileHash $source).Hash.ToLowerInvariant() -ne $browser.sourceSha256 -or (Get-FileHash $export).Hash.ToLowerInvariant() -ne $browser.exportSha256){throw 'Stale browser evidence.'}
    $deck=$app.Presentations.Open($export,-1,0,0)
    $actual=Snapshot $deck;$deck.SaveAs((Join-Path $directory 'browser-export.pdf'),32)
    Ungroup $deck;$actualEffective=Snapshot $deck
    WriteJson 'comparison.json' @{expected=$expected;actual=$actual;expectedEffective=$effective;actualEffective=$actualEffective}
    foreach($pair in @(@{expected=$expected;actual=$actual},@{expected=$effective;actual=$actualEffective})){
      for($i=0;$i -lt 4;$i++){
        if($pair.expected[$i].Count -ne $pair.actual[$i].Count){throw 'Leaf count changed.'}
        for($j=0;$j -lt $pair.expected[$i].Count;$j++){
          $p=$pair.expected[$i][$j];$q=$pair.actual[$i][$j]
          if($p.id -ne $q.id -or $p.name -cne $q.name){throw 'Leaf identity changed.'}
          foreach($key in @('x','y','w','h','rotation')){if([Math]::Abs($p.$key-$q.$key) -gt 0.001){throw "Native frame differs: slide $i leaf $j $key"}}
        }
      }
    }
    WriteJson 'native-report.json' ([ordered]@{passed=$true;sourceSha256=$browser.sourceSha256;exportSha256=$browser.exportSha256;expectedSha256=(Get-FileHash $expectedPath).Hash.ToLowerInvariant();scriptSha256=(Get-FileHash $PSCommandPath).Hash.ToLowerInvariant();executableSha256=$native.sha256;slides=$actual;effective=$actualEffective;tolerancePt=0.001})
    Write-Output 'Rotated/reflected group child exports match all 12 grouped and 12 effective native frames.'
  }
} finally {
  if ($deck) {$deck.Saved=-1;$deck.Close()}
  if ($app) {
    if ($null -ne $security) {$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if ($app.Presentations.Count -eq 0) {$app.Quit()}}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  }
}
