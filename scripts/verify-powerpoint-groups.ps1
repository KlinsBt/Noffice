param([switch]$Compare)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root '.local/powerpoint-groups'
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
  if (!$Compare) {
    $deck=$app.Presentations.Add(0)
    $deck.PageSetup.SlideWidth=720;$deck.PageSetup.SlideHeight=405
    for ($i=1;$i -le 4;$i++) {
      $slide=$deck.Slides.Add($i,12)
      $a=$slide.Shapes.AddShape(1,50,60,40,30);$a.Name='Nested A';$a.Line.Visible=0;$a.Fill.ForeColor.RGB=255
      $b=$slide.Shapes.AddShape(9,110,80,30,40);$b.Name='Nested B';$b.Line.Visible=0;$b.Fill.ForeColor.RGB=65280
      $inner=$slide.Shapes.Range([object[]]@('Nested A','Nested B')).Group();$inner.Name='Inner'
      $inner.LockAspectRatio=0;$inner.Width=135;$inner.Height=90;$inner.Left=70;$inner.Top=80
      $c=$slide.Shapes.AddShape(1,230,100,40,50);$c.Name='Nested C';$c.Line.Visible=0;$c.Fill.ForeColor.RGB=16711680
      $outer=$slide.Shapes.Range([object[]]@('Inner','Nested C')).Group();$outer.Name='Outer'
      $outer.LockAspectRatio=0;$outer.Width=400;$outer.Height=135;$outer.Left=120;$outer.Top=140
      if ($i -eq 2) {$outer.Height=180}
      if ($i -eq 3) {$outer.Rotation=30}
      if ($i -eq 4) {$outer.Flip(0)}
    }
    foreach ($key in @('Author','Last Author','Company','Manager')) {try {$deck.BuiltInDocumentProperties.Item($key).Value=''} catch {}}
    $source=Join-Path $directory 'source.pptx';$deck.SaveAs($source,24);$deck.Close();$deck=$null
    $deck=$app.Presentations.Open($source,-1,0,0)
    $grouped=Snapshot $deck
    $deck.SaveAs((Join-Path $directory 'source.pdf'),32)
    Ungroup $deck
    WriteJson 'source-report.json' ([ordered]@{application='PowerPoint';version=$native.version;sourceSha256=(Get-FileHash $source).Hash.ToLowerInvariant();scriptSha256=(Get-FileHash $PSCommandPath).Hash.ToLowerInvariant();width=720;height=405;grouped=$grouped;effective=(Snapshot $deck)})
    Write-Output 'Authored nested scale/translation, rotation and flip fixtures; measured grouped and ungrouped native bounds.'
  } else {
    $source=Join-Path $root 'tests/fixtures/powerpoint-groups.pptx'
    $export=Join-Path $directory 'edited.pptx'
    $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json') | ConvertFrom-Json
    if (!$browser.passed -or (Get-FileHash $source).Hash.ToLowerInvariant() -ne $browser.sourceSha256 -or (Get-FileHash $export).Hash.ToLowerInvariant() -ne $browser.exportSha256) {throw 'Browser evidence does not match the tested bytes.'}
    $deck=$app.Presentations.Open($source,0,0,0)
    # COM child coordinates are recorded by the source probe before choosing these edits.
    $child=FindLeaf $deck.Slides.Item(1).Shapes 2
    if (!$child -or $child.Name -ne 'Nested A') {throw 'The authored group leaf is missing.'}
    $child.Left=$child.Left+12;$child.Top=$child.Top+9
    $child.Width=$child.Width+12;$child.Height=$child.Height+9
    $expectedPath=Join-Path $directory 'native-expected.pptx'
    $deck.SaveAs($expectedPath,24);$deck.Close();$deck=$null
    $deck=$app.Presentations.Open($expectedPath,-1,0,0)
    $expected=Snapshot $deck;$deck.SaveAs((Join-Path $directory 'native-expected.pdf'),32)
    $deck.Close();$deck=$null
    $deck=$app.Presentations.Open($export,-1,0,0)
    $actual=Snapshot $deck;$deck.SaveAs((Join-Path $directory 'browser-export.pdf'),32)
    WriteJson 'expected-snapshot.json' $expected;WriteJson 'actual-snapshot.json' $actual
    if (($expected | ConvertTo-Json -Depth 20 -Compress) -cne ($actual | ConvertTo-Json -Depth 20 -Compress)) {throw 'Native group-child edits differ from browser export.'}
    WriteJson 'native-report.json' ([ordered]@{passed=$true;sourceSha256=$browser.sourceSha256;exportSha256=$browser.exportSha256;expectedSha256=(Get-FileHash $expectedPath).Hash.ToLowerInvariant();scriptSha256=(Get-FileHash $PSCommandPath).Hash.ToLowerInvariant();application='PowerPoint';version=$native.version;slides=$actual})
    Write-Output 'Native group-child edits match the actual browser export on every leaf of all four slides.'
  }
} finally {
  if ($deck) {$deck.Saved=-1;$deck.Close()}
  if ($app) {
    if ($null -ne $security) {$app.DisplayAlerts=$alerts;$app.AutomationSecurity=$security;if ($app.Presentations.Count -eq 0) {$app.Quit()}}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  }
}
