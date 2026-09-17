param([switch]$Compare)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory = Join-Path $root '.local/powerpoint-inheritance'
[void][IO.Directory]::CreateDirectory($directory)
$baseline = Get-Content -Raw (Join-Path $root 'docs/parity/baseline.json') | ConvertFrom-Json
$native = $baseline.applications | Where-Object executable -eq 'POWERPNT.EXE'
if ((Get-FileHash -LiteralPath $native.path).Hash.ToLowerInvariant() -ne $native.sha256) { throw 'Native PowerPoint baseline changed.' }
function Snapshot($deck) {
  $slides = @()
  foreach ($slide in $deck.Slides) {
    $shapes = @()
    foreach ($shape in $slide.Shapes) {
      $record = [ordered]@{id=[int]$shape.Id; name=[string]$shape.Name; type=[int]$shape.Type; x=[double]$shape.Left; y=[double]$shape.Top; w=[double]$shape.Width; h=[double]$shape.Height}
      if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
        $range = $shape.TextFrame.TextRange
        $characters = @()
        for ($i=1; $i -le $range.Length; $i++) {
          $char = $range.Characters($i,1)
          $characters += [ordered]@{text=[string]$char.Text; size=[double]$char.Font.Size; font=[string]$char.Font.Name; bold=($char.Font.Bold -eq -1); italic=($char.Font.Italic -eq -1); underline=($char.Font.Underline -eq -1); rgb=[int]$char.Font.Color.RGB}
        }
        $record.text = [string]$range.Text
        $record.characters = $characters
      }
      $shapes += $record
    }
    $slides += [ordered]@{design=[string]$slide.Design.Name; layout=[string]$slide.CustomLayout.Name; shapes=$shapes}
  }
  return ,$slides
}
$app=$null; $deck=$null
try {
  $app=New-Object -ComObject PowerPoint.Application
  # PowerPoint can return an existing singleton. Never change preferences in a user session.
  if ($app.Presentations.Count -ne 0) { throw 'PowerPoint has open presentations; native oracle requires an idle application.' }
  $alerts=$app.DisplayAlerts; $security=$app.AutomationSecurity
  $app.DisplayAlerts=1; $app.AutomationSecurity=3
  if (!$Compare) {
    $deck=$app.Presentations.Add(0)
    $deck.PageSetup.SlideWidth=720; $deck.PageSetup.SlideHeight=405
    $first=$deck.Designs.Item(1); $first.Name='Noffice Master One'
    $second=$deck.Designs.Clone($first,2); $second.Name='Noffice Master Two'
    $index=0
    foreach ($design in @($first,$second)) {
      $index++
      $masterFont=$design.SlideMaster.TextStyles.Item(1).Levels.Item(1).Font
      $masterFont.Name='Arial'; $masterFont.Size=30 + $index * 2
      $masterFont.Bold=0; $masterFont.Italic=0
      $layout=$design.SlideMaster.CustomLayouts.Item(1)
      $layout.Name="Noffice Layout $index"
      $title=$layout.Shapes.Title
      $title.Left=48 + $index * 12; $title.Top=30; $title.Width=560; $title.Height=85
      $title.TextFrame.TextRange.Font.Size=36 + $index * 2
      $title.TextFrame.TextRange.Font.Italic=-1
      $slide=$deck.Slides.AddSlide($index,$layout)
      $slide.Shapes.Title.Name="Inherited title $index"
      $slide.Shapes.Title.TextFrame.TextRange.Text="Master $index title"
      $sub=$slide.Shapes.Placeholders.Item(2)
      $sub.Name="Rich subtitle $index"
      $sub.TextFrame.TextRange.Text='Plain BOLD tail'
      $sub.TextFrame.TextRange.Font.Name='Arial'; $sub.TextFrame.TextRange.Font.Size=20
      $sub.TextFrame.TextRange.Font.Bold=0
      $sub.TextFrame.TextRange.Characters(7,4).Font.Bold=-1
      $sub.TextFrame.TextRange.Characters(7,4).Font.Size=28
      $sub.TextFrame.TextRange.Characters(1,6).Font.Underline=-1
      $sub.TextFrame.TextRange.Characters(11,5).Font.Italic=-1
      $sub.TextFrame.TextRange.Characters(11,5).Font.Color.RGB=3351057
      $a=$slide.Shapes.AddShape(1,80,280,40,30); $a.Name='Nested A'
      $b=$slide.Shapes.AddShape(1,140,280,40,30); $b.Name='Nested B'
      $inner=$slide.Shapes.Range([object[]]@('Nested A','Nested B')).Group(); $inner.Name='Inner group'
      $c=$slide.Shapes.AddShape(1,210,280,40,30); $c.Name='Nested C'
      $outer=$slide.Shapes.Range([object[]]@('Inner group','Nested C')).Group(); $outer.Name='Outer group'
    }
    foreach ($key in @('Author','Last Author','Company','Manager')) { try {$deck.BuiltInDocumentProperties.Item($key).Value=''} catch {} }
    $path=Join-Path $directory 'source.pptx'; $deck.SaveAs($path,24); $deck.Close(); $deck=$null
    $deck=$app.Presentations.Open($path,-1,0,0)
    $report=[ordered]@{application='PowerPoint';version=$native.version;sourceSha256=(Get-FileHash $path).Hash.ToLowerInvariant();width=720;height=405;slides=(Snapshot $deck)}
    $deck.SaveAs((Join-Path $directory 'source.pdf'),32)
    [IO.File]::WriteAllText((Join-Path $directory 'source-report.json'),($report | ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
    Write-Output 'Authored and reopened two native masters/layouts, rich placeholders and nested groups.'
  } else {
    $source=Join-Path $root 'tests/fixtures/powerpoint-inheritance.pptx'
    $export=Join-Path $directory 'edited.pptx'
    $browser=Get-Content -Raw (Join-Path $directory 'browser-report.json') | ConvertFrom-Json
    if (!$browser.passed -or (Get-FileHash $source).Hash.ToLowerInvariant() -ne $browser.sourceSha256 -or (Get-FileHash $export).Hash.ToLowerInvariant() -ne $browser.exportSha256) { throw 'Browser evidence does not match source/export bytes.' }
    $deck=$app.Presentations.Open($source,0,0,0)
    $slide=$deck.Slides.Item(1)
    $slide.Shapes.Title.Left=72
    $slide.Shapes.Title.TextFrame.TextRange.Text='Master 1 title edited'
    $slide.Shapes.Item('Rich subtitle 1').TextFrame.TextRange.Characters(7,4).Text='BRAVE'
    $expectedPath=Join-Path $directory 'native-expected.pptx'
    $deck.SaveAs($expectedPath,24); $deck.Close(); $deck=$null
    $deck=$app.Presentations.Open($expectedPath,-1,0,0)
    $expected=Snapshot $deck
    $deck.SaveAs((Join-Path $directory 'native-expected.pdf'),32)
    $deck.Close(); $deck=$null
    $deck=$app.Presentations.Open($export,-1,0,0)
    $actual=Snapshot $deck
    $deck.SaveAs((Join-Path $directory 'browser-export.pdf'),32)
    $expectedJson=$expected | ConvertTo-Json -Depth 12 -Compress
    $actualJson=$actual | ConvertTo-Json -Depth 12 -Compress
    [IO.File]::WriteAllText((Join-Path $directory 'expected-snapshot.json'),$expectedJson,(New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $directory 'actual-snapshot.json'),$actualJson,(New-Object Text.UTF8Encoding($false)))
    if ($actualJson -cne $expectedJson) { throw 'Native source edit and browser export differ; inspect saved snapshots.' }
    $report=[ordered]@{passed=$true;application='PowerPoint';version=$native.version;sourceSha256=$browser.sourceSha256;exportSha256=$browser.exportSha256;expectedSha256=(Get-FileHash $expectedPath).Hash.ToLowerInvariant();scriptSha256=(Get-FileHash $PSCommandPath).Hash.ToLowerInvariant();slides=$actual}
    [IO.File]::WriteAllText((Join-Path $directory 'native-report.json'),($report | ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
    Write-Output 'Both slides match native shape identities, design/layout, geometry and every text character style after independent edits.'
  }
} finally {
  if ($deck) { $deck.Close() }
  if ($app) {
    if ($null -ne $security) { $app.DisplayAlerts=$alerts; $app.AutomationSecurity=$security; if ($app.Presentations.Count -eq 0) { $app.Quit() } }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
  }
}
