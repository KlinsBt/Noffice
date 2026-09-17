param([switch]$Paragraphs)
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$root=Join-Path $workspace '.local/word-side-stories'
function Hash($path){(Get-FileHash -LiteralPath $path).Hash.ToLowerInvariant()}
$scriptHash=Hash $PSCommandPath
function Snapshot($document){
 $document.Activate();$document.ActiveWindow.View.ReadingLayout=$false;$document.ActiveWindow.View.Type=3
 $document.Repaginate();$sections=@()
 foreach($section in $document.Sections){
  $stories=@()
  foreach($kind in @('Headers','Footers')){foreach($slot in @(1,2,3)){
   $story=$section.$kind.Item($slot);$paragraphs=@();$characters=@()
   if($story.Exists){
    foreach($p in $story.Range.Paragraphs){
     $r=$p.Range;$f=$r.ParagraphFormat
     $paragraphs+=@{text=[string]$r.Text;lineRule=[int]$f.LineSpacingRule;line=[double]$f.LineSpacing;before=[double]$f.SpaceBefore;after=[double]$f.SpaceAfter}
    }
    foreach($c in $story.Range.Characters){$characters+=@{text=[string]$c.Text;font=[string]$c.Font.Name;size=[double]$c.Font.Size;bold=[int]$c.Font.Bold;italic=[int]$c.Font.Italic;underline=[int]$c.Font.Underline}}
   }
   $stories+=@{kind=$kind;slot=$slot;exists=[bool]$story.Exists;linked=[bool]$story.LinkToPrevious;paragraphs=$paragraphs;characters=$characters}
  }}
  $p=$section.PageSetup
  $sections+=@{stories=$stories;width=[double]$p.PageWidth;height=[double]$p.PageHeight;top=[double]$p.TopMargin;bottom=[double]$p.BottomMargin;header=[double]$p.HeaderDistance;footer=[double]$p.FooterDistance;first=[bool]$p.DifferentFirstPageHeaderFooter;even=[bool]$p.OddAndEvenPagesHeaderFooter}
 }
 return @{pages=[int]$document.ComputeStatistics(2);body=[string]$document.Content.Text;sections=$sections}
}
$app=$null;$doc=$null;$owned=$false;$rows=@();$hashes=@();$returns=@()
try{
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$owned=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen
 $app.Visible=$false;$app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $baseline=(Get-Content (Join-Path $workspace 'docs/parity/baseline.json') -Raw|ConvertFrom-Json).applications|Where-Object executable -eq 'WINWORD.EXE'
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exe -ne $baseline.sha256){throw 'Pinned Word build changed'}
 $cases=@('default','first-even','long-header','long-footer','varying-slots')|ForEach-Object {@{name=$_;source=$_;folder="browser/$_";kind='Headers';grow=$false}}
 $cases+=@{name='grow-header';source='default';folder='flow/header';kind='Headers';grow=$true}
 $cases+=@{name='grow-footer';source='default';folder='flow/footer';kind='Footers';grow=$true}
 $cases+=@{name='all-slots';source='first-even';folder='slots/all';kind='Headers';allSlots=$true}
 if($Paragraphs){$cases=@(
  @{name='paragraph-header';source='default';folder='paragraphs/header';kind='Headers';paragraphs=$true},
  @{name='paragraph-footer';source='default';folder='paragraphs/footer';kind='Footers';paragraphs=$true}
 )}
 foreach($case in $cases){
  $name=$case.name;$folder=$case.folder;$caseRoot=Join-Path $root $folder
  $browser=Get-Content (Join-Path $caseRoot 'browser-report.json') -Raw|ConvertFrom-Json
  $source=Join-Path $workspace ("tests/fixtures/word-side-stories/"+$case.source+'.docx')
  if((Hash $source) -ne $browser.hashes.source){throw 'Native source binding changed'}
  $stages=if($Paragraphs){@('undo','edited','reloaded','joined')}else{@('undo','edited','reloaded')}
  foreach($stage in $stages){
   $actualPath=Join-Path $caseRoot "$stage.docx"
   if((Hash $actualPath) -ne $browser.hashes.$stage){throw 'Actual export binding changed'}
   $expectedPath=[string](Join-Path $caseRoot "expected-$stage.docx")
   $comparisonSource=if($case.paragraphs -and $stage -eq 'joined'){Join-Path $caseRoot 'expected-reloaded.docx'}else{$source}
   Copy-Item -LiteralPath $comparisonSource -Destination $expectedPath -Force
   $doc=$app.Documents.Open($expectedPath,$false,$false,$false)
   $doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
   if($stage -ne 'undo'){
    $kind=$case.kind;$range=$doc.Sections.Item(1).$kind.Item(1).Range.Duplicate
    if($case.paragraphs){
     if($stage -eq 'joined'){
      # Reopen the independently grown native document and perform the same
      # paragraph-mark joins/deletions as the browser. Replacing the entire
      # story with final text would erase native run/session boundaries.
      $first=$doc.Sections.Item(1).$kind.Item(1).Range.Paragraphs.Item(1).Range
      $range.SetRange([int]$first.End-1,[int]$first.End);$range.Text=''
      $last=$doc.Sections.Item(1).$kind.Item(1).Range.Paragraphs.Last.Range
      $range.SetRange([int]$last.Start,[int]$last.End-1);$range.Text=''
      $last=$doc.Sections.Item(1).$kind.Item(1).Range.Paragraphs.Last.Range
      $range.SetRange([int]$last.Start-1,[int]$last.Start);$range.Text=''
     }else{
      $end=[int]$range.End-1;$range.SetRange($end,$end)
      $range.Text=([string][char]13+'extra1'+[char]13+'extra2'+[char]13+'extra3')
     }
    }elseif($case.allSlots){
     foreach($storyKind in @('Headers','Footers')){
      $label=if($storyKind -eq 'Headers'){'header'}else{'footer'}
      foreach($slot in @(@{name='default';index=1},@{name='first';index=2},@{name='even';index=3})){
       $range=$doc.Sections.Item(1).$storyKind.Item($slot.index).Range.Duplicate
       $range.SetRange([int]$range.Start,[int]$range.End-1)
       $range.Text=('Changed '+$slot.name+' '+$label)
      }
     }
    }elseif($case.grow){
     $end=[int]$range.End-1;$range.SetRange($end,$end)
     # Insert plain text at the collapsed story-local range. The installed
     # build rejects InsertAfter here even in print view; Text is the separate
     # documented range replacement operation and leaves the final mark out.
     $range.Text=([string][char]11+'extra1'+[char]11+'extra2'+[char]11+'extra3')
    }else{
     $text=[string]$range.Text;$first=($text -split '[\r\x0b]')[0]
     $start=[int]$range.Start;$range.SetRange($start,$start+$first.Length)
     $range.Text='Updated header';$range.SetRange($start,$start+14);$range.Font.Bold=-1
    }
   }
   $doc.SaveAs2([ref]$expectedPath,[ref]12);$doc.Close([ref]0);$doc=$null
   $doc=$app.Documents.Open($expectedPath,$false,$true,$false)
   $wanted=Snapshot $doc;$expectedPdf=[string](Join-Path $caseRoot "expected-$stage.pdf")
   $doc.ExportAsFixedFormat($expectedPdf,17);$doc.Close([ref]0);$doc=$null
   $doc=$app.Documents.Open($actualPath,$false,$true,$false)
   $actual=Snapshot $doc;$actualPdf=[string](Join-Path $caseRoot "native-$stage.pdf")
   $doc.ExportAsFixedFormat($actualPdf,17);$doc.Close([ref]0);$doc=$null
   $rows+=@{name=$name;folder=$folder;stage=$stage;expected=$wanted;actual=$actual}
   foreach($path in @($actualPath,$expectedPath,$expectedPdf,$actualPdf)){
    $hashes+=@{path=$path.Substring($root.Length+1).Replace('\','/');sha256=Hash $path}
   }
   if($case.paragraphs -and $stage -in @('edited','joined')){
    $download=Join-Path $caseRoot "$stage.pdf"
    if((Hash $download) -ne $browser.hashes."$stage.pdf"){throw 'Paragraph PDF binding changed'}
    $hashes+=@{path="$folder/$stage.pdf";sha256=Hash $download}
   }
   Write-Output "Captured $name $stage"
  }
  if((Hash $source) -ne $browser.hashes.source){throw 'Source changed during native verification'}
  $hashes+=@{path="$folder/browser-report.json";sha256=Hash (Join-Path $caseRoot 'browser-report.json')}
  if($case.grow -or $case.allSlots){
   if((Hash (Join-Path $caseRoot 'download.pdf')) -ne $browser.hashes.pdf){throw 'Growth PDF binding changed'}
   $hashes+=@{path="$folder/download.pdf";sha256=Hash (Join-Path $caseRoot 'download.pdf')}
  }
  if($case.paragraphs){
   $returnPath=[string](Join-Path $caseRoot 'native-return.docx')
   Copy-Item -LiteralPath (Join-Path $caseRoot 'joined.docx') -Destination $returnPath -Force
   $doc=$app.Documents.Open($returnPath,$false,$false,$false)
   $doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
   $range=$doc.Sections.Item(1).$kind.Item(1).Range.Duplicate
   $end=[int]$range.End-1;$range.SetRange($end,$end);$range.Text=' native'
   $doc.Save();$doc.Close([ref]0);$doc=$null
   $doc=$app.Documents.Open($returnPath,$false,$true,$false)
   $state=Snapshot $doc;$returnPdf=[string](Join-Path $caseRoot 'native-return.pdf')
   $doc.ExportAsFixedFormat($returnPdf,17);$doc.Close([ref]0);$doc=$null
   $returns+=@{name=$name;folder=$folder;native=$state}
   foreach($path in @($returnPath,$returnPdf)){$hashes+=@{path=$path.Substring($root.Length+1).Replace('\','/');sha256=Hash $path}}
  }
 }
 if((Hash $PSCommandPath) -ne $scriptHash){throw 'Native verifier changed during capture'}
 $report=if($Paragraphs){'story-paragraph-native-report.json'}else{'story-edit-native-report.json'}
 @{captured=$true;scriptHash=$scriptHash;executableHash=$exe;rows=$rows;hashes=$hashes;returns=$returns}|ConvertTo-Json -Depth 20|Set-Content (Join-Path $root $report) -Encoding UTF8
}finally{
 if($doc){$doc.Close([ref]0)}
 if($app){if($owned){$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($app.Documents.Count -eq 0){$app.Quit()}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
