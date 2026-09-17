param([switch]$SourceOnly, [switch]$Create, [switch]$Cascade)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-line-metrics'))
if($Cascade){$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-font-cascade'))}
$baseline=Get-Content -LiteralPath (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
function Hash($path){return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Snapshot($document) {
 $document.Activate();$document.ActiveWindow.View.ReadingLayout=$false;$document.ActiveWindow.View.Type=3
 $document.Repaginate();$paragraphs=@()
 foreach($p in $document.Paragraphs){
  $lines=@();$last='';$r=$p.Range;$fontRuns=@();$lastFont=''
  for($i=[int]$r.Start;$i -lt [int]$r.End-1;$i++){
   $c=$document.Range($i,$i+1);$page=[int]$c.Information(3);$y=[double]$c.Information(6)
   $font=@{family=[string]$c.Font.Name;size=[double]$c.Font.Size};$fontKey="$($font.family)/$($font.size)"
   if($fontKey -cne $lastFont){$fontRuns+=@{from=$i-[int]$r.Start;to=$i-[int]$r.Start+1;family=$font.family;size=$font.size};$lastFont=$fontKey}
   else{$fontRuns[$fontRuns.Count-1].to=$i-[int]$r.Start+1}
   $key="$page/$y"
   if($key -cne $last){$lines+=@{from=$i-[int]$r.Start;to=$i-[int]$r.Start+1;page=$page;y=$y;text=[string]$c.Text};$last=$key}
   else{$line=$lines[$lines.Count-1];$line.to=$i-[int]$r.Start+1;$line.text+=[string]$c.Text}
  }
  $end=[int]$r.End-1;$caret=$document.Range($end,$end)
  $mark=$document.Range($end,$end+1)
  $paragraphs+=@{fontRuns=$fontRuns;markFont=@{family=[string]$mark.Font.Name;size=[double]$mark.Font.Size};rule=[int]$p.LineSpacingRule;spacing=[double]$p.LineSpacing;caret=@{page=[int]$caret.Information(3);y=[double]$caret.Information(6)};text=[string]$r.Text;keep=[int]$p.KeepTogether;widow=[int]$p.WidowControl;lines=$lines}
 }
 # ComputeStatistics can return the pre-repagination count on this build. Use the
 # observed physical page assignments; the PDF comparison independently checks it.
 return @{pages=[int](($paragraphs | ForEach-Object {$_.lines} | ForEach-Object {$_.page} | Measure-Object -Maximum).Maximum);paragraphs=$paragraphs}
}

$app=$null;$doc=$null;$links=$null;$owned=$false
try {
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$owned=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts
 $app.Visible=$false;$app.DisplayAlerts=0;$app.AutomationSecurity=3
 $links=$app.Options.UpdateLinksAtOpen;$app.Options.UpdateLinksAtOpen=$false
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
 $source=[string](Join-Path $root 'source.docx')
 $binding=Get-Content -LiteralPath (Join-Path $root 'source.json') -Raw|ConvertFrom-Json
 if((Hash $source) -ne $binding.sha256){throw 'Stale source'}
 $doc=$app.Documents.Open($source,$false,$false,$false)
 if($Create){
  $path=[string](Join-Path $root 'word-saved-source.docx');$doc.SaveAs2([ref]$path,[ref]12)
  $doc.Close([ref]0);$doc=$null
  Write-Output "Created owned Word-saved baseline: $path"
  return
 }
 $initial=Snapshot $doc
 $doc.ExportAsFixedFormat([string](Join-Path $root 'source-native.pdf'),17)
 $doc.Activate();$doc.ActiveWindow.View.ReadingLayout=$false;$doc.ActiveWindow.View.Type=3
 $cases=@(@{name='minimum';rule=3;spacing=18},@{name='exact';rule=4;spacing=18},@{name='auto';rule=1;spacing=18})
 if($Cascade){$cases+=@{name='typing';text='Mark text'}}
 foreach($case in $cases){
  if($case.text){$start=[int]$doc.Paragraphs.Item(4).Range.Start;$doc.Range($start,$start).Select();$app.Selection.TypeText([string]$case.text)}
  else{$p=$doc.Paragraphs.Item(2);$p.LineSpacingRule=$case.rule;$p.LineSpacing=$case.spacing}
  $copy=[string](Join-Path $root ('expected-'+$case.name+'.docx'));$doc.SaveAs2([ref]$copy,[ref]12)
 }
 $doc.Close([ref]0);$doc=$null
 $fonts=@();foreach($name in @('arial.ttf','arialbd.ttf','ariali.ttf','arialbi.ttf')){$fonts+=@{name=$name;sha256=Hash (Join-Path $env:WINDIR "Fonts/$name")}}
 $report=@{source=$initial;sourceHash=Hash $source;version=[string]$app.Version;build=[string]$app.Build;executableHash=$exeHash;fonts=$fonts;printer=[string]$app.ActivePrinter;scriptHash=Hash $PSCommandPath;cases=@();hashes=@()}
 if(!$SourceOnly){
  $browser=Get-Content -LiteralPath (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
  if(!$browser.passed -or $browser.sourceHash -ne (Hash $source)){throw 'Stale browser receipt'}
  for($i=0;$i -lt $initial.paragraphs.Count;$i++){
   $actualFont=$browser.measurements[$i];$expectedFont=$initial.paragraphs[$i].markFont
   if([Math]::Abs($actualFont.fontSize-$expectedFont.size) -gt 0.01 -or $actualFont.fontFamily -ne $expectedFont.family){throw "Browser paragraph mark font differs: $i"}
   foreach($run in $actualFont.runs){
    $p=$initial.paragraphs[$i];$start=$p.text.IndexOf([string]$run.text,[StringComparison]::Ordinal)
    if($start -lt 0){throw 'Browser font text not present in native paragraph'}
    $nativeRun=$p.fontRuns|Where-Object {$_.from -le $start -and $_.to -gt $start}|Select-Object -First 1
    if(!$nativeRun -or [Math]::Abs($run.size-$nativeRun.size) -gt 0.01 -or $run.family -ne $nativeRun.family){throw "Browser run font differs: $i $($run.text)"}
   }
  }
  $report.browser=$browser
  $wrapped=$initial.paragraphs[4].lines
  $nativeAdvance=[double]$wrapped[1].y-[double]$wrapped[0].y
  if(-not $browser.wrappedAdvances -or @($browser.wrappedAdvances).Count -ne $wrapped.Count-1){throw 'Browser wrapped line measurements are incomplete'}
  foreach($advance in $browser.wrappedAdvances){
   if([Math]::Abs([double]$advance*0.75-$nativeAdvance) -gt 0.01){throw 'Browser automatic line advance differs from native'}
  }
  if($Cascade){
   $nativeEmpty=[double]$initial.paragraphs[4].lines[0].y-[double]$initial.paragraphs[3].caret.y
   if([Math]::Abs([double]$browser.measurements[3].height*0.75-$nativeEmpty) -gt 0.01){throw 'Browser empty minimum paragraph advance differs from native'}
  }
 }
 foreach($case in $cases){
  $name=$case.name;$copy=[string](Join-Path $root ('expected-'+$name+'.docx'))
  $doc=$app.Documents.Open($copy,$false,$true,$false);$expected=Snapshot $doc
  $doc.ExportAsFixedFormat([string](Join-Path $root ('expected-'+$name+'.pdf')),17)
  $doc.Close([ref]0);$doc=$null
  if(!$SourceOnly){
   $export=[string](Join-Path $root ('browser-'+$name+'.docx'))
   if((Hash $export) -ne $browser.exports.$name){throw 'Stale browser export'}
   $doc=$app.Documents.Open($export,$false,$true,$false);$actual=Snapshot $doc
   $doc.ExportAsFixedFormat([string](Join-Path $root ('browser-'+$name+'.pdf')),17)
   $doc.Close([ref]0);$doc=$null
   if(($actual|ConvertTo-Json -Depth 12 -Compress) -cne ($expected|ConvertTo-Json -Depth 12 -Compress)){throw "Native spacing differs: $name"}
  }
  $report.cases+=@{name=$name;expected=$expected}
 }
 foreach($f in Get-ChildItem -LiteralPath $root -File){if($f.Name -match '\.(docx|pdf)$' -or $f.Name -eq 'browser-report.json'){$report.hashes+=@{path=$f.Name;sha256=Hash $f.FullName}}}
 $report|ConvertTo-Json -Depth 16|Set-Content -Encoding UTF8 (Join-Path $root $(if($SourceOnly){'source-native.json'}else{'native-report.json'}))
 Write-Output "Native line metrics: $($initial.pages) source pages; $($cases.Count) edited files measured."
}finally{
 try{if($doc){$doc.Close([ref]0)}}finally{if($app){try{if($owned){if($null -ne $links){$app.Options.UpdateLinksAtOpen=$links};$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;if($app.Documents.Count -eq 0){$app.Quit([ref]0)}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}
