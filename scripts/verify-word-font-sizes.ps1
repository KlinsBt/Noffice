param([switch]$Compare)
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../.local/word-font-sizes'))
[void][IO.Directory]::CreateDirectory($root)
function Hash($path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Sizes($document) {
 $values = @()
 foreach ($p in $document.Paragraphs) { $values += [double]$document.Range([int]$p.Range.Start, [int]$p.Range.Start+1).Font.Size }
 return ,$values
}
$app=$null; $doc=$null; $links=$null
try {
 $app=New-Object -ComObject Word.Application
 $app.Visible=$false; $app.DisplayAlerts=0; $app.AutomationSecurity=3
 $links=$app.Options.UpdateLinksAtOpen; $app.Options.UpdateLinksAtOpen=$false
 $baseline=Get-Content (Join-Path $PSScriptRoot '../docs/parity/baseline.json') -Raw|ConvertFrom-Json
 $exeHash=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE')
 if($exeHash -ne ($baseline.applications|Where-Object executable -eq 'WINWORD.EXE').sha256){throw 'Word baseline changed'}
 $sizes=@(1,1.5,6,7.5,10,160,160.5,200,1638)
 $source=[string](Join-Path $root 'source.docx')
 if(!$Compare){
  $doc=$app.Documents.Add(); $doc.Content.Text=(@('a')*$sizes.Count)-join "`r"
  for($i=0;$i -lt $sizes.Count;$i++){$r=$doc.Paragraphs.Item($i+1).Range;$r.Font.Name='Arial';$r.Font.Size=[single]$sizes[$i]}
  $doc.SaveAs2([ref]$source,[ref]12);$doc.Close([ref]0);$doc=$null
 }
 $doc=$app.Documents.Open($source,$false,$true,$false);$actual=Sizes $doc
 if(($actual -join ',') -ne ($sizes -join ',')){throw 'Native source size mismatch'}
 $doc.Close([ref]0);$doc=$null
 $report=@{sourceHash=Hash $source;executableHash=$exeHash;version=[string]$app.Version;build=[string]$app.Build;sizes=$actual;scriptHash=Hash $PSCommandPath;cases=@();steps=@()}
 $doc=$app.Documents.Add();$doc.Content.Text='a'
 foreach($size in @(1,1.5,6,7.5,8,9,10,13,28,36,48,71,72,73,79,80,96,160,160.5,200,1638)){
  $font=$doc.Range(0,1).Font;$font.Size=[single]$size;$font.Grow();$grow=[double]$font.Size
  $font.Size=[single]$size;$font.Shrink();$shrink=[double]$font.Size
  $report.steps+=@{size=$size;grow=$grow;shrink=$shrink}
 }
 $doc.Close([ref]0);$doc=$null
 if($Compare){
  $browser=Get-Content (Join-Path $root 'browser-report.json') -Raw|ConvertFrom-Json
  if(!$browser.passed -or $browser.sourceHash -ne $report.sourceHash){throw 'Stale browser source'}
  foreach($case in $browser.exports){
   $path=[string](Join-Path $root $case.file)
   if((Hash $path) -ne $case.sha256){throw 'Stale browser export'}
   $doc=$app.Documents.Open($path,$false,$true,$false);$actual=Sizes $doc
   $expected=@($sizes);$expected[1]=[double]$case.size
   if(($actual -join ',') -ne ($expected -join ',')){throw "Native export sizes differ: $($case.file)"}
   $doc.Close([ref]0);$doc=$null
   $report.cases+=@{file=$case.file;sha256=Hash $path;sizes=$actual}
  }
 }
 $report|ConvertTo-Json -Depth 8|Set-Content -Encoding UTF8 (Join-Path $root $(if($Compare){'native-report.json'}else{'source-native.json'}))
 Write-Output "Native font sizes: $($sizes.Count) source boundaries; $($report.cases.Count) exports checked."
} finally {
 try {if($doc){$doc.Close([ref]0)}} finally {if($app){try{if($null -ne $links){$app.Options.UpdateLinksAtOpen=$links};$app.Quit([ref]0)}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}}
}

