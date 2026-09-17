param([string]$Case='terminal-column',[switch]$Probe,[switch]$NoSentenceCaps,[ValidateSet('original','start','selected')][string]$Context='original',[switch]$BookmarkProbe)
$ErrorActionPreference='Stop'
$workspace=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$caseId="$Case$(if($Context -ne 'original'){"-$Context"})"
$directory=Join-Path $workspace ".local/word-inline-ui/$caseId$(if($NoSentenceCaps){'-no-caps'})$(if($BookmarkProbe){'-bookmark-probe'})"
if($BookmarkProbe -and !$Probe){throw 'Bookmark authoring is discovery-only'}
$referencePath=Join-Path $workspace 'tests/fixtures/native-word-inline-flow.json'
$reference=(Get-Content $referencePath -Raw|ConvertFrom-Json).cases.$Case
if(!$reference){throw 'Unknown inline fixture'}
$commandReferencePath=Join-Path $workspace 'tests/fixtures/native-word-ui-flow.json'
$commandReference=(Get-Content $commandReferencePath -Raw|ConvertFrom-Json).cases.$caseId
if(!$Probe -and (!$commandReference -or !$NoSentenceCaps)){throw 'Command acceptance currently requires a captured case and -NoSentenceCaps; use -Probe for further discovery'}
[void][IO.Directory]::CreateDirectory($directory)
function Hash($path){
 $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $digest=[Security.Cryptography.SHA256]::Create()
 try{return ([BitConverter]::ToString($digest.ComputeHash($stream))).Replace('-','').ToLowerInvariant()}
 finally{$stream.Dispose();$digest.Dispose()}
}
. (Join-Path $workspace 'scripts/excel-test-window.ps1')
function Retry([scriptblock]$action){for($attempt=0;$attempt -lt 24;$attempt++){try{return (& $action)}catch{if($_.Exception.ToString() -notmatch 'RPC_E_CALL_REJECTED|0x80010001|0x8001010A|Native document body is not ready' -or $attempt -eq 23){throw};if($script:ownedProcess){[ExcelTestWindow]::DismissReminder($script:ownedProcess)};Start-Sleep -Milliseconds 250}}}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class OwnedWordFlowInput {
 [StructLayout(LayoutKind.Sequential)] struct Keyboard {public ushort key,scan;public uint flags,time;public UIntPtr extra;}
 [StructLayout(LayoutKind.Sequential)] struct Mouse {public int x,y;public uint data,flags,time;public UIntPtr extra;}
 [StructLayout(LayoutKind.Explicit)] struct Payload {[FieldOffset(0)]public Keyboard keyboard;[FieldOffset(0)]public Mouse mouse;}
 [StructLayout(LayoutKind.Sequential)] struct Input {public uint type;public Payload payload;}
 [DllImport("user32.dll",SetLastError=true)] static extern uint SendInput(uint count,Input[] inputs,int size);
 [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
 public static void Click(int x,int y) {
  if(!SetCursorPos(x,y))throw new InvalidOperationException("Native pointer move rejected");
  var inputs=new[]{new Input{type=0,payload=new Payload{mouse=new Mouse{flags=2}}},new Input{type=0,payload=new Payload{mouse=new Mouse{flags=4}}}};
  if(SendInput(2,inputs,Marshal.SizeOf(typeof(Input)))!=2)throw new InvalidOperationException("Native pointer injection rejected");
 }
 public static void Press(byte[] keys) {
  var inputs=new Input[keys.Length*2];
  for(int i=0;i<keys.Length;i++){
   ushort key=keys[i];uint extended=(key==35||key==36||key==39||key==46)?1u:0u;
   inputs[i]=new Input{type=1,payload=new Payload{keyboard=new Keyboard{key=key,flags=extended}}};
   inputs[inputs.Length-i-1]=new Input{type=1,payload=new Payload{keyboard=new Keyboard{key=key,flags=extended|2u}}};
  }
  if(SendInput((uint)inputs.Length,inputs,Marshal.SizeOf(typeof(Input)))!=inputs.Length)throw new InvalidOperationException("Native keyboard injection was rejected");
 }
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h,uint flags);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint id);
}
'@
function Keys($keys){
 $frame=(Get-Process -Id $script:ownedProcess).MainWindowHandle
 if($frame -eq [IntPtr]::Zero){throw 'Owned Word frame is not available'}
 [void][OwnedWordFlowInput]::SetForegroundWindow($frame)
 $element=[Windows.Automation.AutomationElement]::FromHandle($frame)
 $focusable=$element.FindAll([Windows.Automation.TreeScope]::Descendants,(New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::IsKeyboardFocusableProperty,$true)))
 $focusable|ForEach-Object{@{name=$_.Current.Name;type=$_.Current.ControlType.ProgrammaticName;class=$_.Current.ClassName;id=$_.Current.AutomationId}}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $directory 'focusable.json') -Encoding UTF8
 $documents=@($focusable|Where-Object {$_.Current.AutomationId -eq 'UIA_AutomationId_Word_Content_Page_1'})
 if($documents.Count -ne 1){throw "Expected one owned editable Word document; found $($documents.Count)"}
 if(!$script:inputFocused){
  [int]$screenLeft=0;[int]$screenTop=0;[int]$screenWidth=0;[int]$screenHeight=0
  Retry {$app.ActiveWindow.GetPoint([ref]$screenLeft,[ref]$screenTop,[ref]$screenWidth,[ref]$screenHeight,$doc.Range(0,1))}
  $point=New-Object Windows.Point(($screenLeft+[Math]::Max(1,$screenWidth/2)),($screenTop+[Math]::Max(1,$screenHeight/2)))
  if([Windows.Automation.AutomationElement]::FromPoint($point).Current.ProcessId -ne $script:ownedProcess){throw 'Native text point lies outside owned Word'}
  [OwnedWordFlowInput]::Click([int]$point.X,[int]$point.Y)
  $script:inputFocused=$true
 }
 Start-Sleep -Milliseconds 100
 [uint32]$front=0
 [void][OwnedWordFlowInput]::GetWindowThreadProcessId([OwnedWordFlowInput]::GetForegroundWindow(),[ref]$front)
 if($front -ne $script:ownedProcess){throw "Refusing keyboard input outside the owned Word process: expected $script:ownedProcess, actual $front, frame $frame"}
 Add-Content -LiteralPath (Join-Path $directory 'keys.log') -Value $keys -Encoding UTF8
 $focused=[Windows.Automation.AutomationElement]::FocusedElement
 Add-Content -LiteralPath (Join-Path $directory 'keys.log') -Value "focus: $($focused.Current.Name) / $($focused.Current.ControlType.ProgrammaticName); insertEnabled: $($app.CommandBars.GetEnabledMso('PageBreakInsertWord'))" -Encoding UTF8
 $virtualKeys=switch($keys){'^{HOME}'{@(17,36)};'{RIGHT}'{@(39)};'+{RIGHT}'{@(16,39)};'{DELETE}'{@(46)};'^{ENTER}'{@(17,13)};'^+{ENTER}'{@(17,16,13)};'^z'{@(17,90)};'^y'{@(17,89)};'^{END}'{@(17,35)};'!'{@(16,49)};default{throw 'Unexpected native key command'}}
 [OwnedWordFlowInput]::Press([byte[]]$virtualKeys)
 Start-Sleep -Milliseconds 100
 Add-Content -LiteralPath (Join-Path $directory 'keys.log') -Value "selection: $($app.Selection.Start)/$($app.Selection.End)" -Encoding UTF8
}
function Snapshot($document){
 return Retry {
 $document.Repaginate();$characters=@();$paragraphs=@()
 $paragraphCount=[int]$document.Paragraphs.Count
 $bodyText=[string]$document.Content.Text
 if($paragraphCount -lt 1 -or !$bodyText){throw 'Native document body is not ready'}
 for($paragraphIndex=1;$paragraphIndex -le $paragraphCount;$paragraphIndex++){
 $paragraph=$document.Paragraphs.Item($paragraphIndex)
 $p=$paragraph.Range;$localCharacters=@()
 for($i=[int]$p.Start;$i -lt [int]$p.End-1;$i++){
  $c=$document.Range($i,$i+1)
  $localCharacters+=@{text=[string]$c.Text;page=[int]$c.Information(3);x=[double]$c.Information(5);y=[double]$c.Information(6);font=[string]$c.Font.Name;size=[double]$c.Font.Size}
 }
 $caret=$document.Range([int]$p.End-1,[int]$p.End-1)
 $lastCaret=@{page=[int]$caret.Information(3);x=[double]$caret.Information(5);y=[double]$caret.Information(6)}
 $paragraphs+=@{text=[string]$p.Text;characters=$localCharacters;caret=$lastCaret;lineRule=[int]$paragraph.Format.LineSpacingRule;lineSpacing=[double]$paragraph.Format.LineSpacing;keepLines=[int]$paragraph.Format.KeepTogether;keepNext=[int]$paragraph.Format.KeepWithNext;widowControl=[int]$paragraph.Format.WidowControl}
 $characters+=$localCharacters
 }
 if($bodyText -cne (($paragraphs|ForEach-Object{$_.text}) -join '')){throw 'Native document body is not ready: paragraph snapshot differs'}
 return @{text=$bodyText;characters=$characters;caret=$lastCaret;paragraphs=$paragraphs;selection=@{from=[int]$app.Selection.Start;to=[int]$app.Selection.End}}
 }
}
function AssertSnapshot($expected,$actual){
 if($expected.text -cne $actual.text -or $expected.characters.Count -ne $actual.characters.Count){throw 'Native UI text differs'}
 for($i=0;$i -lt $expected.characters.Count;$i++){
  $a=$expected.characters[$i];$b=$actual.characters[$i]
  if($a.text -cne $b.text -or $a.page -ne $b.page -or [Math]::Abs($a.x-$b.x) -gt .001 -or [Math]::Abs($a.y-$b.y) -gt .001){throw "Native UI character differs: $i"}
 }
 $a=$expected.caret;$b=$actual.caret
 if($a.page -ne $b.page -or [Math]::Abs($a.x-$b.x) -gt .001 -or [Math]::Abs($a.y-$b.y) -gt .001){throw 'Native UI terminal caret differs'}
}
$app=$null;$doc=$null;$idle=$false
try{
 $fixture=[string](Join-Path $workspace "tests/fixtures/word-inline-flow-$Case.docx")
 if((Hash $fixture) -ne $reference.sourceSha256){throw 'Source binding changed'}
 $app=New-Object -ComObject Word.Application
 if($app.Documents.Count -ne 0){throw 'Requires an idle owned Word instance'};$idle=$true
 $security=$app.AutomationSecurity;$alerts=$app.DisplayAlerts;$links=$app.Options.UpdateLinksAtOpen;$sentenceCaps=$app.AutoCorrect.CorrectSentenceCaps
 @{security=$security;alerts=$alerts;links=$links;sentenceCaps=$sentenceCaps}|ConvertTo-Json|Set-Content (Join-Path $directory 'preferences.json') -Encoding UTF8
 if($NoSentenceCaps){$app.AutoCorrect.CorrectSentenceCaps=$false}
 $app.AutomationSecurity=3;$app.DisplayAlerts=0;$app.Options.UpdateLinksAtOpen=$false
 $exe=Hash (Join-Path ([string]$app.Path) 'WINWORD.EXE');$font=Hash (Join-Path $env:WINDIR 'Fonts/arial.ttf')
 if($exe -ne $reference.native.executableSha256 -or $font -ne $reference.native.fontSha256){throw 'Native baseline changed'}
 $doc=$app.Documents.Open($fixture,$false,$false,$false)
 $handle=[IntPtr]$doc.ActiveWindow.Hwnd;[uint32]$script:ownedProcess=0
 [void][OwnedWordFlowInput]::GetWindowThreadProcessId($handle,[ref]$script:ownedProcess)
 if(!$script:ownedProcess){throw 'Owned Word process unavailable'}
 Retry {$app.Visible=$true}
 [ExcelTestWindow]::DismissReminder($script:ownedProcess)
 [void][OwnedWordFlowInput]::SetForegroundWindow($handle);Retry {$doc.Activate()}
 $bookmarkedInput=$null
 if($BookmarkProbe){
  [void]$doc.Bookmarks.Add('NofficeBoundaryAnchor',$doc.Range(0,0))
  $bookmarkedInput=[string](Join-Path $directory 'bookmarked-source.docx')
  $doc.SaveAs2([ref]$bookmarkedInput,[ref]12)
 }
 $initial=Snapshot $doc
 $initial|ConvertTo-Json -Depth 12|Set-Content (Join-Path $directory 'initial.json') -Encoding UTF8
 AssertSnapshot $reference.native.paragraphs[0] $initial
 $index=([string]$initial.text).IndexOfAny([char[]]@([char]12,[char]14))
 if($index -lt 0){throw 'No source flow break'}
 $kind=([string]$initial.text)[$index]
 $binding=$app.FindKey($(if($kind -eq [char]12){$app.BuildKeyCode(512,[ref]13)}else{$app.BuildKeyCode(512,[ref]256,[ref]13)}))
 @{command=[string]$binding.Command;keyString=[string]$binding.KeyString;category=[int]$binding.KeyCategory}|ConvertTo-Json|Set-Content (Join-Path $directory 'binding.json') -Encoding UTF8
 Keys '^{HOME}';for($i=0;$i -lt $index;$i++){Keys '{RIGHT}'}
 Keys '{DELETE}';$deleted=Snapshot $doc
 $deleted|ConvertTo-Json -Depth 8|Set-Content (Join-Path $directory 'deleted.json') -Encoding UTF8
 if($deleted.text -cne ([string]$initial.text).Remove($index,1)){throw 'Native UI deletion did not remove exactly the break'}
 if($Context -eq 'start'){Keys '^{HOME}'}
 if($Context -eq 'selected'){for($selectionStep=0;$selectionStep -lt 3;$selectionStep++){Keys '+{RIGHT}'}}
 $beforeCommand=Snapshot $doc
 Keys $(if($kind -eq [char]12){'^{ENTER}'}else{'^+{ENTER}'})
 $inserted=Snapshot $doc
 if($BookmarkProbe){
  $anchor=$doc.Bookmarks.Item('NofficeBoundaryAnchor').Range
  @{sourceHash=Hash $bookmarkedInput;insertedAnchor=@{from=[int]$anchor.Start;to=[int]$anchor.End};beforeCommand=$beforeCommand;inserted=$inserted}|ConvertTo-Json -Depth 16|Set-Content (Join-Path $directory 'bookmark-discovery.json') -Encoding UTF8
 }
 @{initial=$initial;inserted=$inserted;paragraphs=@($doc.Paragraphs|ForEach-Object{[string]$_.Range.Text})}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $directory 'inserted.json') -Encoding UTF8
 if($Probe){
  Keys '^z';$undo=Snapshot $doc
  Keys '^y';$redo=Snapshot $doc
  $saved=[string](Join-Path $directory 'ui-inserted.docx');$doc.SaveAs2([ref]$saved,[ref]12)
  $doc.Close([ref]0);$doc=$app.Documents.Open($saved,$false,$false,$false);$doc.Activate();$script:inputFocused=$false
  $reopened=Snapshot $doc;AssertSnapshot $redo $reopened
  Keys '^{END}';Keys '!';$typed=Snapshot $doc
  Keys '^z';$typingUndo=Snapshot $doc
  Keys '^y';$typingRedo=Snapshot $doc
  $edited=[string](Join-Path $directory 'ui-typed.docx');$doc.SaveAs2([ref]$edited,[ref]12)
  $doc.Close([ref]0);$doc=$app.Documents.Open($edited,$false,$true,$false)
  AssertSnapshot $typingRedo (Snapshot $doc)
  $pdf=[string](Join-Path $directory 'ui-typed.pdf');$doc.ExportAsFixedFormat($pdf,17)
  @{result='discovery';case=$caseId;sourceCase=$Case;context=$Context;sourceHash=Hash $fixture;referenceHash=Hash $referencePath;scriptHash=Hash $PSCommandPath;executableHash=$exe;fontHash=$font;printer=[string]$app.ActivePrinter;version=[string]$app.Version;build=[string]$app.Build;sentenceCaps=[bool]$app.AutoCorrect.CorrectSentenceCaps;initial=$initial;deleted=$deleted;beforeCommand=$beforeCommand;inserted=$inserted;undo=$undo;redo=$redo;reopened=$reopened;typed=$typed;typingUndo=$typingUndo;typingRedo=$typingRedo;savedHash=Hash $saved;typedHash=Hash $edited;pdfHash=Hash $pdf}|ConvertTo-Json -Depth 18|Set-Content (Join-Path $directory 'discovery.json') -Encoding UTF8
  Write-Output "Captured native UI insertion/history/typing/reopen for discovery: $Case"
  return
 }
 AssertSnapshot $commandReference.steps.inserted $inserted
 if($inserted.selection.from -ne $commandReference.steps.inserted.selection.from -or $inserted.selection.to -ne $commandReference.steps.inserted.selection.to){throw 'Native command selection differs'}
 Keys '^z';$undo=Snapshot $doc;AssertSnapshot $deleted $undo
 Keys '^y';$redo=Snapshot $doc;AssertSnapshot $inserted $redo
 $restored=[string](Join-Path $directory 'ui-restored.docx');$doc.SaveAs2([ref]$restored,[ref]12)
 $doc.Close([ref]0);$doc=$app.Documents.Open($restored,$false,$false,$false);$doc.Activate()
 $script:inputFocused=$false
 AssertSnapshot $inserted (Snapshot $doc)
 [void][OwnedWordFlowInput]::SetForegroundWindow((Get-Process -Id $script:ownedProcess).MainWindowHandle)
 Keys '^{END}';Keys '!';$typed=Snapshot $doc
 AssertSnapshot $commandReference.steps.typed $typed
 Keys '^z';AssertSnapshot $inserted (Snapshot $doc)
 Keys '^y';AssertSnapshot $typed (Snapshot $doc)
 $edited=[string](Join-Path $directory 'ui-edited.docx');$doc.SaveAs2([ref]$edited,[ref]12)
 $doc.Close([ref]0);$doc=$app.Documents.Open($edited,$false,$true,$false)
 AssertSnapshot $typed (Snapshot $doc)
 $pdf=[string](Join-Path $directory 'ui-edited.pdf');$doc.ExportAsFixedFormat($pdf,17)
 @{passed=$true;case=$Case;sourceHash=Hash $fixture;referenceHash=Hash $referencePath;commandReferenceHash=Hash $commandReferencePath;scriptHash=Hash $PSCommandPath;executableHash=$exe;fontHash=$font;printer=[string]$app.ActivePrinter;sentenceCaps=[bool]$app.AutoCorrect.CorrectSentenceCaps;initial=$initial;deleted=$deleted;inserted=$inserted;undo=$undo;redo=$redo;typed=$typed;restoredHash=Hash $restored;editedHash=Hash $edited;pdfHash=Hash $pdf}|ConvertTo-Json -Depth 16|Set-Content (Join-Path $directory 'ui-report.json') -Encoding UTF8
 Write-Output "Native UI break insertion/deletion/history/typing/reopen verified: $Case"
}finally{
 if($doc){Retry {$doc.Close([ref]0)}}
 if($app){if($idle){Retry {$app.AutomationSecurity=$security;$app.DisplayAlerts=$alerts;$app.Options.UpdateLinksAtOpen=$links;if($null -ne $sentenceCaps){$app.AutoCorrect.CorrectSentenceCaps=$sentenceCaps};if($app.Documents.Count -eq 0){$app.Quit()}}};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}

