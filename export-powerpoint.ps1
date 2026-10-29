param([Parameter(Mandatory=$true)][string]$PlanPath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$powerPoint=$null;$presentation=$null;$previousSecurity=$null
$hadPowerPoint=$true;$officeMutex=New-Object Threading.Mutex($false,'Local\SlideStudioPowerPoint');$hasMutex=$false
$assetFolder=Split-Path -Parent $PlanPath
function RGB-Value([string]$hex){if($hex -notmatch '^#[0-9a-fA-F]{6}$'){return 0};$r=[Convert]::ToInt32($hex.Substring(1,2),16);$g=[Convert]::ToInt32($hex.Substring(3,2),16);$b=[Convert]::ToInt32($hex.Substring(5,2),16);return $r+256*$g+65536*$b}
function Image-File([string]$data){
  if($data -notmatch '^data:image/(png|jpeg);base64,([A-Za-z0-9+/=\r\n]+)$'){throw 'Invalid image data.'}
  $extension=if($Matches[1] -eq 'png'){'.png'}else{'.jpg'};$file=Join-Path $assetFolder (([Guid]::NewGuid().ToString())+$extension)
  [IO.File]::WriteAllBytes($file,[Convert]::FromBase64String($Matches[2]));return $file
}
function Apply-Font($shape,$o,$scale){
  $range=$shape.TextFrame.TextRange
  $range.Font.Name=[string]$o.fontFamily;$range.Font.Size=[single]([double]$o.fontSize*$scale)
  $range.Font.Bold=if($o.bold){-1}else{0};$range.Font.Italic=if($o.italic){-1}else{0};$range.Font.Color.RGB=RGB-Value $o.color
  switch([string]$o.align){'center'{$range.ParagraphFormat.Alignment=2}'right'{$range.ParagraphFormat.Alignment=3}'justify'{$range.ParagraphFormat.Alignment=4}default{$range.ParagraphFormat.Alignment=1}}
}
function Set-Geometry($shape,$o,$scale){
  try{$shape.LockAspectRatio=0}catch{};if($o.canRotate -ne $false -and [Math]::Abs([double]$shape.Rotation) -gt 0.001){$shape.Rotation=0}
  $shape.Left=[single]([double]$o.x*$scale);$shape.Top=[single]([double]$o.y*$scale)
  $shape.Width=[single]([Math]::Max(0.1,[double]$o.width*$scale));$shape.Height=[single]([Math]::Max(0.1,[double]$o.height*$scale));if($o.canRotate -ne $false -and [Math]::Abs([double]$o.rotation) -gt 0.001){$shape.Rotation=[single]$o.rotation}
}
function Add-Object($slide,$o,$scale){
  if($o.type -eq 'picture'){$image=Image-File $o.image;return $slide.Shapes.AddPicture($image,0,-1,[single]($o.x*$scale),[single]($o.y*$scale),[single]($o.width*$scale),[single]($o.height*$scale))}
  if($o.type -eq 'graphic'){if($o.image){$image=Image-File $o.image;return $slide.Shapes.AddPicture($image,0,-1,[single]($o.x*$scale),[single]($o.y*$scale),[single]($o.width*$scale),[single]($o.height*$scale))};throw 'Cannot create unsupported graphic without an image.'}
  if($o.geometry -eq 'line'){return $slide.Shapes.AddLine([single]($o.x*$scale),[single]($o.y*$scale),[single](($o.x+$o.width)*$scale),[single](($o.y+$o.height)*$scale))}
  $kind=1;if($o.geometry -eq 'roundRect'){$kind=5};if($o.geometry -eq 'ellipse'){$kind=9}
  if($o.type -eq 'text' -and -not $o.fillVisible -and $o.lineWidth -eq 0){return $slide.Shapes.AddTextbox(1,[single]($o.x*$scale),[single]($o.y*$scale),[single]($o.width*$scale),[single]($o.height*$scale))}
  return $slide.Shapes.AddShape($kind,[single]($o.x*$scale),[single]($o.y*$scale),[single]($o.width*$scale),[single]($o.height*$scale))
}
function Apply-Object($slide,$shape,$o,$scale,$isNew){
  if($o.changes.image -and $o.type -eq 'picture' -and -not $isNew){$z=[int]$shape.ZOrderPosition;$shape.Delete();$shape=Add-Object $slide $o $scale;try{$shape.Name=[string]$o.name}catch{};while($shape.ZOrderPosition -gt $z){$shape.ZOrder(3)};$isNew=$true}
  if($isNew -or $o.changes.geometry){Set-Geometry $shape $o $scale}
  if($isNew -or $o.changes.text){
    if($shape.HasTextFrame -eq -1 -and $o.type -ne 'table'){
      $shape.TextFrame.TextRange.Text=([string]$o.text).Replace("`n","`r");$shape.TextFrame.AutoSize=0;$shape.TextFrame.WordWrap=-1
      if($o.margins.Count -eq 4){$shape.TextFrame.MarginLeft=[single]($o.margins[0]*$scale);$shape.TextFrame.MarginTop=[single]($o.margins[1]*$scale);$shape.TextFrame.MarginRight=[single]($o.margins[2]*$scale);$shape.TextFrame.MarginBottom=[single]($o.margins[3]*$scale)}
      switch([string]$o.verticalAlign){'middle'{$shape.TextFrame.VerticalAnchor=3}'bottom'{$shape.TextFrame.VerticalAnchor=4}default{$shape.TextFrame.VerticalAnchor=1}}
      Apply-Font $shape $o $scale
    }
  }elseif($o.changes.font -or $o.changes.color){if($shape.HasTextFrame -eq -1 -and $o.type -ne 'table'){Apply-Font $shape $o $scale}}
  if(($isNew -or $o.changes.fill) -and $o.type -ne 'picture'){
    try{if($o.fillVisible){$shape.Fill.Visible=-1;$shape.Fill.Solid();$shape.Fill.ForeColor.RGB=RGB-Value $o.fill;$shape.Fill.Transparency=1-[double]$o.fillOpacity}else{$shape.Fill.Visible=0}}catch{}
  }
  if(($isNew -or $o.changes.line) -and $o.type -ne 'picture'){
    try{if($o.lineWidth -gt 0){$shape.Line.Visible=-1;$shape.Line.Weight=[single]($o.lineWidth*$scale);$shape.Line.ForeColor.RGB=RGB-Value $o.lineColor}else{$shape.Line.Visible=0}}catch{}
  }
  if($o.type -eq 'table' -and $o.changes.table){
    for($r=0;$r -lt $o.table.Count;$r++){for($c=0;$c -lt $o.table[$r].Count;$c++){$cell=$o.table[$r][$c];$range=$shape.Table.Cell($r+1,$c+1).Shape.TextFrame.TextRange;$text=([string]$cell.text).Replace("`n","`r");if([string]$range.Text -ne $text){$range.Text=$text}}}
  }
  return $shape
}
try{
  try{$hasMutex=$officeMutex.WaitOne(180000)}catch [Threading.AbandonedMutexException]{$hasMutex=$true};if(-not $hasMutex){throw 'PowerPoint is busy. Please try again.'}
  $hadPowerPoint=@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
  $plan=Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $powerPoint=New-Object -ComObject PowerPoint.Application;$previousSecurity=$powerPoint.AutomationSecurity;$powerPoint.AutomationSecurity=3
  if($plan.sourcePath){$presentation=$powerPoint.Presentations.Open([string]$plan.sourcePath,-1,0,0)}else{$presentation=$powerPoint.Presentations.Add(0);$presentation.PageSetup.SlideWidth=960;$presentation.PageSetup.SlideHeight=[single](960*$plan.slides[0].height/$plan.slides[0].width)}
  $tag='SLIDESTUDIO_TMP_'+[Guid]::NewGuid().ToString('N')
  $sourceSlides=@{}
  foreach($slide in $presentation.Slides){$sourceSlides[[string]$slide.SlideID]=$slide;foreach($shape in $slide.Shapes){$shape.Tags.Add($tag,[string]$shape.Id)}}
  $position=0
  foreach($item in $plan.slides){
    $position++
    if($item.sourceSlideId){$original=$sourceSlides[[string]$item.sourceSlideId];if(-not $original){throw 'Source slide no longer exists.'};$range=$original.Duplicate();$slide=$range.Item(1);$slide.MoveTo($position)}else{$slide=$presentation.Slides.Add($position,12);$slide.DisplayMasterShapes=0;$slide.FollowMasterBackground=0;$slide.Background.Fill.Solid();$slide.Background.Fill.ForeColor.RGB=RGB-Value $item.backgroundColor}
    if($item.title){$slide.Tags.Add('SLIDESTUDIO_TITLE',[string]$item.title)}
    $scale=[double]$presentation.PageSetup.SlideWidth/[double]$item.width
    if($item.backgroundImage){$image=Image-File $item.backgroundImage;$picture=$slide.Shapes.AddPicture($image,0,-1,0,0,$presentation.PageSetup.SlideWidth,$presentation.PageSetup.SlideHeight);$picture.ZOrder(1)}
    $map=@{};foreach($shape in $slide.Shapes){$sourceId=$shape.Tags.Item($tag);if($sourceId){$map[[string]$sourceId]=$shape}}
    $clones=@{};foreach($o in $item.objects){if($o.cloneSourceId){if(-not $map.ContainsKey([string]$o.cloneSourceId)){throw 'Clone source shape not found.'};$clones[[string]$o.id]=$map[[string]$o.cloneSourceId].Duplicate().Item(1)}}
    foreach($o in $item.objects){
      $isNew=$null -eq $o.sourceId
      if($isNew -and $o.cloneSourceId){$shape=$clones[[string]$o.id];$isNew=$false}elseif($isNew){$shape=Add-Object $slide $o $scale;try{$shape.Name=[string]$o.name}catch{}}else{$shape=$map[[string]$o.sourceId];if(-not $shape){throw ('Source shape not found: '+$o.sourceId)}}
      $shape=Apply-Object $slide $shape $o $scale $isNew
      if($item.reorderObjects){$shape.ZOrder(0)}
    }
    foreach($id in $item.deletedObjects){if($map.ContainsKey([string]$id)){$map[[string]$id].Delete()}}
    foreach($shape in $slide.Shapes){try{$shape.Tags.Delete($tag)}catch{}}
    $written=$false;foreach($note in $slide.NotesPage.Shapes){try{if($note.Type -eq 14 -and $note.PlaceholderFormat.Type -eq 2){$note.TextFrame.TextRange.Text=([string]$item.notes).Replace("`n","`r");$written=$true;break}}catch{}}
    if(-not $written -and $item.notes){$note=$slide.NotesPage.Shapes.AddTextbox(1,50,100,500,300);$note.TextFrame.TextRange.Text=[string]$item.notes;$note.Tags.Add('SLIDESTUDIO_NOTES','1')}
  }
  for($i=$presentation.Slides.Count;$i -gt $position;$i--){$presentation.Slides.Item($i).Delete()}
  $presentation.SaveAs($OutputPath,24)
}catch{[Console]::Error.WriteLine($_.Exception.Message);exit 1}
finally{
  if($presentation){try{$presentation.Saved=-1;$presentation.Close()}catch{}}
  if($powerPoint){if($null -ne $previousSecurity){try{$powerPoint.AutomationSecurity=$previousSecurity}catch{}};if(-not $hadPowerPoint){try{if($powerPoint.Presentations.Count -eq 0){$powerPoint.Quit()}}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  if($hasMutex){$officeMutex.ReleaseMutex()};$officeMutex.Dispose()
}
