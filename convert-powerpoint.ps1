param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$presentation=$null; $powerPoint=$null; $previousSecurity=$null
$hadPowerPoint=$true;$officeMutex=New-Object Threading.Mutex($false,'Local\SlideStudioPowerPoint');$hasMutex=$false
function Color-Hex($value) { $n=[long]$value;return '#{0:X2}{1:X2}{2:X2}' -f ($n -band 255),(($n -shr 8) -band 255),(($n -shr 16) -band 255) }
function Font-Info($range,$scale) {
  $font=$range.Font
  if($range.Length -gt 0){$font=$range.Characters(1,1).Font}
  $size=24;try{if($font.Size -gt 0){$size=[double]$font.Size*$scale}}catch{}
  $name='Malgun Gothic';try{if($font.Name){$name=[string]$font.Name}}catch{}
  $color='#20374d';try{$color=Color-Hex $font.Color.RGB}catch{}
  $bold=$false;try{$bold=$font.Bold -eq -1}catch{}
  $italic=$false;try{$italic=$font.Italic -eq -1}catch{}
  return @{fontFamily=$name;fontSize=$size;color=$color;bold=$bold;italic=$italic}
}
try {
  try{$hasMutex=$officeMutex.WaitOne(180000)}catch [Threading.AbandonedMutexException]{$hasMutex=$true};if(-not $hasMutex){throw 'PowerPoint is busy. Please try again.'}
  $hadPowerPoint=@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
  $powerPoint=New-Object -ComObject PowerPoint.Application
  $previousSecurity=$powerPoint.AutomationSecurity;$powerPoint.AutomationSecurity=3
  $presentation=$powerPoint.Presentations.Open($SourcePath,-1,0,0)
  if($presentation.Slides.Count -gt 200){throw '200 slides maximum.'}
  $normalized=Join-Path $OutputPath 'original.pptx'
  $presentation.SaveCopyAs($normalized,24)
  $width=[double]$presentation.PageSetup.SlideWidth;$height=[double]$presentation.PageSetup.SlideHeight
  $imageWidth=1600;$imageHeight=[int][Math]::Round($imageWidth*$height/$width);$scale=$imageWidth/$width
  $slides=@()
  foreach($slide in $presentation.Slides){
    $number=[int]$slide.SlideIndex;$imageName="slide-$number.png"
    $slide.Export((Join-Path $OutputPath $imageName),'PNG',$imageWidth,$imageHeight)
    $title="Slide $number";try{if($slide.Shapes.HasTitle){$title=[string]$slide.Shapes.Title.TextFrame.TextRange.Text}}catch{}
    $notes=@();foreach($noteShape in $slide.NotesPage.Shapes){try{if($noteShape.HasTextFrame -and $noteShape.TextFrame.HasText -and (($noteShape.Type -eq 14 -and $noteShape.PlaceholderFormat.Type -eq 2) -or $noteShape.Tags.Item('SLIDESTUDIO_NOTES') -eq '1')){$notes += [string]$noteShape.TextFrame.TextRange.Text}}catch{}}
    $objects=@();$frozen=0
    foreach($shape in $slide.Shapes){
      if($shape.Visible -eq 0){continue}
      $id=[int]$shape.Id;$rotation=[double]$shape.Rotation
      $o=@{id="ppt-$number-$id";sourceId=$id;name=[string]$shape.Name;type='graphic';x=[double]$shape.Left*$scale;y=[double]$shape.Top*$scale;width=[Math]::Max(1,[double]$shape.Width*$scale);height=[Math]::Max(1,[double]$shape.Height*$scale);rotation=$rotation;geometry='custom';changes=@{};text='';fontFamily='Malgun Gothic';fontSize=32;color='#20374d';bold=$false;italic=$false;align='left';verticalAlign='top';margins=@(0,0,0,0);fill='#FFFFFF';fillVisible=$false;fillOpacity=1;lineWidth=0;lineColor='#197b92'}
      $hasText=$false;try{$hasText=($shape.HasTextFrame -eq -1) -and (($shape.TextFrame.HasText -eq -1) -or $shape.Type -eq 17 -or $shape.Type -eq 14)}catch{}
      $o.canRotate=([int]$shape.Type -notin @(3,19,24))
      if($shape.Type -eq 13 -or $shape.Type -eq 11){$o.type='picture'}
      if($shape.Type -eq 17 -or $shape.Type -eq 14){$o.geometry='rect'}
      if($shape.Type -eq 1){$o.type='shape';switch([int]$shape.AutoShapeType){1{$o.geometry='rect'}5{$o.geometry='roundRect'}9{$o.geometry='ellipse'}}}
      if($shape.Type -eq 9){$o.type='shape';$o.geometry='line'}
      try{if($shape.Fill.Visible -eq -1){$o.fillVisible=$true;$o.fill=Color-Hex $shape.Fill.ForeColor.RGB;$o.fillOpacity=1-[double]$shape.Fill.Transparency}}catch{}
      try{if($shape.Line.Visible -eq -1){$o.lineWidth=[double]$shape.Line.Weight*$scale;$o.lineColor=Color-Hex $shape.Line.ForeColor.RGB}}catch{}
      if($shape.Type -eq 14 -and $shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -ne -1 -and -not $o.fillVisible -and $o.lineWidth -eq 0){continue}
      if($hasText){
        $o.type='text';$range=$shape.TextFrame.TextRange;$o.text=([string]$range.Text).Replace("`r","`n").Replace([string][char]11,"`n")
        $font=Font-Info $range $scale;foreach($key in $font.Keys){$o[$key]=$font[$key]}
        try{$o.margins=@([double]$shape.TextFrame.MarginLeft*$scale,[double]$shape.TextFrame.MarginTop*$scale,[double]$shape.TextFrame.MarginRight*$scale,[double]$shape.TextFrame.MarginBottom*$scale)}catch{}
        try{switch([int]$range.ParagraphFormat.Alignment){2{$o.align='center'}3{$o.align='right'}4{$o.align='justify'}}}catch{}
        try{switch([int]$shape.TextFrame.VerticalAnchor){3{$o.verticalAlign='middle'}4{$o.verticalAlign='bottom'}}}catch{}
      }
      try{if($shape.HasTable -eq -1){
        $o.type='table';$rows=@()
        for($r=1;$r -le $shape.Table.Rows.Count;$r++){$cells=@();for($c=1;$c -le $shape.Table.Columns.Count;$c++){$cellShape=$shape.Table.Cell($r,$c).Shape;$cell=Font-Info $cellShape.TextFrame.TextRange $scale;$cell.text=([string]$cellShape.TextFrame.TextRange.Text).Replace("`r","`n");$cell.width=[double]$shape.Table.Columns.Item($c).Width*$scale;$cell.height=[double]$shape.Table.Rows.Item($r).Height*$scale;$cell.fill='#FFFFFF';$cell.align='left';try{$cell.fill=Color-Hex $cellShape.Fill.ForeColor.RGB}catch{};try{switch([int]$cellShape.TextFrame.TextRange.ParagraphFormat.Alignment){2{$cell.align='center'}3{$cell.align='right'}}}catch{};$cells+= $cell};$rows+= ,$cells};$o.table=$rows
      }}catch{}
      $preview="shape-$number-$id.png";$backdrop="shape-$number-$id-base.png"
      try{
        if([Math]::Abs($rotation) -gt 0.001){$shape.Rotation=0}
        $shape.Export((Join-Path $OutputPath $preview),2,[int]($imageWidth*0.75),[int]($imageHeight*0.75),1)
        $o.preview=$preview
        if($hasText -and $o.type -ne 'table'){try{$shape.TextFrame.TextRange.Text='';$shape.Export((Join-Path $OutputPath $backdrop),2,[int]($imageWidth*0.75),[int]($imageHeight*0.75),1);$o.backdrop=$backdrop}catch{}}
        $shape.Visible=0;$objects+= $o
      }catch{try{if([Math]::Abs($rotation) -gt 0.001){$shape.Rotation=$rotation}}catch{};$frozen++}
    }
    if($title -eq "Slide $number"){try{$savedTitle=$slide.Tags.Item('SLIDESTUDIO_TITLE');if($savedTitle){$title=$savedTitle}else{$firstText=@($objects|Where-Object{$_.text})[0];if($firstText){$title=([string]$firstText.text).Split("`n")[0]}}}catch{}}
    $background="background-$number.png";$slide.Export((Join-Path $OutputPath $background),'PNG',$imageWidth,$imageHeight)
    $slides+=@{title=$title.Trim();image=$imageName;background=$background;objects=$objects;frozen=$frozen;sourceSlideId=[int]$slide.SlideID;notes=(($notes -join "`n`n") -replace "`r","`n");width=$imageWidth;height=$imageHeight;pointWidth=$width}
  }
  $result=@{slides=$slides;width=$imageWidth;height=$imageHeight;pptx='original.pptx';editable=$true}
  [IO.File]::WriteAllText((Join-Path $OutputPath 'result.json'),(ConvertTo-Json -InputObject $result -Depth 16 -Compress),(New-Object Text.UTF8Encoding($false)))
}catch{[Console]::Error.WriteLine($_.Exception.Message);exit 1}
finally{
  if($presentation){try{$presentation.Saved=-1;$presentation.Close()}catch{}}
  if($powerPoint){if($null -ne $previousSecurity){try{$powerPoint.AutomationSecurity=$previousSecurity}catch{}};if(-not $hadPowerPoint){try{if($powerPoint.Presentations.Count -eq 0){$powerPoint.Quit()}}catch{}};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
  if($hasMutex){$officeMutex.ReleaseMutex()};$officeMutex.Dispose()
}
