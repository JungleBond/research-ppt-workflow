param(
 [Parameter(Mandatory=$true)][string]$Presentation,
 [Parameter(Mandatory=$true)][string]$ReportDirectory
)
$ErrorActionPreference='Stop'
$src=(Resolve-Path -LiteralPath $Presentation).Path
$out=[IO.Path]::GetFullPath($ReportDirectory)
if($src.StartsWith($out+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){
 throw 'Keep the final presentation outside the verification directory.'
}
New-Item -ItemType Directory -Path $out -Force|Out-Null
$copy=Join-Path $out 'verification-copy.pptx'
Copy-Item -LiteralPath $src -Destination $copy -Force
$before=(Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash
$app=New-Object -ComObject PowerPoint.Application
$deck=$null;$chartTests=@()
try{
 $deck=$app.Presentations.Open($copy,0,0,0)
 foreach($slide in $deck.Slides){
  $slide.Export((Join-Path $out ('slide-'+('{0:D2}' -f $slide.SlideIndex)+'.png')),'PNG',1600,900)
 }
 $deck.SaveAs((Join-Path $out 'final-preview.pdf'),32)
 foreach($slide in $deck.Slides){
  foreach($shape in $slide.Shapes){
   if($shape.HasChart -eq -1){
    $chart=$shape.Chart
    $series=$chart.SeriesCollection(1)
    $old=@($series.Values)
    $changed=@($old);$changed[0]=[double]$old[0]+1
    $series.Values=$changed
    $readback=@($series.Values)
    $dataPassed=[Math]::Abs([double]$readback[0]-[double]$changed[0]) -lt 0.000001
    $series.Values=$old
    $width=$shape.Width
    $shape.Width=[double]($width+12)
    $sizePassed=[Math]::Abs($shape.Width-($width+12)) -lt 0.01
    $shape.Width=[double]$width
    $chartTests += [pscustomobject]@{
     Slide=$slide.SlideIndex
     OriginalValues=$old
     EditedValue=$readback[0]
     DataEditPassed=$dataPassed
     ResizePassed=$sizePassed
     OriginalValuesRestored=$true
    }
    if(-not $dataPassed -or -not $sizePassed){throw 'Native chart edit test failed.'}
   }
  }
 }
 $count=$deck.Slides.Count
 $deck.Saved=-1
 $after=(Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash
 if($before -ne $after){throw 'The final file unexpectedly changed.'}
 [pscustomobject]@{
  Slides=$count;PowerPointReopened=$true;AllSlidesRendered=$true
  FinalSha256=$after;FinalUnchanged=$true;ChartTests=$chartTests
  Scope='Native PowerPoint reopen/render and temporary chart edit/resize; no video playback test.'
 }|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $out 'native-verification.json') -Encoding UTF8
 Write-Output "Verified $count slides in PowerPoint. Final file unchanged."
}finally{
 if($deck){$deck.Close()}
 [Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null
}
