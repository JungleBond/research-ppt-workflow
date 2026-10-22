param([Parameter(Mandatory=$true)][string]$BuildDirectory)
$ErrorActionPreference='Stop'
$w=(Resolve-Path -LiteralPath $BuildDirectory).Path
$app=New-Object -ComObject PowerPoint.Application
$deck=$null;$audit=@();$tableAudit=@();$chartAudit=@()
$content=Get-Content -LiteralPath (Join-Path $w 'deck.resolved.json') -Raw -Encoding UTF8|ConvertFrom-Json
$highlightAudit=@()
function Set-Emphasis($shape,$terms,$page){
 $tr=$shape.TextFrame.TextRange
 foreach($term in $terms){
  $offset=0
  while($offset -lt $tr.Text.Length){
   $at=$tr.Text.IndexOf([string]$term,$offset,[StringComparison]::Ordinal)
   if($at -lt 0){break}
   # A numeric emphasis token must not match a digit inside another value.
   if(([string]$term) -match '^\d+(\.\d+)?$'){
    $before=if($at -gt 0){$tr.Text.Substring($at-1,1)}else{''}
    $afterAt=$at+([string]$term).Length
    $after=if($afterAt -lt $tr.Text.Length){$tr.Text.Substring($afterAt,1)}else{''}
    if($before -match '[0-9.]' -or $after -match '[0-9.]'){$offset=$afterAt;continue}
   }
   $run=$tr.Characters($at+1,([string]$term).Length)
   $run.Font.Color.RGB=3342528; $run.Font.Bold=-1
   $script:highlightAudit += [pscustomobject]@{Page=$page;Shape=$shape.Name;Text=$term;Start=$at+1;Color='#C00033'}
   $offset=$at+([string]$term).Length
  }
 }
}
New-Item -ItemType Directory -Path (Join-Path $w 'renders') -Force|Out-Null
try{
 $deck=$app.Presentations.Open((Join-Path $w 'base.pptx'),0,0,0)
 # Copy the user's actual native TOC shapes; preserve the target slide and its notes.
 for($pi=0;$pi -lt $content.pages.Count;$pi++){
  $q=$content.pages[$pi]
  if($q.tocEntries){
   $items=@($q.tocEntries); if($items.Count -lt 3 -or $items.Count -gt 5){throw 'TOC template supports 3-5 chapters; use a second TOC for more.'}
   $ref=$app.Presentations.Open((Join-Path $PSScriptRoot '..\assets\toc_reference.pptx'),0,0,0)
   try{
    $target=$deck.Slides.Item($pi+1)
    for($j=$target.Shapes.Count;$j -ge 1;$j--){$target.Shapes.Item($j).Delete()}
    $target.DisplayMasterShapes=0;$target.FollowMasterBackground=0;$target.Background.Fill.ForeColor.RGB=16777215
    $ref.Slides.Item(1).Shapes.Range().Copy();$null=$target.Shapes.Paste()
    $nums=@($target.Shapes|Where-Object {$_.HasTextFrame -eq -1 -and $_.TextFrame.HasText -eq -1 -and $_.TextFrame.TextRange.Text -match '^[1-5]$'}|Sort-Object Top)
    $labels=@($target.Shapes|Where-Object {$_.HasTextFrame -eq -1 -and $_.TextFrame.HasText -eq -1 -and $_.Left -gt 370}|Sort-Object Top)
    for($j=4;$j -ge $items.Count;$j--){$nums[$j].Delete();$labels[$j].Delete()}
    for($j=0;$j -lt $items.Count;$j++){
     $yy=106+$j*(274/[Math]::Max(1,$items.Count-1))
     $nums[$j].Top=[double]$yy;$labels[$j].Top=[double]$yy
     $labels[$j].TextFrame.TextRange.Text=[string]$items[$j]
     $labels[$j].TextFrame.TextRange.Font.Bold=-1
    }
    # Institutional identity is not inherited from a visual reference.
    for($j=$target.Shapes.Count;$j -ge 1;$j--){if($target.Shapes.Item($j).Type -eq 13){$target.Shapes.Item($j).Delete()}}
   }finally{$ref.Close()}
  }
 }
 foreach($slide in $deck.Slides){
  $terms=@($content.pages[$slide.SlideIndex-1].emphasis|Where-Object {$_})
  $slide.SlideShowTransition.EntryEffect=1793;$slide.SlideShowTransition.Duration=.35
  foreach($z in $slide.Shapes){
   if($z.HasTextFrame -eq -1 -and $z.TextFrame2.HasText -eq -1){
    $tf=$z.TextFrame2;$tf.AutoSize=0;$tf.WordWrap=-1
    $tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
    $tf.TextRange.Font.Name='Microsoft YaHei';$tf.TextRange.Font.NameFarEast='Microsoft YaHei'
    $tf.TextRange.ParagraphFormat.SpaceWithin=1.08
    $tf.TextRange.ParagraphFormat.SpaceBefore=0;$tf.TextRange.ParagraphFormat.SpaceAfter=0
    Set-Emphasis $z $terms $slide.SlideIndex
    $t=$tf.TextRange;$audit += [pscustomobject]@{Slide=$slide.SlideIndex;Name=$z.Name;Text=$t.Text;Size=$t.Font.Size;X=$z.Left;Y=$z.Top;W=$z.Width;H=$z.Height;BW=$t.BoundWidth;BH=$t.BoundHeight;Align=$t.ParagraphFormat.Alignment}
   }
   if($z.HasTable -eq -1){
    for($i=1;$i -le $z.Table.Rows.Count;$i++){for($j=1;$j -le $z.Table.Columns.Count;$j++){
     $cell=$z.Table.Cell($i,$j).Shape;$tf=$cell.TextFrame2
     $tf.MarginLeft=9;$tf.MarginRight=9;$tf.MarginTop=6;$tf.MarginBottom=6
     $tf.TextRange.Font.Name='Microsoft YaHei';$tf.TextRange.Font.NameFarEast='Microsoft YaHei'
     $tf.TextRange.ParagraphFormat.SpaceWithin=1.05
     Set-Emphasis $cell $terms $slide.SlideIndex
     $t=$tf.TextRange;$tableAudit += [pscustomobject]@{Slide=$slide.SlideIndex;Row=$i;Col=$j;Text=$t.Text;Size=$t.Font.Size;H=$cell.Height;BH=$t.BoundHeight;W=$cell.Width;BW=$t.BoundWidth}
    }}
   }
   if($z.HasChart -eq -1){
    $ch=$z.Chart;$ch.ChartArea.Font.Name='Microsoft YaHei';$ch.ChartArea.Font.Size=18
    for($a=1;$a -le 2;$a++){
     try{$axis=$ch.Axes($a);$axis.TickLabels.Font.Name='Microsoft YaHei';$axis.TickLabels.Font.Size=18;if($axis.HasTitle){$axis.AxisTitle.Font.Name='Microsoft YaHei';$axis.AxisTitle.Font.Size=18}}catch{}
    }
    foreach($ser in $ch.SeriesCollection()){if($ser.HasDataLabels){$ser.DataLabels().Font.Name='Microsoft YaHei';$ser.DataLabels().Font.Size=18}}
    $chartAudit += [pscustomobject]@{Slide=$slide.SlideIndex;Series=$ch.SeriesCollection().Count;Font=$ch.ChartArea.Font.Size;Editable=$true}
   }
  }
  $slide.Export((Join-Path $w ('renders\slide-'+('{0:D2}' -f $slide.SlideIndex)+'.png')),'PNG',1600,900)
 }
 $deck.SaveAs((Join-Path $w 'candidate.pptx'),24)
 $deck.SaveAs((Join-Path $w 'preview.pdf'),32)
 ConvertTo-Json -InputObject @($audit) -Depth 6|Set-Content -LiteralPath (Join-Path $w 'powerpoint_audit.json') -Encoding UTF8
 ConvertTo-Json -InputObject @($tableAudit) -Depth 6|Set-Content -LiteralPath (Join-Path $w 'table_audit.json') -Encoding UTF8
 ConvertTo-Json -InputObject @($chartAudit) -Depth 6|Set-Content -LiteralPath (Join-Path $w 'chart_audit.json') -Encoding UTF8
 ConvertTo-Json -InputObject @($highlightAudit) -Depth 6|Set-Content -LiteralPath (Join-Path $w 'highlight_audit.json') -Encoding UTF8
 Write-Output ('Rendered and reopened authoring deck: '+$deck.Slides.Count+' slides')
}finally{if($deck){$deck.Close()};[Runtime.InteropServices.Marshal]::ReleaseComObject($app)|Out-Null}
