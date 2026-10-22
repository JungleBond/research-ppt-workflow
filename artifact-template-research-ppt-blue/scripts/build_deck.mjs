import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const req=createRequire(path.join(process.env.RUNTIME_NODE_MODULES||path.join(root,'node_modules'),'package.json'));
const {PresentationFile,FileBlob}=req('@oai/artifact-tool');
const sharp=req('sharp');
const [dataArg,outArg]=process.argv.slice(2);
if(!dataArg||!outArg)throw Error('Usage: node build_deck.mjs deck.json build-directory');
const input=path.resolve(dataArg),out=path.resolve(outArg),base=path.dirname(input);
const D=JSON.parse(await fs.readFile(input,'utf8'));
await fs.mkdir(out,{recursive:true});
if(!D.pages?.length||!D.chapters?.length)throw Error('pages and chapters required');
const C={blue:'#005D97',dark:'#004675',cyan:'#00B0F0',ink:'#172D3C',muted:'#526674',pale:'#DAE3F6',mint:'#E4F3DA',peach:'#FCE6D5',white:'#FFFFFF'};
const p=await PresentationFile.importPptx(await FileBlob.load(path.join(root,'assets/template_clean.pptx')));
const layout=p.slides.items[0].layout;for(const x of [...p.slides.items])x.delete();
const manifest=[],tableOwners=[],chartOwners=[];let s,visible,n;
function rect(x,y,w,h,fill,line='none'){return s.shapes.add({geometry:'rect',position:{left:x,top:y,width:w,height:h},fill,line:{fill:line,width:line==='none'?0:1}})}
function text(t,x,y,w,h,size=27,color=C.ink,bold=false,align='left',valign='top',name=''){
 if(size<24)throw Error(`Page ${n}: font ${size}px is below 18pt`);
 visible.push(t);const z=s.shapes.add({geometry:'textbox',name:name||`text_${n}_${visible.length}`,position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 z.text=t;z.text.style={typeface:'Microsoft YaHei',fontSize:size,color,bold,alignment:align,verticalAlignment:valign,autoFit:'none',wrap:'square',insets:0,lineSpacing:1.08};return z;
}
for(const [i,q]of D.pages.entries()){
 n=i+1;visible=[];s=p.slides.add();s.setLayout(layout);s.background.fill=C.white;
 for(const field of ['title','question','previous','next','speech','cue','sources'])if(q[field]===undefined)throw Error(`Page ${n} missing ${field}`);
 if(!q.fullBleed){
  rect(0,0,1280,57,'#F2F2F2');
  const cw=1232/D.chapters.length;
  D.chapters.forEach((ch,j)=>{if(j===q.chapter)rect(24+j*cw,0,cw,57,C.white);text(ch,38+j*cw,0,cw-24,57,24,j===q.chapter?C.ink:'#999999',j===q.chapter,'left','middle')});
  rect(33,96,11,16,C.cyan);text(q.title,57,82,1187,64,38,C.blue,true,'left','middle','title');
  text(D.shortTitle||D.title,28,684,1050,36,24,C.white,false,'left','middle');
  text(String(n).padStart(2,'0'),1190,684,60,36,24,C.white,false,'left','middle');
  if(q.shortSource)text(q.shortSource,36,650,1208,30,24,C.muted);
 }
 for(const [k,o]of (q.objects||[]).entries()){
  const {x,y,w,h}=o;
  if([x,y,w,h].some(v=>!Number.isFinite(v))||x<0||y<0||x+w>1280.5||y+h>720.5)throw Error(`Page ${n}, object ${k}: invalid bounds`);
  if(o.type==='rect')rect(x,y,w,h,o.fill||C.pale,o.line);
  else if(o.type==='text')text(o.text,x,y,w,h,o.size||27,o.color||C.ink,!!o.bold,o.align||'left',o.valign||'top',o.name);
  else if(o.type==='image'){
   if(!o.role||!o.source)throw Error(`Page ${n}: image role and source required`);
   const f=path.resolve(base,o.file);const blob=await fs.readFile(f);
   if(o.crop){
    const edges=['left','top','right','bottom'];
    if(edges.some(e=>!Number.isFinite(o.crop[e])||o.crop[e]<0||o.crop[e]>=1)||o.crop.left+o.crop.right>=1||o.crop.top+o.crop.bottom>=1)throw Error(`Page ${n}: invalid normalized crop`);
   }
   const metadata=await sharp(blob).metadata();
   s.images.add({blob,contentType:metadata.format==='jpeg'?'image/jpeg':'image/png',alt:o.alt||o.role,fit:o.fit||'contain',...(o.crop?{crop:o.crop}:{}),position:{left:x,top:y,width:w,height:h}});
   manifest.push({page:n,file:o.file,source:o.source,role:o.role,sourceRegion:o.sourceRegion||null,crop:o.crop||null,x,y,w,h});
  }else if(o.type==='table'){
   visible.push(...o.values.flat());const tb=s.tables.add({rows:o.values.length,columns:o.values[0].length,left:x,top:y,width:w,height:h,columnWidths:o.widths,values:o.values});
   tb.borders.assign({style:'solid',fill:'#CDDEE8',width:1});const f=o.size||26;if(f<24)throw Error('Table font below 18pt');
   const colAlign=o.values[0].map((_,c)=>o.alignments?.[c]||(o.values.slice(1).every(row=>/^\s*[+−-]?\d[\d., /()%×]*\s*(?:[KMG]|FPS|ms|秒)?\s*$/.test(String(row[c])))?'center':'left'));
   for(let r=0;r<o.values.length;r++)for(let c=0;c<o.values[0].length;c++){
    const cell=tb.getCell(r,c);cell.fill=r===0?C.blue:(r%2?C.white:'#EEF4F8');
    cell.text.style={typeface:'Microsoft YaHei',fontSize:f,color:r===0?C.white:C.ink,bold:r===0||c===0,alignment:colAlign[c],verticalAlignment:'middle',autoFit:'none',wrap:'square',insets:12};
   }tableOwners.push(n);
  }else if(o.type==='chart'){
   if(!o.source||!o.config?.series)throw Error('Chart source and literal data required');
   s.charts.add(o.chartType||'bar',{...o.config,position:{left:x,top:y,width:w,height:h}});
   chartOwners.push(n);manifest.push({page:n,kind:'native-chart',source:o.source,role:o.role,values:o.config.series,categories:o.config.categories});
  }else if(o.type==='thanks'){
   rect(0,0,1280,720,C.white);
   const shadow=await sharp(path.join(root,'assets/template_shadow.png')).rotate(90).resize(1096,28,{fit:'fill'}).png().toBuffer();
   for(const yy of [190,510])s.images.add({blob:shadow,contentType:'image/png',alt:'模板横带阴影',fit:'contain',position:{left:92,top:yy,width:1096,height:28}});
   rect(0,204,1280,320,C.blue);text('谢谢大家',240,292,800,144,76,C.white,true,'center','middle','thank_you');
  }else throw Error('Unsupported object type '+o.type);
 }
 q.visible=visible;
 const refs=q.sources.map(id=>{const v=D.sources.find(z=>z.id===id);if(!v)throw Error(`Missing source ${id}`);return `${v.title}\n${v.url}\n定位：${q.sourceLocator||v.locator||''}`}).join('\n\n');
 s.speakerNotes.text=`第 ${n} 页：${q.title}\n预计 ${q.minutes} 分钟\n本页问题：${q.question}\n承接上页：${q.previous}\n接下来：${q.next}\n\n【完整口播】\n${q.speech}\n\n【操作提示】\n${q.cue}\n\n【画面文字】\n${visible.join('\n')}\n\n【资料与来源】\n${refs}`;
}
await fs.writeFile(path.join(out,'deck.resolved.json'),JSON.stringify(D,null,2));
await fs.writeFile(path.join(out,'visual_manifest.json'),JSON.stringify(manifest,null,2));
await fs.writeFile(path.join(out,'requirements.json'),JSON.stringify({explicitTotalSlideCount:D.pages.length,requiredNativeTableOwnerSlides:[...new Set(tableOwners)],requiredNativeChartOwnerSlides:[...new Set(chartOwners)],noVideo:D.noVideo!==false},null,2));
await(await PresentationFile.exportPptx(p)).save(path.join(out,'base.pptx'));
console.log(JSON.stringify({slides:D.pages.length,tables:tableOwners,charts:chartOwners,output:path.join(out,'base.pptx')}));
