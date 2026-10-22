from pathlib import Path
import sys,json,zipfile,re,hashlib,base64,io,html
from collections import defaultdict
import xml.etree.ElementTree as ET
from PIL import Image,ImageDraw,ImageFont

w=Path(sys.argv[1]).resolve(); final=Path(sys.argv[2]).resolve() if len(sys.argv)>2 else w/'candidate.pptx'
D=json.loads((w/'deck.resolved.json').read_text('utf8'));R=json.loads((w/'requirements.json').read_text('utf8'))
read=lambda name:json.loads((w/name).read_text('utf-8-sig'))
a=read('powerpoint_audit.json');ta=read('table_audit.json')
bad=[x for x in a if x['Size']<18 or x['BH']>x['H']+1 or x['BW']>x['W']+2]
tb=[x for x in ta if x['Size']<18 or x['BH']>x['H']-8 or x['BW']>x['W']-13]
ns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
issues=[];fonts=[];tables=[];charts=[]
with zipfile.ZipFile(final) as z:
    slides=[n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)]
    if len(slides)!=len(D['pages']):issues.append('slide count mismatch')
    media=[n for n in z.namelist() if n.startswith('ppt/media/') and n.lower().endswith(('.mp4','.mov','.avi','.wmv','.mp3','.wav','.m4a'))]
    if R.get('noVideo') and media:issues.append('unexpected audio/video')
    for i,p in enumerate(D['pages'],1):
        x=ET.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
        fonts += [int(a.attrib['sz'])/100 for a in x.findall('.//a:rPr',ns) if 'sz'in a.attrib]
        if x.findall('.//a:tbl',ns):tables.append(i)
        if b'/chart' in z.read(f'ppt/slides/_rels/slide{i}.xml.rels'):charts.append(i)
        nt=''.join(t.text or '' for t in ET.fromstring(z.read(f'ppt/notesSlides/notesSlide{i}.xml')).findall('.//a:t',ns))
        squash=lambda s:re.sub(r'\s','',s)
        if squash(p['speech']) not in squash(nt):issues.append(f'notes mismatch {i}')
        for sid in p['sources']:
            source=next(s for s in D['sources'] if s['id']==sid)
            if source['url'] not in nt:issues.append(f'missing note source {i}')
    chart_fonts=[]
    for f in z.namelist():
        if re.fullmatch(r'ppt/charts/chart\d+\.xml',f):
            x=ET.fromstring(z.read(f));chart_fonts += [int(t.attrib['sz'])/100 for t in x.iter() if t.tag in [f'{{{ns["a"]}}}rPr',f'{{{ns["a"]}}}defRPr',f'{{{ns["a"]}}}endParaRPr'] and 'sz'in t.attrib]
if min(fonts)<18:issues.append('native font below 18')
if chart_fonts and min(chart_fonts)<18:issues.append('chart font below 18')
if tables!=R['requiredNativeTableOwnerSlides']:issues.append('native table mismatch')
if charts!=R['requiredNativeChartOwnerSlides']:issues.append('native chart mismatch')
report={'slides':len(D['pages']),'minimum_native_pt':min(fonts),'minimum_chart_pt':min(chart_fonts) if chart_fonts else None,'native_tables':tables,'native_charts':charts,'media_count':len(media),'text_issues':bad,'table_issues':tb,'package_issues':issues,'notes_match':not any('notes' in s for s in issues)}
# A broad left-aligned frame can pass overflow checks and still look unbalanced.
# These are review candidates, not a mandate to add filler or center all paragraphs.
balance=[x for x in a if x.get('Align')==1 and x['W']>=330 and 110<x['Y']<480 and x['BW']/max(x['W'],1)<.58 and x['H']>=45]
report['balance_review_candidates']=balance
# Flag duplicated conditions and spacious prose tables for human review.
# These are not failures: a concise numeric comparison can legitimately be sparse.
table_candidates=[]
for number,page in enumerate(D['pages'],1):
    for index,o in enumerate(page.get('objects',[])):
        if o.get('type')!='table':continue
        rows=o['values'][1:]
        reasons=[]
        norm=lambda value:re.sub(r'\s','',str(value))
        if len(rows)>=3:
            for column in range(1,len(o['values'][0])):
                values=[norm(row[column]) for row in rows]
                if len(set(values))==1 and len(values[0])>=5 and re.search(r'[A-Za-z\u4e00-\u9fff]',values[0]):
                    reasons.append('重复的说明列：'+str(o['values'][0][column]))
        if 2<=len(rows)<=4 and o['w']>=1000 and o['h']>=280:
            row_lengths=[sum(len(norm(v)) for v in row) for row in rows]
            text=''.join(norm(v) for row in rows for v in row)
            letters=len(re.findall(r'[A-Za-z\u4e00-\u9fff]',text))
            if sum(row_lengths)/len(rows)<50 and letters>len(text)*.6:
                reasons.append('宽大文字表格内容偏简短，检查行高、实际比较关系和证据缺口')
        if reasons:table_candidates.append({'page':number,'object_index':index,'reasons':reasons})
report['table_review_candidates']=table_candidates
report['emphasis_runs']=len(read('highlight_audit.json')) if (w/'highlight_audit.json').exists() else 0
(w/'automated_review.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),'utf8')
# Contact sheets support whole-deck scanning, never replace full-size review.
for first in range(1,len(D['pages'])+1,9):
    im=Image.new('RGB',(1600,975),'#dfe8ef')
    for j,n in enumerate(range(first,min(first+9,len(D['pages'])+1))):
        pic=Image.open(w/f'renders/slide-{n:02}.png').convert('RGB');pic.thumbnail((528,297));im.paste(pic,(j%3*534,j//3*325+26))
    im.save(w/f'contact-{(first-1)//9+1}.jpg',quality=90)
E=html.escape
H=['<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>完整讲稿与逐页对应</title><style>body{font:19px/1.85 "Microsoft YaHei",sans-serif;margin:0;background:#edf2f6;color:#172d3c}header{padding:32px 5vw;background:#005d97;color:white}main{max-width:1450px;margin:auto;padding:28px}article{background:white;padding:28px;margin:25px 0}h2,h3{color:#005d97}.pair{display:grid;grid-template-columns:1.05fr 1fr;gap:30px}img{width:100%}aside{background:#e4f3da;padding:15px}table{width:100%;border-collapse:collapse}td,th{border:1px solid #aac4d5;padding:12px;text-align:left}@media(max-width:1000px){.pair{display:block}}</style><header><h1>'+E(D['title'])+'</h1><p>完整讲稿与逐页逻辑 · 无视频案例</p></header><main><h2>逐页逻辑</h2><table><tr><th>页</th><th>本页问题</th><th>下一页为何出现</th></tr>']
for i,p in enumerate(D['pages'],1):H.append(f'<tr><td>{i}　{E(p["title"])}</td><td>{E(p["question"])}</td><td>{E(p["next"])}</td></tr>')
H.append('</table><h2>讲稿与画面</h2>')
txt=[D['title'],'']
for i,p in enumerate(D['pages'],1):
    im=Image.open(w/f'renders/slide-{i:02}.png').convert('RGB');im.thumbnail((1280,720));buf=io.BytesIO();im.save(buf,'JPEG',quality=89)
    b64=base64.b64encode(buf.getvalue()).decode()
    H.append(f'<article><h2>{i:02}　{E(p["title"])}</h2><div class="pair"><div><img src="data:image/jpeg;base64,{b64}" alt="{i}"><aside>{E(p["cue"])}</aside></div><div>')
    H.extend('<p>'+E(t)+'</p>' for t in p['speech'].split('\n') if t)
    H.append('<p>来源定位：'+E(p.get('sourceLocator',''))+'</p></div></div></article>')
    txt += [f'第 {i} 页　{p["title"]}',p['speech'],'操作：'+p['cue'],'']
H.append('<h2>资料</h2>')
for s in D['sources']:H.append('<p><a href="'+E(s['url'])+'">'+E(s['title'])+'</a></p>')
H.append('</main></html>')
(w/'完整讲稿与逐页对应.html').write_text(''.join(H),'utf8')
(w/'完整演讲稿.txt').write_text('\n\n'.join(txt),'utf8')
print(json.dumps(report,ensure_ascii=False))
if issues or bad or tb:sys.exit(2)
