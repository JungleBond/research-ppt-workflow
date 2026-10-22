"""Remove common author metadata from PPTX/PDF; visual and source review remains required."""
from pathlib import Path
from lxml import etree as ET
import argparse, zipfile, posixpath, re

def clean_pptx(source, output, remove_pictures=False):
    with zipfile.ZipFile(source) as z:
        data={n:z.read(n) for n in z.namelist()}
    drop={n for n in data if n.startswith(('docProps/thumbnail','docProps/custom','ppt/comments/','ppt/commentAuthors','ppt/persons/'))}
    if remove_pictures:
        # Reserved for a reviewed template whose only pictures are identity logos.
        for n,b in list(data.items()):
            if re.match(r'ppt/(slides|slideMasters|slideLayouts)/[^/]+\.xml$',n):
                root=ET.fromstring(b)
                for parent in root.iter():
                    for child in list(parent):
                        if child.tag=='{http://schemas.openxmlformats.org/presentationml/2006/main}pic':parent.remove(child)
                data[n]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
        drop.update(n for n in data if n.startswith('ppt/media/'))
    core='docProps/core.xml'
    if core in data:
        root=ET.fromstring(data[core])
        for child in list(root):
            if child.tag.rsplit('}',1)[-1] in ['creator','lastModifiedBy','description','keywords','subject','category','contentStatus','lastPrinted']:
                root.remove(child)
        data[core]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    if 'docProps/app.xml' in data:
        root=ET.fromstring(data['docProps/app.xml'])
        for child in list(root):
            if child.tag.rsplit('}',1)[-1] in ['Company','Manager','HyperlinkBase']:root.remove(child)
        data['docProps/app.xml']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    for n,b in list(data.items()):
        if not n.endswith('.rels'):continue
        root=ET.fromstring(b); base=posixpath.dirname(posixpath.dirname(n)) if n!='_rels/.rels' else ''
        changed=False
        for child in list(root):
            target=posixpath.normpath(posixpath.join(base,child.get('Target',''))).lstrip('/')
            if target in drop:
                root.remove(child);changed=True
        if changed:data[n]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    root=ET.fromstring(data['[Content_Types].xml'])
    for child in list(root):
        if child.get('PartName','').lstrip('/') in drop:root.remove(child)
    data['[Content_Types].xml']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    output=Path(output);output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in data.items():
            if n not in drop:z.writestr(n,b)

def clean_pdf(source,output):
    from pypdf import PdfReader,PdfWriter
    reader=PdfReader(source);writer=PdfWriter(clone_from=reader)
    writer.metadata=None
    if '/Metadata' in writer._root_object:del writer._root_object['/Metadata']
    writer.write(output)

if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('input');p.add_argument('output');a=p.parse_args()
    if Path(a.input).resolve()==Path(a.output).resolve():p.error('Use a separate output file and retain the original.')
    if Path(a.input).suffix.lower()=='.pptx':clean_pptx(a.input,a.output)
    elif Path(a.input).suffix.lower()=='.pdf':clean_pdf(a.input,a.output)
    else:p.error('Only PPTX/PDF are supported.')
