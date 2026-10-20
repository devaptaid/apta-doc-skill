#!/usr/bin/env python3
"""Build Aptaworks presentations: Markdown -> editable HTML -> native PPTX."""
from __future__ import annotations

import argparse
import html
import json
import re
import struct
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote
from xml.sax.saxutils import escape

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
EMU = 9525  # one CSS pixel at 96 DPI
W, H = 1280, 720
NAVY, DEEP, ORANGE, PAPER, RULE, MUTED = '17234E', '0B1533', 'EF7636', 'F6F3EA', 'C9D2DF', '5E6B82'
LAYOUTS = {'cover', 'section', 'ledger', 'process', 'flow', 'timeline', 'kpi', 'cards', 'proof', 'quote', 'columns', 'table', 'image', 'closing'}
ASSETS = Path(__file__).resolve().parent.parent / 'assets'
FONT_METRICS = json.loads((ASSETS / 'font-metrics.json').read_text(encoding='utf-8'))


def xml(body):
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + body


def rels(entries):
    return xml(f'<Relationships xmlns="{PKG}">' + ''.join(
        f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{html.escape(target, quote=True)}"/>'
        for rid, kind, target in entries) + '</Relationships>')


def paragraph(text, size=24, color=NAVY, bold=False, align='l', font='Lato', italic=False, opacity=100, line_height=1.15, language='en'):
    lang = 'en-US' if language == 'en' else 'id-ID'
    return (f'<a:p><a:pPr algn="{align}"><a:lnSpc><a:spcPct val="{int(line_height*100000)}"/></a:lnSpc></a:pPr>'
            f'<a:r><a:rPr lang="{lang}" sz="{int(size * 75)}" b="{int(bold)}" i="{int(italic)}">'
            f'<a:solidFill><a:srgbClr val="{color}"><a:alpha val="{int(opacity*1000)}"/></a:srgbClr></a:solidFill><a:latin typeface="{html.escape(font,quote=True)}"/>'
            f'</a:rPr><a:t xml:space="preserve">{escape(text)}</a:t></a:r>'
            f'<a:endParaRPr lang="{lang}" sz="{int(size * 75)}"/></a:p>')


def wrap(text, width, size, max_lines, font='Lato', bold=False, italic=False):
    # Packaged glyph advances make the density check portable; native rendering
    # remains authoritative for kerning, line breaks, and substituted fonts.
    metric_key = 'Georgia Italic' if font == 'Georgia' else 'Lato Bold' if font == 'Lato' and bold else font
    metrics = FONT_METRICS.get(metric_key, FONT_METRICS['Lato'])
    def measure(value):
        return sum(metrics.get(c, .58) for c in value)*size
    lines = 0
    for raw in text.split('\n'):
        current = ''
        lines += 1
        for word in raw.split():
            if measure(word) > width:
                raise ValueError(f'Unbreakable text exceeds layout width: {word[:70]!r}')
            candidate = (current+' '+word).strip()
            if current and measure(candidate) > width:
                lines += 1
                current = word
            else:
                current = candidate
    if lines > max_lines:
        raise ValueError(f'Text exceeds layout capacity ({lines}>{max_lines} lines): {text[:80]!r}. Split or shorten it.')
    return lines


def read_source(path):
    raw = path.read_text(encoding='utf-8-sig')
    meta = {'language': 'en', 'title': 'Aptaworks', 'subtitle': '', 'client': '', 'date': '', 'author': 'Aptaworks Team'}
    if raw.startswith('---\n'):
        parts = raw.split('\n---\n', 1)
        if len(parts) != 2:
            raise ValueError('Unclosed frontmatter.')
        for line in parts[0][4:].splitlines():
            if not line.strip() or line.lstrip().startswith('#'):
                continue
            if ':' not in line:
                raise ValueError(f'Invalid metadata: {line}')
            key, value = line.split(':', 1)
            if key.strip() not in meta:
                raise ValueError(f'Unknown metadata key: {key.strip()}')
            if line != line.lstrip() or value.strip().startswith(('[','{','|','>')):
                raise ValueError('Frontmatter supports single-line scalar values only.')
            meta[key.strip()] = value.strip().strip('"\'')
        raw = parts[1]
    if meta['language'] not in {'en', 'id'}:
        raise ValueError("language must be 'en' or 'id'.")
    if 'AptaWorks' in raw or 'AptaWorks' in str(meta) or 'Aptaworks Consultant' in str(meta):
        raise ValueError('Use Aptaworks / Aptaworks Team branding.')
    chunks = re.split(r'^## (.+?) \{([a-z]+)\}\s*$', raw, flags=re.M)
    if chunks[0].strip() or len(chunks) == 1:
        raise ValueError('Slides must start with: ## Slide title {layout}')
    slides = []
    for index in range(1, len(chunks), 3):
        title, layout, body = chunks[index:index + 3]
        if layout not in LAYOUTS:
            raise ValueError(f'Unsupported layout: {layout}')
        notes = re.findall(r'<!-- notes:\s*(.*?)-->', body, flags=re.S)
        body = re.sub(r'<!-- notes:.*?-->', '', body, flags=re.S).strip()
        lines = [line.strip() for line in body.splitlines() if line.strip()]
        for line in lines:
            check = re.sub(r'!\[[^\]]+\]\([^)]+\)', '', line) if layout == 'image' else line
            if re.search(r'(^#{1,6}\s|```|~~~|\*\*|__|`|<[^>]+>|\[[^\]]+\]\([^)]+\))', check):
                raise ValueError(f'Unsupported rich Markdown or HTML on slide {title!r}: {line}')
        if not title.strip():
            raise ValueError('Slide title cannot be empty.')
        slides.append({'title': title.strip(), 'layout': layout, 'lines': lines, 'notes': '\n'.join(notes).strip()})
    return meta, slides


class Slide:
    def __init__(self, layout, number, total, meta, furniture=True):
        self.shapes, self.media, self.next_id = [], [], 2
        self.elements = []
        self.layout = layout
        self.language = meta['language']
        self.dark = layout in {'cover', 'section', 'closing'}
        self.color = 'FFFFFF' if self.dark else NAVY
        self.bg = DEEP if self.dark else PAPER
        if layout == 'closing':
            stops = [(0,'182B63'),(48000,'152B5B'),(100000,'16A6C8')]
            self.background = {'stops': stops, 'angle': 5400000}
            self.bg_fill = self.gradient(stops,5400000)
        elif self.dark:
            self.background = {'stops': [(0,'124C6E'),(46000,NAVY),(100000,DEEP)], 'angle': 2700000}
            self.bg_fill = self.gradient([(0,'124C6E'),(46000,NAVY),(100000,DEEP)],2700000)
        else:
            self.background = {'color': PAPER}
            self.bg_fill = f'<a:solidFill><a:srgbClr val="{PAPER}"/></a:solidFill>'
        # The user requested clean presentation backgrounds without a paper grid.
        if not furniture:
            return
        self.text(f'{number:02d} / {total:02d}',1080,656,116,24,12,NAVY if layout == 'closing' else self.color,align='r',font='Consolas')
        if layout != 'closing':
            self.rect(84,688,1112,1,'FFFFFF' if self.dark else NAVY,opacity=46 if self.dark else 55)
        if not self.dark:
            self.rect(84,29,34,3,ORANGE)
            self.text(layout.upper()+' / COMPONENT LIBRARY',128,22,640,20,11,MUTED,font='Consolas')
            self.text('Aptaworks',855,22,341,20,12,MUTED,align='r',font='Consolas')

    @staticmethod
    def gradient(stops, angle):
        return '<a:gradFill rotWithShape="1"><a:gsLst>'+''.join(f'<a:gs pos="{pos}"><a:srgbClr val="{color}"/></a:gs>' for pos,color in stops)+f'</a:gsLst><a:lin ang="{angle}" scaled="1"/></a:gradFill>'

    def ident(self, name, description=''):
        value = self.next_id
        self.next_id += 1
        return f'<p:cNvPr id="{value}" name="{html.escape(name, quote=True)}" descr="{html.escape(description, quote=True)}"/>'

    def pos(self, x, y, w, h):
        if x < 0 or y < 0 or x + w > W + 1 or y + h > H + 1:
            raise ValueError(f'Shape outside slide: {x, y, w, h}')
        return f'<a:xfrm><a:off x="{int(x*EMU)}" y="{int(y*EMU)}"/><a:ext cx="{int(w*EMU)}" cy="{int(h*EMU)}"/></a:xfrm>'

    def rect(self, x, y, w, h, color, opacity=100, stroke=None, preset='rect', name='Rule'):
        self.elements.append(dict(kind='shape',x=x,y=y,w=w,h=h,color=color,opacity=opacity,stroke=stroke,preset=preset,name=name))
        fill = f'<a:solidFill><a:srgbClr val="{color}"><a:alpha val="{int(opacity*1000)}"/></a:srgbClr></a:solidFill>' if color else '<a:noFill/>'
        line = f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{stroke}"/></a:solidFill><a:prstDash val="solid"/></a:ln>' if stroke else '<a:ln><a:noFill/></a:ln>'
        adjust = '<a:gd name="adj" fmla="val 8000"/>' if preset == 'roundRect' else ''
        self.shapes.append(f'<p:sp><p:nvSpPr>{self.ident(name)}<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                           f'<p:spPr>{self.pos(x,y,w,h)}<a:prstGeom prst="{preset}"><a:avLst>{adjust}</a:avLst></a:prstGeom>'
                           f'{fill}{line}</p:spPr></p:sp>')

    def text(self, text, x, y, w, h, size=24, color=None, bold=False, align='l', font='Lato', italic=False, opacity=100, line_height=1.15, role=None):
        wrap(text, w, size, max(1, int(h / (size * line_height))),font,bold,italic)
        self.elements.append(dict(kind='text',text=text,x=x,y=y,w=w,h=h,size=size,color=color or self.color,bold=bold,align=align,font=font,italic=italic,opacity=opacity,line_height=line_height,role=role))
        self.shapes.append(f'<p:sp><p:nvSpPr>{self.ident("Text: " + text[:50])}<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
                           f'<p:spPr>{self.pos(x,y,w,h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>'
                           '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
                           + ''.join(paragraph(t,size,color or self.color,bold,align,font,italic,opacity,line_height,self.language) for t in text.split('\n')) + '</p:txBody></p:sp>')

    def image(self, path, x, y, w, h, description='Aptaworks brand asset'):
        data = path.read_bytes()
        if data[:8] != b'\x89PNG\r\n\x1a\n':
            raise ValueError(f'Only PNG images supported; convert this asset first: {path}')
        iw, ih = struct.unpack('>II', data[16:24])
        scale = min(w / iw, h / ih)
        pw, ph = iw * scale, ih * scale
        x, y = x + (w - pw)/2, y + (h - ph)/2
        self.elements.append(dict(kind='image',path=path,x=x,y=y,w=pw,h=ph,description=description))
        rid = f'rIdImage{len(self.media)+1}'
        self.media.append((rid, path))
        self.shapes.append(f'<p:pic><p:nvPicPr>{self.ident(path.name,description)}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
                           f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                           f'<p:spPr>{self.pos(x,y,pw,ph)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')

    def table(self, rows, x=84, y=300, width=1112, height=None):
        count = len(rows[0])
        if not 2 <= count <= 4 or not 2 <= len(rows) <= 6 or any(len(row) != count for row in rows):
            raise ValueError('Table needs 2–4 columns, 2–6 rows including header, and equal columns in every row.')
        height = min(336, 56 * len(rows)) if height is None else height
        self.pos(x,y,width,height)
        self.elements.append(dict(kind='table',rows=rows,x=x,y=y,w=width,h=height))
        col_width, row_height = width/count, height/len(rows)
        cells = []
        for ri, row in enumerate(rows):
            row_xml = ''
            for cell in row:
                wrap(cell, col_width-24, 23, max(1, int((row_height-16)/27)))
                color, fill = ('FFFFFF', NAVY) if ri == 0 else (NAVY, 'FFFFFF' if ri % 2 else 'F0EFEA')
                borders = ''.join(f'<a:ln{edge} w="9525"><a:solidFill><a:srgbClr val="{RULE}"/></a:solidFill></a:ln{edge}>' for edge in ('L','R','T','B'))
                row_xml += ('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>' + paragraph(cell,23,color,ri==0,language=self.language)
                            + f'</a:txBody><a:tcPr marL="114300" marR="114300" marT="76200" marB="76200">'
                            + borders +
                            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></a:tcPr></a:tc>')
            cells.append(f'<a:tr h="{int(row_height*EMU)}">{row_xml}</a:tr>')
        self.shapes.append(f'<p:graphicFrame><p:nvGraphicFramePr>{self.ident("Editable table")}<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
                           f'<p:xfrm><a:off x="{int(x*EMU)}" y="{int(y*EMU)}"/><a:ext cx="{int(width*EMU)}" cy="{int(height*EMU)}"/></p:xfrm>'
                           '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/>'
                           '<a:tblGrid>' + f'<a:gridCol w="{int(col_width*EMU)}"/>'*count + '</a:tblGrid>'
                           + ''.join(cells) + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')

    def export(self):
        return xml(f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld name="{self.layout}">'
                   f'<p:bg><p:bgPr>{self.bg_fill}<a:effectLst/></p:bgPr></p:bg>'
                   '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                   '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
                   + ''.join(self.shapes) + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def clean_lines(lines):
    values = []
    for line in lines:
        if not line.startswith('- '):
            raise ValueError(f'Expected a plain Markdown list item, got: {line}')
        values.append(line[2:])
    return values


def compose(item, number, total, meta, source_dir):
    from layouts import compose as compose_layout
    return compose_layout(item, number, total, meta, source_dir)


def export_pptx(meta, items, slides, output):
    # Include the complete minimal master/theme structure required by PowerPoint.
    group = '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree>'
    clrmap = '<p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>'
    parts = {}
    parts['_rels/.rels'] = rels([('rId1','officeDocument','ppt/presentation.xml')])
    slide_ids = ''.join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(len(slides)))
    parts['ppt/presentation.xml'] = xml(f'<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>{slide_ids}</p:sldIdLst><p:sldSz cx="{W*EMU}" cy="{H*EMU}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
    parts['ppt/_rels/presentation.xml.rels'] = rels([('rId1','slideMaster','slideMasters/slideMaster1.xml')] + [(f'rId{i+2}','slide',f'slides/slide{i+1}.xml') for i in range(len(slides))])
    parts['ppt/slideMasters/slideMaster1.xml'] = xml(f'<p:sldMaster xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld>{group}</p:cSld>{clrmap}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>')
    parts['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = rels([('rId1','slideLayout','../slideLayouts/slideLayout1.xml'),('rId2','theme','../theme/theme1.xml')])
    parts['ppt/slideLayouts/slideLayout1.xml'] = xml(f'<p:sldLayout xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" type="blank" preserve="1"><p:cSld name="Blank">{group}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
    parts['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = rels([('rId1','slideMaster','../slideMasters/slideMaster1.xml')])
    colors = dict(dk1=NAVY,lt1='FFFFFF',dk2=DEEP,lt2=PAPER,accent1=NAVY,accent2=ORANGE,accent3='0A6F93',accent4='2448FF',accent5=MUTED,accent6=RULE,hlink='2448FF',folHlink=MUTED)
    fills = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'*3
    line = '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
    parts['ppt/theme/theme1.xml'] = xml(f'<a:theme xmlns:a="{A}" name="Aptaworks"><a:themeElements><a:clrScheme name="Aptaworks">'+''.join(f'<a:{key}><a:srgbClr val="{value}"/></a:{key}>' for key,value in colors.items())+'</a:clrScheme><a:fontScheme name="Aptaworks"><a:majorFont><a:latin typeface="Lato"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Lato"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Aptaworks"><a:fillStyleLst>'+fills+'</a:fillStyleLst><a:lnStyleLst>'+line*3+'</a:lnStyleLst><a:effectStyleLst>'+ '<a:effectStyle><a:effectLst/></a:effectStyle>'*3 + '</a:effectStyleLst><a:bgFillStyleLst>'+fills+'</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>')
    types = {'/ppt/presentation.xml':'presentationml.presentation.main','/ppt/slideMasters/slideMaster1.xml':'presentationml.slideMaster','/ppt/slideLayouts/slideLayout1.xml':'presentationml.slideLayout','/ppt/theme/theme1.xml':'theme'}
    for i,(slide,item) in enumerate(zip(slides,items),1):
        parts[f'ppt/slides/slide{i}.xml'] = slide.export()
        relationships = [('rIdLayout','slideLayout','../slideLayouts/slideLayout1.xml')]
        for j,(rid,path) in enumerate(slide.media,1):
            name = f'image{i}_{j}.png'
            parts[f'ppt/media/{name}'] = path.read_bytes()
            relationships.append((rid,'image','../media/'+name))
        if item['notes']:
            relationships.append(('rIdNotes','notesSlide',f'../notesSlides/notesSlide{i}.xml'))
            parts[f'ppt/notesSlides/notesSlide{i}.xml'] = xml(f'<p:notes xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'+paragraph(item['notes'])+'</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>')
            parts[f'ppt/notesSlides/_rels/notesSlide{i}.xml.rels'] = rels([('rId1','slide',f'../slides/slide{i}.xml')])
            types[f'/ppt/notesSlides/notesSlide{i}.xml'] = 'presentationml.notesSlide'
        parts[f'ppt/slides/_rels/slide{i}.xml.rels'] = rels(relationships)
        types[f'/ppt/slides/slide{i}.xml'] = 'presentationml.slide'
    parts['[Content_Types].xml'] = xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'+''.join(f'<Override PartName="{name}" ContentType="application/vnd.openxmlformats-officedocument.{kind}+xml"/>' for name,kind in types.items())+'</Types>')
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as archive:
        for name,value in parts.items():
            archive.writestr(name,value.encode('utf-8') if isinstance(value,str) else value)
    return len(slides)


def build(source, output):
    from deck_html import compile_html, export_html
    if source.suffix.lower() == '.html':
        if output.suffix.lower() != '.pptx':
            raise ValueError('HTML input must be exported to .pptx.')
        return export_html(source,output)
    if source.suffix.lower() != '.md':
        raise ValueError('Input must be a Markdown .md or generated presentation .html file.')
    html_path = output if output.suffix.lower() == '.html' else output.with_suffix('.html')
    count = compile_html(source,html_path)
    if output.suffix.lower() == '.pptx':
        return export_html(html_path,output)
    return count


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path,help='Markdown source or generated presentation HTML')
    parser.add_argument('--output',type=Path,required=True)
    args = parser.parse_args()
    try:
        if args.output.suffix.lower() not in {'.pptx','.html'}:
            raise ValueError('Output must end in .html or .pptx')
        count = build(args.source.resolve(),args.output.resolve())
        print(f'{args.output.suffix[1:].upper()}: {args.output} ({count} slides)')
        if args.source.suffix.lower() == '.md' and args.output.suffix.lower() == '.pptx':
            print(f'HTML: {args.output.with_suffix(".html")}')
        return 0
    except (OSError, ValueError, KeyError, struct.error) as exc:
        print(f'ERROR: {exc}',file=sys.stderr)
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
