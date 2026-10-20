"""Presentation HTML interchange: visible DOM is the source for native PPTX.

Supports this skill's positioned HTML subset, not arbitrary web pages. No browser,
network access, or hidden copy of the Markdown is used when exporting the HTML.
"""
from __future__ import annotations
import base64
import html
import math
import re
import tempfile
from html.parser import HTMLParser
from pathlib import Path
from build_deck import Slide, W, H, NAVY, RULE, read_source, compose, export_pptx

# Office percentage line spacing is relative to the font's natural line box;
# CSS unitless spacing is relative to em size. Translate the supported fonts.
FONT_LINE_BOX = {'Lato':1.2,'Lato Black':1.2,'Georgia':1.137,'Consolas':1.172}


def esc(value):
    return html.escape(str(value),quote=True)


def box(e):
    return ';'.join(f'{key}:{e[field]:g}px' for key,field in [('left','x'),('top','y'),('width','w'),('height','h')])


def color(value,opacity=100):
    if value is None:
        return 'transparent'
    if opacity==100:
        return '#'+value
    rgb=tuple(int(value[i:i+2],16) for i in (0,2,4))
    return f'rgba({rgb[0]},{rgb[1]},{rgb[2]},{opacity/100:g})'


def background_css(background):
    if 'color' in background:
        return '#'+background['color']
    angle=(background['angle']/60000+90)%360
    stops=', '.join(f'#{value} {pos/1000:g}%' for pos,value in background['stops'])
    return f'linear-gradient({angle:g}deg, {stops})'


def element_html(e,title):
    geometry=box(e)
    if e['kind']=='text':
        align={'l':'left','r':'right','ctr':'center'}[e['align']]
        style=(f"{geometry};font-family:'{e['font']}';font-size:{e['size']:g}px;"
               f"font-weight:{'700' if e['bold'] else '400'};font-style:{'italic' if e['italic'] else 'normal'};"
               f"line-height:{e['line_height']*FONT_LINE_BOX.get(e['font'],1.2):g};text-align:{align};color:{color(e['color'],e['opacity'])}")
        heading=' data-role="title" role="heading" aria-level="1"' if e.get('role')=='title' else ''
        paragraphs=''.join(f'<p>{esc(line)}</p>' for line in e['text'].split('\n'))
        return f'<div class="apta-text"{heading} style="{esc(style)}">{paragraphs}</div>'
    if e['kind']=='shape':
        radius='50%' if e['preset']=='ellipse' else '12px' if e['preset']=='roundRect' else '0'
        border=f"1px solid #{e['stroke']}" if e['stroke'] else 'none'
        style=f"{geometry};background:{color(e['color'],e['opacity'])};border:{border};border-radius:{radius}"
        return f'<div class="apta-shape" data-shape="{e["preset"]}" data-name="{esc(e["name"])}" aria-hidden="true" style="{esc(style)}"></div>'
    if e['kind']=='image':
        encoded=base64.b64encode(e['path'].read_bytes()).decode('ascii')
        return f'<img class="apta-image" data-name="{esc(e["path"].name)}" alt="{esc(e["description"])}" style="{geometry}" src="data:image/png;base64,{encoded}">'
    if e['kind']=='table':
        rows=[]
        for index,row in enumerate(e['rows']):
            tag='th' if index==0 else 'td'
            rows.append('<tr>'+''.join(f'<{tag}>{esc(cell)}</{tag}>' for cell in row)+'</tr>')
        return f'<table class="apta-table" style="{geometry}" aria-label="{esc(title)}"><tbody>{"".join(rows)}</tbody></table>'
    raise ValueError(f'Unsupported scene element: {e["kind"]}')


CSS = """
*{box-sizing:border-box}html,body{margin:0;background:#e5e5e5}body{font-family:Lato,Arial,sans-serif}
.deck{width:1280px;margin:0 auto}.slide{position:relative;width:1280px;height:720px;overflow:hidden;break-after:page}
.apta-text,.apta-shape,.apta-image,.apta-table{position:absolute;margin:0}.apta-text{white-space:normal;overflow:visible}
.apta-text p{margin:0;padding:0}.apta-image{object-fit:contain}.apta-table{border-collapse:collapse;table-layout:fixed;font:23px/1.38 Lato,Arial,sans-serif;color:#17234E}
.apta-table tr{height:20%}.apta-table th,.apta-table td{padding:8px 12px;border:1px solid #C9D2DF;text-align:left;vertical-align:top;overflow-wrap:normal}
.apta-table th{background:#17234E;color:white;font-weight:700}.apta-table td{background:white}.apta-table tr:nth-child(odd) td{background:#F0EFEA}
.speaker-notes{display:none}.sr-only{position:absolute;width:1px;height:1px;padding:0;overflow:hidden;clip-path:inset(50%);white-space:nowrap}
#deck-nav{position:fixed;bottom:20px;left:50%;transform:translateX(-50%);display:flex;gap:7px;align-items:center;padding:7px 10px;border:1px solid rgba(23,35,78,.18);border-radius:999px;background:rgba(246,243,234,.9);z-index:10}
.dot{position:relative;flex:none;display:block;width:7px;height:7px;padding:0;border:0;border-radius:999px;background:rgba(23,35,78,.28);cursor:pointer}
.dot::after{content:'';position:absolute;inset:-3.5px}.dot.active{width:22px;background:#EF7636}.dot:focus-visible{outline:3px solid #EF7636;outline-offset:5px}
body.present{overflow:hidden}body.present .deck-stage{min-height:100vh;display:flex;align-items:center;justify-content:center}.deck-stage{touch-action:pan-x}
body.present .slide{display:none}body.present .slide.active{display:block}
@media(pointer:coarse){#deck-nav{gap:12px;padding:12px}.dot::after{inset:-6px}}
@page{size:13.333333in 7.5in;margin:0}@media print{body.present{overflow:visible;background:white}body.present .deck-stage{display:block;min-height:0}.deck{zoom:1!important;margin:0}#deck-nav,#counter{display:none!important}body.present .slide{display:block!important}}
"""

JS = """
(() => {
const slides=[...document.querySelectorAll('.slide')], deck=document.querySelector('.deck');
const counter=document.querySelector('#counter'), dots=[...document.querySelectorAll('#deck-nav .dot')]; let current=0;
document.body.classList.add('present');
function show(n){current=Math.max(0,Math.min(slides.length-1,n));slides.forEach((s,i)=>s.classList.toggle('active',i===current));dots.forEach((d,i)=>{d.classList.toggle('active',i===current);if(i===current)d.setAttribute('aria-current','true');else d.removeAttribute('aria-current');});counter.textContent=`Slide ${current+1} of ${slides.length}`;}
function fit(){deck.style.zoom=Math.min(innerWidth/1280,innerHeight/720);}
dots.forEach((dot,i)=>dot.addEventListener('click',()=>show(i)));
addEventListener('keydown',e=>{if(e.altKey||e.ctrlKey||e.metaKey||e.target.closest('input,textarea,select,[contenteditable="true"]'))return;if(e.key===' '&&e.target.closest('button'))return;let target=current;if(['ArrowRight','ArrowDown','PageDown',' '].includes(e.key))target++;else if(['ArrowLeft','ArrowUp','PageUp'].includes(e.key))target--;else if(e.key==='Home')target=0;else if(e.key==='End')target=slides.length-1;else return;e.preventDefault();show(target);if(e.target.closest('#deck-nav'))dots[current].focus();});
let wheelAt=-Infinity;addEventListener('wheel',e=>{if(e.ctrlKey)return;const delta=Math.abs(e.deltaY)>=Math.abs(e.deltaX)?e.deltaY:e.deltaX;if(Math.abs(delta)<12)return;e.preventDefault();if(performance.now()-wheelAt<450)return;wheelAt=performance.now();show(current+Math.sign(delta));},{passive:false});
let touchStart=null;const stage=document.querySelector('.deck-stage');stage.addEventListener('touchstart',e=>{touchStart=e.touches.length===1?{x:e.touches[0].clientX,y:e.touches[0].clientY}:null;},{passive:true});stage.addEventListener('touchend',e=>{if(!touchStart||!e.changedTouches.length)return;const dx=touchStart.x-e.changedTouches[0].clientX,dy=touchStart.y-e.changedTouches[0].clientY;touchStart=null;const delta=Math.abs(dx)>Math.abs(dy)?dx:dy;if(Math.abs(delta)>45)show(current+Math.sign(delta));},{passive:true});
addEventListener('resize',fit);show(0);fit();
})();
"""


def compile_html(source,output):
    meta,items=read_source(source)
    slides=[compose(item,i+1,len(items),meta,source.parent) for i,item in enumerate(items)]
    sections=[]
    for slide,item in zip(slides,items):
        elements='\n'.join(element_html(e,item['title']) for e in slide.elements)
        sections.append(f'<section class="slide" data-layout="{slide.layout}" data-title="{esc(item["title"])}" style="background:{background_css(slide.background)}">\n{elements}\n<aside class="speaker-notes">{esc(item["notes"])}</aside>\n</section>')
    metadata='\n'.join(f'<meta name="apta:{esc(k)}" content="{esc(v)}">' for k,v in meta.items())
    dots=''.join(f'<button class="dot" type="button" aria-label="Go to slide {i+1}"></button>' for i in range(len(slides)))
    document=(f'<!doctype html>\n<html lang="{meta["language"]}">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width,initial-scale=1">\n'
              f'<meta name="apta:format" content="positioned-slides-v1">\n{metadata}\n<title>{esc(meta["title"])}</title>\n<style>{CSS}</style>\n</head>\n'
              f'<body><div class="deck-stage"><main class="deck">\n'+ '\n'.join(sections) + '\n</main></div>\n'
              f'<nav id="deck-nav" aria-label="Slide navigation">{dots}</nav><span id="counter" class="sr-only" aria-live="polite"></span>\n'
              f'<script>{JS}</script>\n</body></html>\n')
    output.parent.mkdir(parents=True,exist_ok=True)
    output.write_text(document,encoding='utf-8')
    return len(slides)


class Node:
    def __init__(self,tag='',attrs=()):
        self.tag,self.attrs,self.children=tag,dict(attrs),[]

    def text(self):
        return ''.join(c if isinstance(c,str) else c.text() for c in self.children)

    def all(self,tag):
        result=[]
        for c in self.children:
            if isinstance(c,Node):
                if c.tag==tag:
                    result.append(c)
                result.extend(c.all(tag))
        return result


class DocumentParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root=Node()
        self.stack=[self.root]

    def handle_starttag(self,tag,attrs):
        node=Node(tag,attrs)
        self.stack[-1].children.append(node)
        if tag not in {'meta','img','br','hr','link','col','input'}:
            self.stack.append(node)

    def handle_startendtag(self,tag,attrs):
        self.handle_starttag(tag,attrs)
        if self.stack[-1].tag==tag:
            self.stack.pop()

    def handle_endtag(self,tag):
        for i in range(len(self.stack)-1,0,-1):
            if self.stack[i].tag==tag:
                self.stack=self.stack[:i]
                break

    def handle_data(self,data):
        self.stack[-1].children.append(data)


def styles(node):
    return {key.strip():value.strip() for item in node.attrs.get('style','').split(';') if ':' in item for key,value in [item.split(':',1)]}


def number(value,unit='px'):
    raw=value.removesuffix(unit) if unit else value
    result=float(raw)
    if not math.isfinite(result):
        raise ValueError('HTML geometry must be finite.')
    return result


def geometry(style):
    values=[number(style[key]) for key in ('left','top','width','height')]
    if min(values[:2])<0 or min(values[2:])<=0:
        raise ValueError('HTML elements need nonnegative positions and positive dimensions.')
    return values


def parse_color(value):
    if value=='transparent':
        return None,100
    if re.fullmatch(r'#[0-9a-fA-F]{6}',value):
        return value[1:].upper(),100
    match=re.fullmatch(r'rgba\(\s*(\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\s*\)',value)
    if match:
        rgb=[int(match[i]) for i in (1,2,3)]
        opacity=float(match[4])*100
        if max(rgb)>255 or not 0<=opacity<=100:
            raise ValueError('HTML color components out of range.')
        return ''.join(f'{v:02X}' for v in rgb),opacity
    raise ValueError(f'Unsupported HTML color: {value}')


def set_background(slide,value):
    match=re.fullmatch(r'linear-gradient\(([\d.]+)deg,\s*(.+)\)',value)
    if match:
        stops=[]
        for stop in match[2].split(','):
            color_value,pos=stop.strip().split()
            shade,opacity=parse_color(color_value)
            if opacity!=100 or shade is None:
                raise ValueError('Background gradient needs opaque colors.')
            stops.append((round(number(pos,'%')*1000),shade))
        if len(stops)<2 or stops[0][0]!=0 or stops[-1][0]!=100000 or stops!=sorted(stops):
            raise ValueError('Gradient stops must be ordered from 0% to 100%.')
        angle=round(((float(match[1])-90)%360)*60000)
        slide.background={'stops':stops,'angle':angle}
        slide.bg_fill=slide.gradient(stops,angle)
    else:
        shade,opacity=parse_color(value)
        if opacity!=100 or shade is None:
            raise ValueError('Slide background must be opaque.')
        slide.background={'color':shade}
        slide.bg_fill=f'<a:solidFill><a:srgbClr val="{shade}"/></a:solidFill>'


def parse_html(source,image_dir):
    parser=DocumentParser()
    parser.feed(source.read_text(encoding='utf-8'))
    metadata={node.attrs.get('name','')[5:]:node.attrs.get('content','') for node in parser.root.all('meta') if node.attrs.get('name','').startswith('apta:')}
    if metadata.pop('format',None)!='positioned-slides-v1':
        raise ValueError('Expected presentation HTML generated by this skill; arbitrary HTML/CSS is unsupported.')
    if metadata.get('language') not in {'en','id'}:
        raise ValueError('HTML needs apta:language en or id.')
    sections=[n for n in parser.root.all('section') if 'slide' in n.attrs.get('class','').split()]
    if not sections:
        raise ValueError('HTML contains no presentation slides.')
    items,slides=[],[]
    for index,section in enumerate(sections):
        layout=section.attrs.get('data-layout','')
        slide=Slide(layout,index+1,len(sections),metadata,furniture=False)
        set_background(slide,styles(section)['background'])
        notes=[n.text() for n in section.children if isinstance(n,Node) and n.tag=='aside' and n.attrs.get('class')=='speaker-notes']
        title_nodes=[n for n in section.children if isinstance(n,Node) and n.attrs.get('data-role')=='title']
        if len(title_nodes)!=1:
            raise ValueError('Each HTML slide needs one visible data-role="title" text element.')
        title='\n'.join(n.text() for n in title_nodes[0].all('p'))
        items.append({'title':title,'layout':layout,'lines':[],'notes':'\n'.join(notes)})
        for child in section.children:
            if not isinstance(child,Node) or child.tag=='aside':
                continue
            kind=child.attrs.get('class','')
            style=styles(child)
            x,y,w,h=geometry(style)
            if kind=='apta-text':
                text='\n'.join(n.text() for n in child.all('p'))
                shade,opacity=parse_color(style['color'])
                align={'left':'l','right':'r','center':'ctr'}[style['text-align']]
                font=style['font-family'].strip("'\"")
                line_height=number(style['line-height'],'')/FONT_LINE_BOX.get(font,1.2)
                size=number(style['font-size'])
                if size<=0 or line_height<=0:
                    raise ValueError('HTML font size and line height must be positive.')
                slide.text(text,x,y,w,h,size,shade,style['font-weight']=='700',align,font,style['font-style']=='italic',opacity,line_height)
            elif kind=='apta-shape':
                shade,opacity=parse_color(style['background'])
                stroke=None if style['border']=='none' else parse_color(style['border'].split()[-1])[0]
                preset=child.attrs.get('data-shape')
                if preset not in {'rect','roundRect','ellipse'}:
                    raise ValueError(f'Unsupported HTML shape: {preset}')
                slide.rect(x,y,w,h,shade,opacity,stroke,preset,child.attrs.get('data-name','Shape'))
            elif kind=='apta-image':
                src=child.attrs.get('src','')
                if not src.startswith('data:image/png;base64,'):
                    raise ValueError('Presentation HTML must embed PNG images as data URLs; external/local URLs are unsupported.')
                image_name=Path(child.attrs.get('data-name','image.png')).name
                if not image_name.lower().endswith('.png'):
                    raise ValueError('HTML image name must end in .png.')
                image_path=image_dir/f'{index}'/image_name
                image_path.parent.mkdir(parents=True,exist_ok=True)
                image_path.write_bytes(base64.b64decode(src.split(',',1)[1],validate=True))
                slide.image(image_path,x,y,w,h,child.attrs.get('alt',''))
            elif kind=='apta-table':
                rows=[[cell.text() for cell in row.children if isinstance(cell,Node) and cell.tag in {'th','td'}] for row in child.all('tr')]
                slide.table(rows,x,y,w,h)
            else:
                raise ValueError(f'Unsupported element in slide HTML: {child.tag}.{kind}')
        slides.append(slide)
    return metadata,items,slides


def export_html(source,output):
    with tempfile.TemporaryDirectory(prefix='apta-html-') as folder:
        meta,items,slides=parse_html(source,Path(folder))
        return export_pptx(meta,items,slides,output)


def validate_html(source):
    with tempfile.TemporaryDirectory(prefix='apta-html-check-') as folder:
        return parse_html(source,Path(folder))[1]
