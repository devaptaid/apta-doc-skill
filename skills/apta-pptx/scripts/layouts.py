"""Native slide compositions for the Aptaworks presentation system."""
from __future__ import annotations
import re
from urllib.parse import unquote
from build_deck import Slide, ASSETS, NAVY, ORANGE, MUTED, RULE, clean_lines, wrap


def fields(lines, minimum, maximum, count, layout):
    values = clean_lines(lines)
    if not minimum <= len(values) <= maximum:
        raise ValueError(f'{layout} needs {minimum}–{maximum} rows.')
    result = [value.split(' | ') for value in values]
    if any(len(parts) != count or any(not p.strip() for p in parts) for parts in result):
        raise ValueError(f'{layout} requires exactly {count} nonempty fields per row separated by " | ".')
    return result


def eyebrow(s, label, y=54):
    s.rect(84,y+7,40,3,ORANGE)
    s.text(label.upper(),136,y,1000,22,11,ORANGE,font='Consolas')


def heading(s, title, label):
    eyebrow(s,label)
    lines=wrap(title,1112,48,2,font='Lato Black')
    s.text(title,84,82,1112,60*lines,48,font='Lato Black',line_height=1.02,role='title')
    # Reserve a second title line only when needed; components use the freed space.
    return 184+(lines-1)*60


def compose(item, number, total, meta, source_dir):
    layout,title,lines = item['layout'],item['title'],item['lines']
    s = Slide(layout,number,total,meta)
    if layout == 'cover':
        s.image(ASSETS/'logo_aptaworks_wide_white.png',84,54,150,31)
        eyebrow(s,'Aptaworks presentation library',140)
        s.text(title,84,190,900,200,64,font='Lato Black',line_height=1.02,role='title')
        s.text('\n'.join(lines),84,430,790,120,22,'FFFFFF',opacity=78,line_height=1.45)
        s.text('A',1000,384,200,265,270,'FFFFFF',font='Georgia',italic=True,opacity=14,line_height=.94)
    elif layout == 'section':
        if len(lines)>1:
            raise ValueError('Section accepts one supporting line.')
        s.image(ASSETS/'logo_aptaworks_wide_white.png',84,54,150,31)
        eyebrow(s,'Section divider',198)
        s.text(title,84,245,815,150,48,font='Lato Black',line_height=1.02,role='title')
        s.text('\n'.join(lines),84,420,740,90,23,'FFFFFF',opacity=78,line_height=1.4)
        s.text(f'{number:02d}',880,415,320,240,255,'FFFFFF',font='Georgia',italic=True,opacity=16,line_height=.94)
    elif layout == 'closing':
        if len(lines)!=2:
            raise ValueError('Closing requires a secondary headline and the consultation subtitle, one line each.')
        # These positions and sizes follow the 1280×720 upstream HTML reference.
        s.text(title,84,156,1112,54,41.6,font='Lato Black',align='ctr',line_height=1.14,role='title')
        s.text(lines[0],84,210,1112,44,31.6,'FFFFFF',bold=True,align='ctr',opacity=78,line_height=1.14)
        s.text(lines[1],84,328,1112,54,35.8,'FFFFFF',align='ctr',line_height=1.2)
        s.rect(569,447,142,142,'FFFFFF',preset='roundRect',name='Consultation QR card')
        s.image(ASSETS/'qr-apta.png',577,455,126,126,'Aptaworks consultation QR code')
        s.image(ASSETS/'Footer.png',80,610,1120,56,'(+62) 811 862 2782; Instagram aptaworks; LinkedIn aptaworks; www.apta.id; info@apta.id')
    elif layout == 'table':
        top=heading(s,title,'Editable workplan / native table')
        rows=[]
        for line in lines:
            if not line.startswith('|') or not line.endswith('|'):
                raise ValueError('Table accepts Markdown pipe rows only.')
            values=[v.strip() for v in line.strip('|').split('|')]
            if all(re.fullmatch(r':?-{3,}:?',v) for v in values):
                continue
            rows.append(values)
        if not rows:
            raise ValueError('Table is empty.')
        s.table(rows,y=top,height=min(636-top,64*len(rows)))
    elif layout == 'ledger':
        top=heading(s,title,'Option sets / ledger')
        values=clean_lines(lines)
        if not 1<=len(values)<=5:
            raise ValueError('Ledger accepts 1–5 rows.')
        rich=any(' | ' in v for v in values)
        if rich and (len(values)>3 or any(len(v.split(' | '))!=3 for v in values)):
            raise ValueError('Detailed ledger accepts 1–3 rows: Label | Description | Status.')
        step=min(132 if rich else 82,(636-top)/len(values))
        s.rect(84,top,1112,1.5,NAVY)
        for i,value in enumerate(values):
            y=top+20+i*step
            s.text(f'{i+1:02d}',84,y-2,77,70,54,NAVY,font='Georgia',italic=True,line_height=1)
            if rich:
                label,description,state=value.split(' | ')
                s.text(label,186,y,800,40,27,font='Lato Black')
                s.text(description,186,y+42,800,49,18,MUTED,line_height=1.4)
                state_color={'READY':'2C8D63','PILOT':ORANGE,'REVIEW':'D94B35'}.get(state.upper(),MUTED)
                s.rect(1038,y+4,158,34,None,stroke=state_color,name='Ledger state')
                s.text(state,1046,y+10,142,24,12,state_color,align='ctr',font='Consolas')
            else:
                s.text(value,186,y+5,1010,47,27,font='Lato Black')
            s.rect(84,y+step-20,1112,1,NAVY,opacity=42)
    elif layout in {'kpi','proof'}:
        top=heading(s,title,'Illustrative metrics' if layout=='kpi' else 'Structured evidence / proof grid')
        values=fields(lines,3 if layout=='kpi' else 6,4 if layout=='kpi' else 6,3,layout)
        cols=len(values) if layout=='kpi' else 3
        width=1112/cols
        row_h=min(330,636-top) if layout=='kpi' else min(196,(636-top)/2)
        for i,(value,label,context) in enumerate(values):
            x,y=84+(i%cols)*width,top+(i//cols)*row_h
            s.rect(x,y,width,row_h,'BEC1CD',stroke=NAVY,name='Metric cell' if layout=='kpi' else 'Proof cell')
            size=65 if layout=='kpi' else 39
            s.text(value,x+22,y+22,width-44,92 if layout=='kpi' else 52,size,NAVY,font='Georgia',italic=True,line_height=1)
            label_y=y+126 if layout=='kpi' else y+82
            s.text(label,x+22,label_y,width-44,43,21 if layout=='kpi' else 18,NAVY,bold=True)
            s.text(context,x+22,label_y+48 if layout=='kpi' else label_y+28,width-44,60 if layout=='kpi' else 35,17 if layout=='kpi' else 14,MUTED,line_height=1.35)
    elif layout=='cards':
        top=heading(s,title,'Rule-framed cards / workstreams')
        height=min(420,636-top)
        values=fields(lines,2,3,3,layout)
        gap=24
        width=(1112-gap*(len(values)-1))/len(values)
        for i,(label,description,detail) in enumerate(values):
            x=84+i*(width+gap)
            s.rect(x,top,width,height,'FFFFFF',opacity=58,stroke=NAVY,name='Editable workstream card')
            s.text(f'{i+1:02d}',x+25,top+24,width-50,55,42,NAVY,font='Georgia',italic=True,line_height=1)
            s.text(label,x+25,top+98,width-50,72,28,font='Lato Black',line_height=1.12)
            s.text(description,x+25,top+183,width-50,height-235,19,MUTED,line_height=1.4)
            s.text(detail,x+25,top+height-36,width-50,24,12,ORANGE,font='Consolas')
    elif layout in {'process','flow'}:
        top=heading(s,title,'Delivery logic / connected flow')
        values=fields(lines,2,4,2,layout)
        gap=32
        width=(1112-gap*(len(values)-1))/len(values)
        s.rect(84,top+48,1112,3,ORANGE)
        for i,(label,description) in enumerate(values):
            x=84+i*(width+gap)
            s.rect(x,top+28,42,42,ORANGE,preset='ellipse',name='Editable flow node')
            s.text(str(i+1),x,top+37,42,24,15,'FFFFFF',align='ctr',font='Consolas')
            s.text(label,x,top+110,width,72,25,font='Lato Black',line_height=1.15)
            s.text(description,x,top+197,width,min(180,636-top-197),20,MUTED,line_height=1.45)
    elif layout=='timeline':
        top=heading(s,title,'Milestones / editable timeline')
        values=fields(lines,3,4,3,layout)
        width=1112/len(values)
        s.rect(84,top+80,1112,3,ORANGE)
        for i,(date,label,description) in enumerate(values):
            x=84+i*width
            s.text(date,x,top+20,width-28,30,15,ORANGE,font='Consolas')
            s.rect(x,top+72,18,18,'FFFFFF',stroke=ORANGE,preset='ellipse',name='Timeline milestone')
            s.text(label,x,top+115,width-28,78,27,font='Lato Black',line_height=1.12)
            s.text(description,x,top+214,width-28,min(170,636-top-214),19,MUTED,line_height=1.4)
    elif layout=='columns':
        top=heading(s,title,'Narrative / two-column proposal')
        height=min(400,636-top)
        values=fields(lines,2,2,2,layout)
        for i,(label,description) in enumerate(values):
            x=84+i*584
            s.rect(x,top,528,height,'FFFFFF',opacity=58,stroke=NAVY,name='Narrative panel')
            s.text(label,x+26,top+28,476,88,32,font='Lato Black',line_height=1.12)
            s.text(description,x+26,top+142,476,height-168,23,MUTED,line_height=1.5)
    elif layout=='image':
        top=heading(s,title,'Product / visual evidence')
        height=min(410,636-top)
        if len(lines)!=2:
            raise ValueError('Image requires a local PNG and one caption.')
        match=re.fullmatch(r'!\[([^\]]+)\]\(([^)]+)\)',lines[0])
        if not match:
            raise ValueError('Image syntax: ![Meaningful alt text](relative.png)')
        s.rect(84,top,710,height,'FFFFFF',stroke=RULE,name='Source image frame')
        s.image(source_dir/unquote(match[2]),98,top+14,682,height-28,match[1])
        s.text(lines[1],846,top+15,350,height-30,25,NAVY,line_height=1.35)
    elif layout=='quote':
        top=heading(s,title,'Quote / principle')
        height=min(350,636-top)
        if len(lines)!=2:
            raise ValueError('Quote needs one statement and one attribution line.')
        s.rect(84,top+12,5,height,ORANGE)
        s.text(lines[0],126,top+10,1020,height-70,43,NAVY,font='Georgia',italic=True,line_height=1.28)
        s.text(lines[1],126,top+height-24,1020,32,15,MUTED,font='Consolas')
    else:
        raise ValueError(f'Layout not implemented: {layout}')
    return s
