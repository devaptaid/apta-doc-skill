#!/usr/bin/env python3
"""Validate PPTX relationships, bounds, content, native text and source coverage."""
from __future__ import annotations
import argparse
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from build_deck import A, P, R, PKG, EMU, W, H, read_source


def validate(path, source=None):
    if path.suffix.lower()=='.html':
        from deck_html import validate_html
        items=validate_html(path)
        if source and source.suffix.lower()=='.md' and len(items)!=len(read_source(source)[1]):
            return ['HTML slide count differs from Markdown.']
        return []
    errors = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        required = {'[Content_Types].xml','_rels/.rels','ppt/presentation.xml','ppt/theme/theme1.xml'}
        if required - names:
            return [f'Missing package parts: {sorted(required - names)}']
        roots = {name:ET.fromstring(archive.read(name)) for name in names if name.endswith(('.xml','.rels'))}
        for name,root in roots.items():
            if not name.endswith('.rels'):
                continue
            owner_dir = posixpath.dirname(name).replace('/_rels','')
            if name == '_rels/.rels':
                owner_dir = ''
            for rel in root:
                if rel.get('TargetMode') == 'External':
                    errors.append(f'Unexpected external relationship: {name}')
                    continue
                target = posixpath.normpath(posixpath.join(owner_dir,rel.get('Target','')))
                if target not in names:
                    errors.append(f'Broken relationship {name}: {target}')
        presentation = roots['ppt/presentation.xml']
        size = presentation.find(f'{{{P}}}sldSz')
        if size is None or int(size.get('cx','0')) != W*EMU or int(size.get('cy','0')) != H*EMU:
            errors.append('Slide size is not 1280×720 / 16:9.')
        slides = sorted((n for n in roots if re.fullmatch(r'ppt/slides/slide\d+\.xml',n)),key=lambda n:int(re.search(r'(\d+)\.xml',n)[1]))
        if not slides:
            errors.append('No slides.')
        if source and source.suffix.lower()=='.html':
            from deck_html import validate_html
            sources=validate_html(source)
        else:
            sources = read_source(source)[1] if source else None
        if sources is not None and len(sources) != len(slides):
            errors.append(f'Slide count differs from source: {len(slides)} vs {len(sources)}')
        slide_ids = presentation.findall(f'{{{P}}}sldIdLst/{{{P}}}sldId')
        if len(slide_ids) != len(slides):
            errors.append('Slide list count differs from slide package parts.')
        for i,name in enumerate(slides):
            root = roots[name]
            text = '\n'.join(node.text or '' for node in root.findall(f'.//{{{A}}}t'))
            brand_evidence = text + ' '.join(n.get('name','') for n in root.findall(f'.//{{{P}}}pic/{{{P}}}nvPicPr/{{{P}}}cNvPr'))
            if 'AptaWorks' in text or 'aptaworks' not in brand_evidence.lower():
                errors.append(f'{name}: incorrect or missing branding')
            if not root.findall(f'.//{{{P}}}sp/{{{P}}}txBody'):
                errors.append(f'{name}: no native editable text')
            if root.find(f'{{{P}}}cSld').get('name') in {'cover','section','closing'} and not root.findall(f'.//{{{P}}}pic'):
                errors.append(f'{name}: branded cover, divider, or closing is missing artwork')
            for xfrm in root.findall(f'.//{{{A}}}xfrm') + root.findall(f'.//{{{P}}}xfrm'):
                off,ext = xfrm.find(f'{{{A}}}off'),xfrm.find(f'{{{A}}}ext')
                if off is None or ext is None:
                    continue
                x,y,cx,cy = (int(off.get('x',0)),int(off.get('y',0)),int(ext.get('cx',0)),int(ext.get('cy',0)))
                if min(x,y,cx,cy)<0 or x+cx>W*EMU+EMU or y+cy>H*EMU+EMU:
                    errors.append(f'{name}: shape outside slide')
            if sources is not None and i < len(sources):
                item = sources[i]
                if item['title'] not in text:
                    errors.append(f'{name}: source title missing')
                expected = []
                for line in item['lines']:
                    if line.startswith('![') or re.fullmatch(r'\|[\s:|\-]+\|',line):
                        continue
                    if line.startswith('|'):
                        expected.extend(v.strip() for v in line.strip('|').split('|'))
                    else:
                        expected.extend(line.removeprefix('- ').split(' | ') if item['layout'] in {'columns','process','flow','timeline','ledger','cards','kpi','proof'} else [line.removeprefix('- ')])
                for value in expected:
                    if value and value not in text:
                        errors.append(f'{name}: source text missing: {value[:70]}')
                if item['layout']=='table' and not root.findall(f'.//{{{A}}}tbl'):
                    errors.append(f'{name}: table is not native')
                if item['layout']=='closing':
                    stops=[n.get('val') for n in root.findall(f'{{{P}}}cSld/{{{P}}}bg/{{{P}}}bgPr/{{{A}}}gradFill/{{{A}}}gsLst/{{{A}}}gs/{{{A}}}srgbClr')]
                    if stops != ['182B63','152B5B','16A6C8']:
                        errors.append(f'{name}: consultation closing gradient differs from the reference')
                    pictures=root.findall(f'.//{{{P}}}pic/{{{P}}}nvPicPr/{{{P}}}cNvPr')
                    picture_names=[n.get('name') for n in pictures]
                    if 'Footer.png' not in picture_names or 'qr-apta.png' not in picture_names:
                        errors.append(f'{name}: reference QR or full contact row missing')
                for graphic in root.findall(f'.//{{{A}}}graphicData'):
                    if graphic.find(f'{{{A}}}tbl') is not None and graphic.get('uri') != 'http://schemas.openxmlformats.org/drawingml/2006/table':
                        errors.append(f'{name}: table graphicData uses an invalid schema URI')
                if item['notes']:
                    notes = roots.get(f'ppt/notesSlides/notesSlide{i+1}.xml')
                    notes_text = '\n'.join(n.text or '' for n in notes.findall(f'.//{{{A}}}t')) if notes is not None else ''
                    if item['notes'] not in notes_text:
                        errors.append(f'{name}: source notes missing')
    return errors


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx',type=Path)
    parser.add_argument('--source',type=Path)
    args=parser.parse_args()
    try:
        errors=validate(args.pptx,args.source)
    except (OSError,ValueError,KeyError,ET.ParseError,zipfile.BadZipFile) as exc:
        errors=[str(exc)]
    for error in errors:
        print('FAIL: '+error,file=sys.stderr)
    if not errors:
        print(f'PASS: {args.pptx} (structural only; render for visual QA)')
    return int(bool(errors))


if __name__=='__main__':
    raise SystemExit(main())
