"""Observable regressions for portable installs, exports, and malformed input."""
from __future__ import annotations
import importlib.util
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
DOC = ROOT/'skills/apta-doc-proposal'
PPT = ROOT/'skills/apta-pptx'


def run(script, *args, cwd=None):
    return subprocess.run([sys.executable,str(script),*map(str,args)],cwd=cwd,capture_output=True,text=True)


class Exports(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.work = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_standalone_ppt_install_preserves_content_and_notes(self):
        installed = self.work/'installed skill'
        shutil.copytree(PPT,installed,ignore=shutil.ignore_patterns('__pycache__'))
        output = self.work/'demo.pptx'
        source = installed/'examples/example.md'
        result = run(installed/'scripts/build_deck.py',source,'--output',output,cwd=self.work)
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertTrue(output.with_suffix('.html').is_file())
        result = run(installed/'scripts/validate_deck.py',output,'--source',source,cwd=self.work)
        self.assertEqual(result.returncode,0,result.stderr)
        with zipfile.ZipFile(output) as archive:
            self.assertEqual(len([name for name in archive.namelist() if re.fullmatch(r'ppt/slides/slide\d+.xml',name)]),13)
            table_xml=archive.read('ppt/slides/slide8.xml').decode()
            self.assertIn('http://schemas.openxmlformats.org/drawingml/2006/table',table_xml)
            self.assertIn('Practical decks',archive.read('ppt/slides/slide1.xml').decode())

    def test_ppt_rejects_overflow_missing_assets_and_unsupported_layouts(self):
        cases = [
            '## Topic {unknown}\n- Row',
            '## Topic {ledger}\n- '+'a '*1000,
            '## Topic {table}\n| A | B |\n|---|---|\n| Unequal |',
            '## Topic {image}\n![Missing](absent.png)\nCaption.',
            '## Topic {ledger}\n- **Unsupported formatting**',
        ]
        for index,text in enumerate(cases):
            with self.subTest(index=index):
                source=self.work/f'bad{index}.md'
                source.write_text(text,encoding='utf-8')
                output=self.work/f'bad{index}.pptx'
                result=run(PPT/'scripts/build_deck.py',source,'--output',output)
                self.assertNotEqual(result.returncode,0)
                self.assertFalse(output.exists())

    def test_ppt_validator_detects_content_loss(self):
        source=self.work/'check.md'
        source.write_text('## Agenda {ledger}\n- Unchanged',encoding='utf-8')
        output=self.work/'check.pptx'
        self.assertEqual(run(PPT/'scripts/build_deck.py',source,'--output',output).returncode,0)
        source.write_text('## Agenda {ledger}\n- Lost source text',encoding='utf-8')
        result=run(PPT/'scripts/validate_deck.py',output,'--source',source)
        self.assertNotEqual(result.returncode,0)
        self.assertIn('source text missing',result.stderr)

    def test_html_is_the_real_pptx_input_and_is_portable(self):
        source=self.work/'brief.md'
        source.write_text('## Original headline {ledger}\n- First item\n- Second item',encoding='utf-8')
        html_path=self.work/'brief.html'
        result=run(PPT/'scripts/build_deck.py',source,'--output',html_path)
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertFalse(html_path.with_suffix('.pptx').exists())
        source.unlink()
        # Edit the visible HTML and its inline typography, then convert without MD.
        html_path.write_text(html_path.read_text(encoding='utf-8').replace('Original headline','Reviewed headline').replace('font-size:48px','font-size:44px'),encoding='utf-8')
        moved=self.work/'portable'/'reviewed.html'
        moved.parent.mkdir()
        shutil.move(html_path,moved)
        output=self.work/'reviewed.pptx'
        result=run(PPT/'scripts/build_deck.py',moved,'--output',output)
        self.assertEqual(result.returncode,0,result.stderr)
        with zipfile.ZipFile(output) as archive:
            xml=archive.read('ppt/slides/slide1.xml').decode()
            self.assertIn('Reviewed headline',xml)
            self.assertNotIn('Original headline',xml)
            self.assertIn('sz="3300"',xml)

    def test_clean_pptx_has_no_background_grid(self):
        output=self.work/'clean.pptx'
        self.assertEqual(run(PPT/'scripts/build_deck.py',PPT/'examples/example.md','--output',output).returncode,0)
        a='http://schemas.openxmlformats.org/drawingml/2006/main'
        with zipfile.ZipFile(output) as archive:
            for name in archive.namelist():
                if not re.fullmatch(r'ppt/slides/slide\d+.xml',name):
                    continue
                root=ET.fromstring(archive.read(name))
                for transform in root.findall(f'.//{{{a}}}xfrm'):
                    size=transform.find(f'{{{a}}}ext')
                    if size is None:
                        continue
                    w,h=int(size.get('cx')),int(size.get('cy'))
                    self.assertFalse((w<=9525 and h==720*9525) or (h<=9525 and w==1280*9525),name)

    def test_repeated_title_and_body_still_round_trip_through_html(self):
        source=self.work/'repeated.md'
        source.write_text('## Aptaworks {cover}\nAptaworks',encoding='utf-8')
        output=self.work/'repeated.pptx'
        result=run(PPT/'scripts/build_deck.py',source,'--output',output)
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertTrue(output.exists())

    def test_html_rejects_arbitrary_pages_and_external_images(self):
        html_path=self.work/'arbitrary.html'
        html_path.write_text('<html><body><h1>Not a slide contract</h1></body></html>',encoding='utf-8')
        output=self.work/'invalid.pptx'
        self.assertNotEqual(run(PPT/'scripts/build_deck.py',html_path,'--output',output).returncode,0)
        self.assertFalse(output.exists())
        source=PPT/'examples/example.md'
        self.assertEqual(run(PPT/'scripts/build_deck.py',source,'--output',html_path).returncode,0)
        html_path.write_text(re.sub(r'src="data:image/png;base64,[^"]+"','src="https://example.invalid/image.png"',html_path.read_text(encoding='utf-8')),encoding='utf-8')
        result=run(PPT/'scripts/build_deck.py',html_path,'--output',output)
        self.assertNotEqual(result.returncode,0)
        self.assertIn('embed PNG',result.stderr)

    def test_document_export_gate_has_no_outputs(self):
        output=self.work/'out'
        result=run(DOC/'scripts/build_proposal.py',DOC/'examples/example.md','--output-dir',output,'--docx')
        self.assertEqual(result.returncode,2)
        self.assertFalse(output.exists())

    def test_document_portability_simple_heading_and_word_fields(self):
        installed=self.work/'document skill'
        shutil.copytree(DOC,installed,ignore=shutil.ignore_patterns('__pycache__'))
        source=self.work/'brief.md'
        shutil.copy2(DOC/'assets/logo_aptaworks_wide_navy.png',self.work/'client logo.png')
        source.write_text('---\ntitle: Demo\nclient: Internal\nauthor: Aptaworks Team\n---\n\n# 1. Scope\n\nA short brief.\n\n![Client](client%20logo.png)\n',encoding='utf-8')
        output=self.work/'out'
        result=run(installed/'scripts/build_proposal.py',source,'--output-dir',output,'--docx','--approved',cwd=self.work)
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertFalse(list(output.glob('*.pdf')))
        self.assertIn('1. Scope',(output/'brief.html').read_text(encoding='utf-8'))
        self.assertNotIn('1. 1. Scope',(output/'brief.html').read_text(encoding='utf-8'))
        self.assertNotIn('file://',(output/'brief.html').read_text(encoding='utf-8'))
        moved=self.work/'portable'
        shutil.copytree(output,moved)
        shutil.rmtree(output)
        result=run(installed/'scripts/validate_outputs.py',moved/'brief.html',moved/'brief.docx',cwd=self.work)
        self.assertEqual(result.returncode,0,result.stdout+result.stderr)
        with zipfile.ZipFile(moved/'brief.docx') as archive:
            root=ET.fromstring(archive.read('word/document.xml'))
            ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            instr=' '.join(n.text or '' for n in root.findall('.//w:instrText',ns))
            self.assertIn('TOC',instr)
            self.assertIn('PAGEREF',instr)
            self.assertNotIn('updateFields',archive.read('word/settings.xml').decode())

    def test_missing_document_image_is_an_error(self):
        source=self.work/'missing.md'
        source.write_text('# Scope\n\n![Missing](missing.png)',encoding='utf-8')
        result=run(DOC/'scripts/build_proposal.py',source,'--output-dir',self.work/'out')
        self.assertNotEqual(result.returncode,0)
        self.assertIn('Image not found',result.stderr)

    def test_english_document_localizes_contents_and_cover(self):
        output=self.work/'english'
        result=run(DOC/'scripts/build_proposal.py',DOC/'examples/example.md','--output-dir',output,'--docx','--approved')
        self.assertEqual(result.returncode,0,result.stderr)
        markup=(output/'example.html').read_text(encoding='utf-8')
        self.assertIn('<html lang="en">',markup)
        self.assertIn('CONTENTS',markup)
        self.assertNotIn('DAFTAR ISI',markup)
        with zipfile.ZipFile(output/'example.docx') as archive:
            document=archive.read('word/document.xml').decode()
            self.assertIn('CONTENTS',document)
            self.assertIn('Author:',document)
            self.assertNotIn('Penulis:',document)
            self.assertNotIn('Tanggal:',document)

    def test_document_exports_the_reviewed_html_without_markdown(self):
        source=self.work/'brief.md'
        source.write_text('---\nlanguage: en\ntitle: Review example\n---\n\n# Scope\n\nOriginal body text.',encoding='utf-8')
        draft=self.work/'draft'
        result=run(DOC/'scripts/build_proposal.py',source,'--output-dir',draft)
        self.assertEqual(result.returncode,0,result.stderr)
        self.assertFalse(list(draft.glob('*.docx')))
        source.unlink()
        html_path=draft/'brief.html'
        html_path.write_text(html_path.read_text(encoding='utf-8').replace('Original body text.','Reviewed body text.'),encoding='utf-8')
        output=self.work/'final'
        result=run(DOC/'scripts/build_proposal.py',html_path,'--output-dir',output,'--docx','--approved')
        self.assertEqual(result.returncode,0,result.stderr)
        with zipfile.ZipFile(output/'brief.docx') as archive:
            document=archive.read('word/document.xml').decode()
            self.assertIn('Reviewed body text.',document)
            self.assertNotIn('Original body text.',document)

    def test_closing_retains_reference_gradient_qr_and_full_contact_artwork(self):
        output=self.work/'library.pptx'
        source=PPT/'examples/example.md'
        self.assertEqual(run(PPT/'scripts/build_deck.py',source,'--output',output).returncode,0)
        with zipfile.ZipFile(output) as archive:
            root=ET.fromstring(archive.read('ppt/slides/slide13.xml'))
            p='http://schemas.openxmlformats.org/presentationml/2006/main'
            a='http://schemas.openxmlformats.org/drawingml/2006/main'
            shapes=root.findall(f'.//{{{p}}}pic')
            descriptions=[node.find(f'{{{p}}}nvPicPr/{{{p}}}cNvPr').get('descr','') for node in shapes]
            self.assertTrue(any('811 862 2782' in value and 'info@apta.id' in value for value in descriptions))
            qr=next(node for node in shapes if node.find(f'{{{p}}}nvPicPr/{{{p}}}cNvPr').get('name')=='qr-apta.png')
            off=qr.find(f'{{{p}}}spPr/{{{a}}}xfrm/{{{a}}}off')
            ext=qr.find(f'{{{p}}}spPr/{{{a}}}xfrm/{{{a}}}ext')
            self.assertAlmostEqual(int(off.get('x'))+int(ext.get('cx'))/2,1280*9525/2,delta=9525)
            rels=ET.fromstring(archive.read('ppt/slides/_rels/slide13.xml.rels'))
            footer_pic=next(node for node in shapes if node.find(f'{{{p}}}nvPicPr/{{{p}}}cNvPr').get('name')=='Footer.png')
            r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            rid=footer_pic.find(f'{{{p}}}blipFill/{{{a}}}blip').get(f'{{{r}}}embed')
            target=next(node.get('Target') for node in rels if node.get('Id')==rid)
            self.assertEqual(archive.read('ppt/'+target.removeprefix('../')),(PPT/'references/Footer.png').read_bytes())


if __name__=='__main__':
    unittest.main()
