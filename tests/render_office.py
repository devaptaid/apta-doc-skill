"""Optional Windows QA: render the real PPTX/DOCX using installed Office.

Needs pywin32 (not a generator dependency). Preserves existing Office sessions;
only the opened artifact is closed. Does not save changes to the source file.
"""
from __future__ import annotations
import argparse
import subprocess
from pathlib import Path


def render(source, target):
    import pythoncom
    from win32com.client import dynamic
    source, target = source.resolve(), target.resolve()
    if not source.is_file():
        raise ValueError(f'Input missing: {source}')
    if target.exists():
        raise ValueError(f'Output already exists: {target}')
    if source.suffix.lower() not in {'.pptx','.docx'}:
        raise ValueError('Expected .pptx or .docx')
    target.parent.mkdir(parents=True,exist_ok=True)
    ppt = source.suffix.lower()=='.pptx'
    program, process = ('PowerPoint.Application','POWERPNT.EXE') if ppt else ('Word.Application','WINWORD.EXE')
    running = subprocess.run(['tasklist','/FI',f'IMAGENAME eq {process}','/NH'],capture_output=True,text=True,check=True)
    had_session = process.lower() in running.stdout.lower()
    # Late binding avoids dependence on Office type-library registration caches.
    app = dynamic.Dispatch(pythoncom.CoCreateInstance(program,None,pythoncom.CLSCTX_LOCAL_SERVER,pythoncom.IID_IDispatch))
    document = None
    try:
        if ppt:
            document = app.Presentations.Open(str(source),True,False,False)
            document.SaveAs(str(target),32)
            print(f'PowerPoint: rendered {document.Slides.Count} slides')
        else:
            document = app.Documents.Open(str(source),False,True)
            document.Repaginate()
            document.ExportAsFixedFormat(str(target),17)
            print(f'Word: rendered {document.ComputeStatistics(2)} pages')
    finally:
        if document is not None:
            document.Close() if ppt else document.Close(False)
        if not had_session:
            app.Quit()
    print(target)


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source',type=Path)
    parser.add_argument('target_pdf',type=Path)
    args=parser.parse_args()
    render(args.source,args.target_pdf)
