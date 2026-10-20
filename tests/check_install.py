"""Install this package with the real skills CLI into an isolated Claude project.

No global profile is changed. Supply a new workspace directory for each run.
"""
from __future__ import annotations
import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source',default=str(Path(__file__).resolve().parents[1]))
    parser.add_argument('--workspace',type=Path,required=True)
    parser.add_argument('--copy',action='store_true',help='Test copy mode instead of the CLI default installation mode')
    args=parser.parse_args()
    workspace=args.workspace.resolve()
    if workspace.exists():
        raise SystemExit('Use a new empty installation-check workspace.')
    workspace.mkdir(parents=True)
    npx=shutil.which('npx.cmd') or shutil.which('npx')
    if not npx:
        raise SystemExit('Missing npx; install a supported Node.js runtime first.')
    command=[npx,'--yes','skills','add',args.source,'--skill','apta-doc-proposal','apta-pptx','--agent','claude-code','--yes']
    if args.copy:
        command.append('--copy')
    result=subprocess.run(command,cwd=workspace,text=True,capture_output=True)
    (workspace/'install.log').write_text(result.stdout+result.stderr,encoding='utf-8')
    if result.returncode:
        print(result.stdout+result.stderr)
        return result.returncode
    checks={}
    for name in ('apta-doc-proposal','apta-pptx'):
        folder=workspace/'.claude/skills'/name
        checks[name]={'installed':(folder/'SKILL.md').is_file()}
        if not checks[name]['installed']:
            raise SystemExit(f'Claude skill folder missing: {folder}')
        output=workspace/'output'/name
        if name=='apta-pptx':
            command=[sys.executable,str(folder/'scripts/build_deck.py'),str(folder/'examples/example.md'),'--output',str(output/'example.pptx')]
        else:
            command=[sys.executable,str(folder/'scripts/build_proposal.py'),str(folder/'examples/example.md'),'--output-dir',str(output),'--docx','--approved']
        built=subprocess.run(command,cwd=workspace,text=True,capture_output=True)
        checks[name]['build_passed']=built.returncode==0
        checks[name]['build_output']=built.stdout+built.stderr
        if built.returncode:
            print(built.stdout+built.stderr)
            return built.returncode
    (workspace/'verification.json').write_text(json.dumps(checks,indent=2),encoding='utf-8')
    print(json.dumps(checks,indent=2))
    return 0


if __name__=='__main__':
    raise SystemExit(main())
