"""Static checks for the public package; not a VBA compiler or a full DLP tool."""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
ALLOWED = {
    'README.md', '.gitignore', 'docs/SETUP.md', 'docs/ADOBE_WORKFLOW.md',
    'docs/PRIVACY_REVIEW.md', 'src/PBI_UI_Export.bas',
    'src/PBI_Example_Control.bas', 'tools/check_release.py',
}

def findings(text):
    issues = []
    if re.search(r'\\{2}[A-Za-z0-9_.-]+\\', text):
        issues.append('UNC/network path')
    if re.search(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', text):
        issues.append('email address')
    if re.search(r'(?i)(?:gh[pousr]_[A-Za-z0-9]{20,}|sk-[A-Za-z0-9]{20,})', text) or re.search(r'^-----BEGIN .*PRIVATE KEY-----', text, re.M):
        issues.append('credential-like string')
    for drive, tail in re.findall(r'([A-Za-z]):\\([^\s"\x27`<>]*)', text):
        if drive != 'C' or not (tail == 'abcd' or tail.startswith('abcd\\')):
            issues.append('non-example Windows path')
    return issues

def main():
    # A few synthetic negative controls prove these patterns are exercised.
    assert findings(chr(92)*2 + 'sample-host' + chr(92) + 'sample-share')
    assert findings('person' + '@' + 'example.invalid')
    assert findings('D:' + chr(92) + 'private-folder')
    assert not findings('C:' + chr(92) + 'abcd' + chr(92) + 'output')
    errors = []
    seen = set()
    for p in ROOT.rglob('*'):
        rel = p.relative_to(ROOT).as_posix()
        if '.git' in p.parts or '__pycache__' in p.parts or p.is_dir():
            continue
        seen.add(rel)
        if p.is_symlink() or rel not in ALLOWED:
            errors.append(f'{rel}: unexpected file')
            continue
        for issue in findings(p.read_text(encoding='utf-8')):
            errors.append(f'{rel}: {issue}')
    errors += [f'{p}: missing' for p in ALLOWED - seen]
    source = (ROOT / 'src/PBI_UI_Export.bas').read_text()
    assert 'Private Const PBI_VerboseSnapshot As Boolean = False' in source
    assert 'If tabs.Count > 50 Then' in source
    assert 'PBI_DeleteSuccessfulLog' in source
    assert '.Visible = xlSheetHidden' in source
    for name in ('Page A', 'Page B', 'Page C', 'Volume A', 'Volume B', 'Volume C'):
        assert name in source
    procedures = re.findall(r'^(?:Public|Private) (?:Sub|Function) (\w+)', source, re.M)
    endings = re.findall(r'^End (?:Sub|Function)$', source, re.M)
    assert len(procedures) == len(set(procedures)) == len(endings)
    if errors:
        raise SystemExit('\n'.join(errors))
    print(f'PASS: {len(seen)} allowlisted text files; privacy patterns and static module checks passed.')
    print('Windows/VBA compilation and interactive execution remain untested.')

if __name__ == '__main__':
    main()
