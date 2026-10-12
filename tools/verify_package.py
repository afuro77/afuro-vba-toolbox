"""Static package checks only. Does not compile or execute Excel VBA."""
from pathlib import Path
import json, re

root = Path(__file__).resolve().parents[1]
modules = sorted((root / 'src/modules').glob('*.bas'))
classes = sorted((root / 'src/classes').glob('*.cls'))
extras = sorted((root / 'examples').glob('*.bas')) + sorted((root / 'tests').glob('*.bas'))
assert len(modules) == 11 and len(classes) == 3
reserved = {'single', 'double', 'integer', 'long', 'byte', 'boolean', 'string', 'date', 'variant', 'object'}

def code_only(line):
    """Remove VBA string literals and comments; validate doubled-quote escapes."""
    out = []; i = 0
    while i < len(line):
        char = line[i]
        if char == "'": break
        if char != '"':
            out.append(char); i += 1; continue
        i += 1; closed = False
        while i < len(line):
            if line[i] == '"':
                if i + 1 < len(line) and line[i + 1] == '"': i += 2; continue
                i += 1; closed = True; break
            i += 1
        assert closed, f'Unclosed VBA string: {line}'
        out.append('STRING_LITERAL')
    return ''.join(out)

procedure_pattern = re.compile(r'^(?:Public |Private |Friend )?(Sub|Function|Property (?:Get|Let|Set)) (\w+)\b', re.M)
global_names = {}
own_names = {}
for path in modules + classes + extras:
    text = path.read_text(encoding='utf-8')
    text.encode('cp932')
    assert f'Attribute VB_Name = "{path.stem}"' in text, path
    assert text.count('Option Explicit') == 1, path
    lines = text.splitlines()
    stripped = '\n'.join(code_only(line) for line in lines)
    starts = procedure_pattern.findall(stripped)
    ends = re.findall(r'^End (Sub|Function|Property)\s*$', stripped, re.M)
    assert len(starts) == len(ends), (path, 'Procedure boundaries')
    assert sum(kind == 'Sub' for kind, _ in starts) == ends.count('Sub'), path
    assert sum(kind == 'Function' for kind, _ in starts) == ends.count('Function'), path
    assert sum(kind.startswith('Property') for kind, _ in starts) == ends.count('Property'), path
    own_names[path.stem] = {name for _, name in starts} | set(re.findall(r'^Public Type (\w+)', text, re.M))
    if path in modules:
        for name in own_names[path.stem]:
            assert name.lower() not in global_names, (path, name, 'Global name collision')
            global_names[name.lower()] = path.stem
    for i, line in enumerate(lines):
        assert len(line.encode('cp932')) <= 1023, (path, i + 1, 'Physical line too long')
        if line.rstrip().endswith(' _'):
            assert i + 1 < len(lines) and lines[i + 1].strip(), (path, i + 1)
        names = re.findall(r'\b(?:ByVal|ByRef|Dim|Private)\s+(\w+)\s+As\b', code_only(line), re.I)
        assert not any(name.lower() in reserved for name in names), (path, i + 1, 'Reserved type used as identifier')
    assert stripped.count('(') == stripped.count(')'), (path, 'Unbalanced parentheses')
    assert not re.search(r'\.(Select|Activate)\b|\bActiveSheet\b|\bActiveWorkbook\b', stripped), (path, 'Implicit selection dependency')

dependencies = {'mdlAT_Csv': {'mdlAT_Text', 'mdlAT_Range', 'mdlAT_Files'}}
for path in modules:
    names = own_names[path.stem]
    calls = set(re.findall(r'\bAT_\w+\b', '\n'.join(code_only(line) for line in path.read_text(encoding='utf-8').splitlines())))
    for name in calls - names:
        assert name.lower() in global_names, (path, name, 'Unknown standard-module dependency')
        assert global_names[name.lower()] in dependencies.get(path.stem, set()), (path, name, 'Undocumented dependency')

for directory in ['src/modules', 'src/classes', 'src/all-in-one', 'examples', 'tests']:
    for path in (root / directory).glob('*'):
        if path.suffix not in {'.bas', '.cls'}: continue
        relative = path.relative_to(root)
        target = root / 'vbe' / Path(*relative.parts[1:]) if relative.parts[0] == 'src' else root / 'vbe' / relative
        expected = path.read_text(encoding='utf-8').replace('\n', '\r\n').encode('cp932')
        assert target.read_bytes() == expected, (path, 'Import encoding mismatch')

all_text = (root / 'src/all-in-one/mdlAT_All.bas').read_text(encoding='utf-8')
all_names = {name.lower() for _, name in procedure_pattern.findall(all_text)}
assert all_names == {name.lower() for path in modules for kind, name in procedure_pattern.findall(path.read_text(encoding='utf-8'))}, 'All-in-one procedure mismatch'
assert all_text.index('Public Type') < all_text.index('Public Function'), 'Declarations must precede procedures'
for path in modules:
    text = path.read_text(encoding='utf-8')
    for body in re.findall(r'^(?:Public |Private )?(?:Sub|Function) \w+.*?^End (?:Sub|Function)\s*$', text, re.M | re.S):
        assert body.strip() in all_text, (path, 'All-in-one body mismatch')

catalog = json.loads((root / 'docs/catalog.json').read_text(encoding='utf-8'))
assert len(catalog) == 100
assert (root / '部品一覧.html').read_text(encoding='utf-8').count('<article ') == len(catalog)
for item in catalog:
    assert (root / item['path']).is_file()
    assert item['name'] in own_names[item['module']]

for path in root.rglob('*'):
    if not path.is_file() or '__pycache__' in path.parts: continue
    assert path.suffix.lower() not in {'.xls', '.xlsx', '.xlsm', '.xlsb', '.bin'}, path
    if 'vbe' in path.parts: continue
    text = path.read_text(encoding='utf-8')
    if path.name == 'verify_package.py': continue
    assert not re.search(r'(?i)C:\\Users\\|sharepoint\.com|https?://|Auto_Open\s*\(', text), (path, 'Publication-sensitive identifier')

print(f'PASS: {len(modules)} modules, {len(classes)} classes, 100 indexed parts, encodings, literals, boundaries, names, dependencies and generated copies.')
print('VBE compilation and Excel execution: NOT RUN.')
