# This Source Code Form is subject to the terms of the Mozilla Public License, v. 2.0.
# If a copy of the MPL was not distributed with this file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Checks the Markdown docs without Excel, so it runs in CI (.github/workflows/
# docs-check.yml) and on any machine with Python 3:
#
#   python3 .github/scripts/check-docs.py
#
# 1. Links and images: every relative link in wiki/ names a wiki page (and,
#    with #anchor, a heading on it); every relative link elsewhere names a file
#    in the repo; every image exists; links to this repo's master tree or wiki
#    on github.com resolve the same way.
# 2. Sidebar: every wiki page is linked from wiki/_Sidebar.md.
# 3. Map: every btn* macro in src/PrintCosts-VBA and every SET_* name (the
#    template's defined names plus those the VBA creates) has a row in
#    wiki/_map.md, and every page the map names exists.
#
# Buttons are created by the VBA at build time, so (3) reads the source, not a
# built workbook. Prints one line per problem and exits 1 if there are any.

import glob
import os
import re
import sys
import zipfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
WIKI = os.path.join(ROOT, 'wiki')
VBA = os.path.join(ROOT, 'src', 'PrintCosts-VBA')
TEMPLATE = os.path.join(ROOT, 'src', 'PrintCosts.xlsx')
REPO_URL = 'https://github.com/Adasha/PrintJob'
SKIP_DIRS = {'.git', 'node_modules', 'wiki'}

# Macros with no control of their own to document.
MAP_EXEMPT = {
    'btnPanelNoop': 'assigned to panel captions so a click does nothing',
    'btnPickerOK': 'OK on the _Picker sheet, covered by Select printers',
    'btnPickerCancel': 'Cancel on the _Picker sheet, covered by Select printers',
}

problems = []


def problem(path, msg):
    problems.append(f'{os.path.relpath(path, ROOT)}: {msg}')


def read(path):
    with open(path, encoding='utf-8') as f:
        return f.read()


def strip_code(text):
    # Links inside fenced blocks and code spans are examples, not links.
    text = re.sub(r'^```.*?^```', '', text, flags=re.M | re.S)
    return re.sub(r'`[^`\n]*`', '', text)


def slugs(text):
    # GitHub's heading anchors: lower case, punctuation dropped, spaces to
    # hyphens, repeats numbered -1, -2...
    seen = {}
    out = set()
    for h in re.findall(r'^#{1,6}\s+(.+?)\s*#*\s*$', strip_code(text), re.M):
        s = re.sub(r'[^\w\- ]', '', h.lower()).replace(' ', '-')
        n = seen.get(s, 0)
        seen[s] = n + 1
        out.add(s if n == 0 else f'{s}-{n}')
    return out


LINK_RE = re.compile(r'(!?)\[[^\]]*\]\(\s*<?([^)\s>]+)>?(?:\s+"[^"]*")?\s*\)')
IMG_TAG_RE = re.compile(r'<img\s[^>]*src="([^"]+)"', re.I)

wiki_pages = {os.path.splitext(os.path.basename(p))[0]: p
              for p in glob.glob(os.path.join(WIKI, '*.md'))}


def check_wiki_target(path, target):
    page, _, anchor = target.partition('#')
    page = page[:-3] if page.endswith('.md') else page
    if page and page not in wiki_pages:
        problem(path, f'link to missing wiki page "{target}"')
        return
    if anchor:
        dest = wiki_pages[page] if page else path
        if anchor.lower() not in slugs(read(dest)):
            problem(path, f'link to missing heading "{target}"')


def check_repo_target(path, base, target):
    rel, _, anchor = target.partition('#')
    dest = os.path.normpath(os.path.join(base, rel)) if rel else path
    if not os.path.exists(dest):
        problem(path, f'link to missing file "{target}"')
    elif anchor and dest.endswith('.md') and anchor.lower() not in slugs(read(dest)):
        problem(path, f'link to missing heading "{target}"')


def check_links(path, in_wiki):
    text = strip_code(read(path))
    base = os.path.dirname(path)
    for bang, target in LINK_RE.findall(text):
        if target.startswith(REPO_URL + '/wiki'):
            check_wiki_target(path, target[len(REPO_URL + '/wiki'):].lstrip('/') or 'Home')
        elif re.match(re.escape(REPO_URL) + r'/(tree|blob)/master/', target):
            check_repo_target(path, ROOT, target.split('/master/', 1)[1])
        elif re.match(r'[a-z]+:', target):
            continue  # other sites, mailto:
        elif bang or not in_wiki:
            check_repo_target(path, base, target)
        else:
            check_wiki_target(path, target)
    for src in IMG_TAG_RE.findall(text):
        if not re.match(r'[a-z]+:', src):
            check_repo_target(path, base, src)


# 1. Links and images
for page in wiki_pages.values():
    check_links(page, True)
for d, dirs, files in os.walk(ROOT):
    dirs[:] = [x for x in dirs if x not in SKIP_DIRS]
    for f in files:
        if f.endswith('.md'):
            check_links(os.path.join(d, f), False)

# 2. Sidebar
sidebar = os.path.join(WIKI, '_Sidebar.md')
linked = {t.partition('#')[0] for _, t in LINK_RE.findall(strip_code(read(sidebar)))}
for name, path in sorted(wiki_pages.items()):
    if not name.startswith('_') and name not in linked:
        problem(path, 'not linked from _Sidebar.md')

# 3. Map
map_path = os.path.join(WIKI, '_map.md')
map_text = read(map_path)
map_codes = set(re.findall(r'`([^`\n]+)`', map_text))
# `btnAddRow<Table>` stands for one macro per catalogue table.
map_patterns = [re.compile(re.escape(c).replace(re.escape('<Table>'), r'\w+') + '$')
                for c in map_codes if '<Table>' in c]

vba_files = glob.glob(os.path.join(VBA, '*'))
macros = set()
set_names = set()
for f in vba_files:
    src = read(f) if not f.endswith('.frx') else ''
    macros |= set(re.findall(r'^(?:Public |Private )?Sub (btn\w+)', src, re.M))
    set_names |= set(re.findall(r'"SET_([A-Z0-9_]+)"', src))
with zipfile.ZipFile(TEMPLATE) as z:
    set_names |= set(re.findall(r'definedName name="SET_(\w+)"', z.read('xl/workbook.xml').decode()))

for m in sorted(macros):
    if m in MAP_EXEMPT or m in map_codes or any(p.match(m) for p in map_patterns):
        continue
    problem(map_path, f'no row for macro {m}')
for k in sorted(set_names):
    if k not in map_codes and f'SET_{k}' not in map_codes:
        problem(map_path, f'no row for setting SET_{k}')
page_col = None
for line in map_text.splitlines():
    if not line.startswith('|'):
        page_col = None
        continue
    cells = [c.strip() for c in line.strip('|').split('|')]
    if page_col is None:
        page_col = cells.index('Page') if 'Page' in cells else -1
    elif page_col >= 0 and page_col < len(cells):
        # A cell may be free text ("the sheet's Setup page"); check page names only.
        for name in re.findall(r'\b[A-Z][a-z]+(?:-[A-Za-z]+)+\b', cells[page_col]):
            if name not in wiki_pages:
                problem(map_path, f'names missing page "{name}"')

for p in problems:
    print(p)
print(f'{len(problems)} problem(s): {len(wiki_pages)} wiki pages, {len(macros)} macros, '
      f'{len(set_names)} settings checked')
sys.exit(1 if problems else 0)
