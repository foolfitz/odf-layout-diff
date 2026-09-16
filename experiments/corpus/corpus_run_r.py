#!/usr/bin/env python3
# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""C1r (pre-registered): C1 population (Word-exported ODT with meta page count = Word's layout).
Base = non-ASCII fixed pitch dropped + text grid filled, then saved once by LibreOffice 26.2
(its settings.xml is complete; compat items injected into Word's minimal settings.xml are ignored).
variants from the resaved file: v1r (as is), v3r +MsWordCompGridMetrics=true,
v4r +AdjustTableLineHeightsToGridHeight=false, v5r +MinRowHeightInclBorder=true, v6r v3r+v5r.
Counting object: document x variant -> rendered page count or error; reference = meta:page-count."""
import json, os, re, signal, subprocess, sys, zipfile, pathlib, concurrent.futures, shutil
sys.path.insert(0, str(pathlib.Path(__file__).parent.parent)); sys.path.insert(0, str(pathlib.Path(__file__).parent))
from prepare_word_odt import fill_grid, drop_fixed_pitch
from inject import with_items

# Resolved once, never guessed: see corpus_run.py for why.
SOFFICE = os.environ.get('SOFFICE') or shutil.which('soffice')
if not SOFFICE or not os.path.isfile(SOFFICE):
    raise SystemExit(f'soffice not found ({SOFFICE or "not on PATH"}): set $SOFFICE or put it on PATH')

SP = pathlib.Path(sys.argv[1]); WORKERS = 4; WORK = SP / 'corpus' / 'c1r'; WORK.mkdir(parents=True, exist_ok=True)
rows = json.load(open(SP / 'corpus' / 'population.json'))
sel = [dict(r, idx=i) for i, r in enumerate(rows) if r['fam'] == 'word' and r['pages']]
VARIANTS = {'v1r': {}, 'v3r': {'MsWordCompGridMetrics': 'true'}, 'v4r': {'AdjustTableLineHeightsToGridHeight': 'false'},
            'v5r': {'MinRowHeightInclBorder': 'true'}, 'v6r': {'MsWordCompGridMetrics': 'true', 'MinRowHeightInclBorder': 'true'}}
def rezip(src, dst, edit):
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, 'w') as zo:
        for n in sorted(zi.namelist(), key=lambda x: x != 'mimetype'):
            d = edit(n, zi.read(n))
            zo.writestr(zipfile.ZipInfo(n, (1980, 1, 1, 0, 0, 0)), d, zipfile.ZIP_STORED if n == 'mimetype' else zipfile.ZIP_DEFLATED)
def pre(n, d):
    if n in ('styles.xml', 'content.xml'):
        t = drop_fixed_pitch(d.decode('utf-8'))[0]
        if n == 'styles.xml': t = fill_grid(t)[0]
        return t.encode('utf-8')
    return d
def soffice(profile, fmt, outdir, files, timeout):
    p = subprocess.Popen([SOFFICE, f'-env:UserInstallation=file://{profile}', '--headless', '--convert-to', fmt, '--outdir', str(outdir)] + [str(f) for f in files],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    try: p.wait(timeout=timeout)
    except subprocess.TimeoutExpired: os.killpg(p.pid, signal.SIGKILL); p.wait()
def pages(pdf):
    m = re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True).stdout); return int(m.group(1)) if m else None
def worker(n, docs):
    wd = WORK / f'w{n}'
    for sub in ('pre', 'res', 'var', 'out'): (wd / sub).mkdir(parents=True, exist_ok=True)
    profile = SP / f'profile-c1r-w{n}'; res = open(wd / 'results.jsonl', 'a')
    for b in range(0, len(docs), 10):
        batch = docs[b:b + 10]; ok = []
        for r in batch:
            f = wd / 'pre' / f"{r['idx']}.odt"
            try: rezip(r['f'], f, pre); ok.append((r, f))
            except Exception as e: res.write(json.dumps({'idx': r['idx'], 'variant': 'all', 'error': 'pre:' + str(e)[:60]}) + '\n')
        soffice(profile, 'odt', wd / 'res', [f for _, f in ok], 900)
        vars_ = []
        for r, f in ok:
            rs = wd / 'res' / f.name
            if not rs.exists(): res.write(json.dumps({'idx': r['idx'], 'variant': 'all', 'error': 'resave'}) + '\n'); continue
            for v, items in VARIANTS.items():
                dst = wd / 'var' / f"{r['idx']}__{v}.odt"
                try:
                    rezip(rs, dst, lambda nm, d: with_items(d.decode('utf-8'), items).encode('utf-8') if (nm == 'settings.xml' and items) else d); vars_.append((r, v, dst))
                except Exception as e: res.write(json.dumps({'idx': r['idx'], 'variant': v, 'error': 'inject:' + str(e)[:60]}) + '\n')
        soffice(profile, 'pdf', wd / 'out', [d for _, _, d in vars_], 900)
        for r, v, dst in vars_:
            pdf = wd / 'out' / (dst.stem + '.pdf'); p = pages(pdf) if pdf.exists() else None
            res.write(json.dumps({'idx': r['idx'], 'variant': v, 'pages': p} if p else {'idx': r['idx'], 'variant': v, 'error': 'no-pdf'}) + '\n')
        res.flush()
        for sub in ('pre', 'res', 'var', 'out'): shutil.rmtree(wd / sub); (wd / sub).mkdir()
with concurrent.futures.ThreadPoolExecutor(WORKERS) as ex:
    list(ex.map(worker, range(WORKERS), [sel[i::WORKERS] for i in range(WORKERS)]))
print('c1r done', len(sel))
