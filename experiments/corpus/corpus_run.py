#!/usr/bin/env python3
# SPDX-License-Identifier: MPL-2.0
# Copyright 2026 OSSII
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
"""Corpus page-count experiments (pre-registered before running).

C1  population: every ODT whose meta:generator starts with MicrosoftOffice and has
    meta:page-count (Word's own layout).  Reference = that page count.
    variants (all with non-ASCII fixed-pitch dropped, prepare_word_odt.drop_fixed_pitch):
      v0        nothing else
      v1        + text grid filled (prepare_word_odt.fill_grid)
      v3gm      v1 + MsWordCompGridMetrics=true
      v4adj     v1 + AdjustTableLineHeightsToGridHeight=false
      v5minrow  v1 + MinRowHeightInclBorder=true
C2  population: ODTs saved by MODA/NDC ODF tools or LibreOffice that carry Word-import
    style names (WW8Num / WW- / WW_ / List_20_Paragraph); 300 random with a text grid,
    100 random without (negative control).  Reference = meta:page-count of the saving app.
    variants: v0 (font fix), v1adj (+AdjustTableLineHeightsToGridHeight=false),
              v2gm (+MsWordCompGridMetrics=true), v3both (both).
Counting object: one document x variant -> rendered page count (pdfinfo) or an error.
"""
import json, os, re, shutil, signal, subprocess, sys, zipfile, random, pathlib, concurrent.futures
sys.path.insert(0, str(pathlib.Path(__file__).parent.parent)); sys.path.insert(0, str(pathlib.Path(__file__).parent))
from prepare_word_odt import fill_grid, drop_fixed_pitch

# The renderer is the oracle for every page count below, so it is resolved once,
# up front, and never guessed: $SOFFICE wins, otherwise whatever is on PATH.
# A missing binary stops the run here rather than after the corpus was built.
SOFFICE = os.environ.get('SOFFICE') or shutil.which('soffice')
if not SOFFICE or not os.path.isfile(SOFFICE):
    raise SystemExit(f'soffice not found ({SOFFICE or "not on PATH"}): set $SOFFICE or put it on PATH')

SP = pathlib.Path(sys.argv[1]); EXP = sys.argv[2]; WORKERS = 4
WORK = SP / 'corpus' / EXP; WORK.mkdir(parents=True, exist_ok=True)
rows = json.load(open(SP / 'corpus' / 'population.json'))
for i, r in enumerate(rows): r['idx'] = i
random.seed(20260916)
if EXP == 'c1':
    sel = [r for r in rows if r['fam'] == 'word' and r['pages']]
    VARIANTS = {'v0': {}, 'v1': {}, 'v3gm': {'MsWordCompGridMetrics': 'true'},
                'v4adj': {'AdjustTableLineHeightsToGridHeight': 'false'},
                'v5minrow': {'MinRowHeightInclBorder': 'true'},
                'v6docx': {'MsWordCompGridMetrics': 'true', 'MinRowHeightInclBorder': 'true'}}
    GRIDFILL = {'v1', 'v3gm', 'v4adj', 'v5minrow', 'v6docx'}
elif EXP == 'c1n':
    sel = [r for r in rows if r['fam'] == 'word' and r['pages']]
    VARIANTS = {'n0e': {'AddExternalLeading': 'true'}, 'n0': {'IsLabelDocument': 'false'},
                'n4': {'AdjustTableLineHeightsToGridHeight': 'false'},
                'n4e': {'AdjustTableLineHeightsToGridHeight': 'false', 'AddExternalLeading': 'true'}}
    GRIDFILL = set(VARIANTS)
else:
    fam = [r for r in rows if r['fam'] in ('moda', 'lo', 'ndc') and r['ww'] and r['pages']]
    grid = [r for r in fam if r['grid'] != 'none']; none = [r for r in fam if r['grid'] == 'none']
    sel = random.sample(grid, min(300, len(grid))) + random.sample(none, min(100, len(none)))
    VARIANTS = {'v0': {}, 'v1adj': {'AdjustTableLineHeightsToGridHeight': 'false'},
                'v2gm': {'MsWordCompGridMetrics': 'true'},
                'v3both': {'AdjustTableLineHeightsToGridHeight': 'false', 'MsWordCompGridMetrics': 'true'}}
    GRIDFILL = set()
json.dump([r['idx'] for r in sel], open(WORK / 'selection.json', 'w'))

from inject import with_items

def build(r, variant, dst):
    with zipfile.ZipFile(r['f']) as zi, zipfile.ZipFile(dst, 'w') as zo:
        names = zi.namelist()
        if VARIANTS[variant] and 'settings.xml' not in names: raise RuntimeError('no-settings')
        for n in sorted(names, key=lambda x: x != 'mimetype'):
            d = zi.read(n)
            if n in ('styles.xml', 'content.xml'):
                t = drop_fixed_pitch(d.decode('utf-8'))[0]
                if n == 'styles.xml' and variant in GRIDFILL: t = fill_grid(t)[0]
                d = t.encode('utf-8')
            elif n == 'settings.xml' and VARIANTS[variant]:
                d = with_items(d.decode('utf-8'), VARIANTS[variant]).encode('utf-8')
            zo.writestr(zipfile.ZipInfo(n, (1980, 1, 1, 0, 0, 0)), d, zipfile.ZIP_STORED if n == 'mimetype' else zipfile.ZIP_DEFLATED)

def soffice(profile, outdir, files, timeout):
    p = subprocess.Popen([SOFFICE, f'-env:UserInstallation=file://{profile}', '--headless', '--convert-to', 'pdf', '--outdir', str(outdir)] + [str(f) for f in files],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    try: p.wait(timeout=timeout); return True
    except subprocess.TimeoutExpired:
        os.killpg(p.pid, signal.SIGKILL); p.wait(); return False

def pages(pdf):
    o = subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True).stdout
    m = re.search(r'Pages:\s+(\d+)', o); return int(m.group(1)) if m else None

def worker(n, jobs):
    wd = WORK / f'w{n}'; ind, outd = wd / 'in', wd / 'out'; ind.mkdir(parents=True, exist_ok=True); outd.mkdir(exist_ok=True)
    profile = SP / f'profile-{EXP}-w{n}'; res = open(wd / 'results.jsonl', 'a')
    done = set()
    if (wd / 'results.jsonl').exists():
        for line in open(wd / 'results.jsonl'): j = json.loads(line); done.add((j['idx'], j['variant']))
    jobs = [j for j in jobs if (j[0]['idx'], j[1]) not in done]
    for b in range(0, len(jobs), 20):
        batch = []
        for r, v in jobs[b:b + 20]:
            f = ind / f"{r['idx']}__{v}.odt"
            try: build(r, v, f); batch.append((r, v, f))
            except Exception as e: res.write(json.dumps({'idx': r['idx'], 'variant': v, 'error': str(e)[:80]}) + '\n')
        soffice(profile, outd, [f for _, _, f in batch], 900)
        for r, v, f in batch:
            pdf = outd / (f.stem + '.pdf')
            if not pdf.exists(): soffice(profile, outd, [f], 180)
            p = pages(pdf) if pdf.exists() else None
            res.write(json.dumps({'idx': r['idx'], 'variant': v, 'pages': p} if p else {'idx': r['idx'], 'variant': v, 'error': 'no-pdf'}) + '\n')
            for x in (f, pdf):
                try: x.unlink()
                except FileNotFoundError: pass
        res.flush()

jobs = [(r, v) for r in sel for v in VARIANTS]
chunks = [jobs[i::WORKERS] for i in range(WORKERS)]
with concurrent.futures.ThreadPoolExecutor(WORKERS) as ex:
    list(ex.map(worker, range(WORKERS), chunks))
print(EXP, 'done', len(jobs), 'jobs')
