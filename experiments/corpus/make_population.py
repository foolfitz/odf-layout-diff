# SPDX-License-Identifier: MPL-2.0
"""Builds corpus/population.json for corpus_run.py / corpus_run_r.py / corpus_agg.py.

    python3 make_population.py <work-dir> <corpus-dir> [<corpus-dir> ...]

One row per readable .odt: path, generator family (word / moda / ndc / lo / other),
meta:page-count, page-grid modes, whether Word-import style names are present, and
whether settings.xml already carries AdjustTableLineHeightsToGridHeight / MsWordCompGridMetrics.
The output holds local paths: keep it in a scratch directory, never in a repository.
"""
import glob
import json
import pathlib
import re
import sys
import zipfile

work = pathlib.Path(sys.argv[1])
(work / 'corpus').mkdir(parents=True, exist_ok=True)
files = [f for d in sys.argv[2:] for f in glob.glob(f'{d}/**/*.odt', recursive=True)]
rows = []
for f in files:
    try:
        z = zipfile.ZipFile(f)
        meta = z.read('meta.xml').decode('utf8', 'replace')
        styles = z.read('styles.xml').decode('utf8', 'replace')
        content = z.read('content.xml').decode('utf8', 'replace')
        settings = z.read('settings.xml').decode('utf8', 'replace') if 'settings.xml' in z.namelist() else ''
    except Exception:
        continue
    g = re.search(r'<meta:generator>(.*?)<', meta)
    g = g.group(1) if g else ''
    pc = re.search(r'meta:page-count="(\d+)"', meta)
    modes = set(re.findall(r'style:layout-grid-mode="(\w+)"', styles)) - {'none'}
    fam = ('word' if g.startswith('MicrosoftOffice') else 'moda' if 'MODA_ODF' in g else 'ndc' if 'NDC_ODF' in g
           else 'lo' if ('LibreOffice' in g or 'OxOffice' in g) else 'other')
    rows.append(dict(
        f=f, gen=g[:60], fam=fam, pages=int(pc.group(1)) if pc else None,
        grid='+'.join(sorted(modes)) or 'none',
        bh='style:layout-grid-base-height=' in styles,
        bw=re.findall(r'style:layout-grid-base-width="([^"]+)"', styles)[:1],
        ww=bool(re.search(r'WW8Num|"WW-|WW_|List_20_Paragraph', content + styles)),
        agh='AdjustTableLineHeightsToGridHeight' in settings,
        gm='MsWordCompGridMetrics' in settings,
        nbytes=len(content)))
json.dump(rows, open(work / 'corpus' / 'population.json', 'w'), ensure_ascii=False)
print(len(rows), 'rows')
