import json, sys, glob, collections, pathlib
SP = pathlib.Path(sys.argv[1]); rows = json.load(open(SP / 'corpus' / 'population.json'))
def load(exp):
    R = collections.defaultdict(dict)
    for f in glob.glob(str(SP / 'corpus' / exp / 'w*' / 'results.jsonl')):
        for line in open(f):
            j = json.loads(line); R[j['idx']][j['variant']] = j.get('pages')
    return R
def bucket(d):
    if d is None: return 'error'
    return 'exact' if d == 0 else '+1' if d == 1 else '+2..' if d >= 2 else '-1' if d == -1 else '-2..'
def table(exp, variants, base, stratum):
    R = load(exp); print(f'\n### {exp}: n documents = {len(R)}')
    groups = collections.defaultdict(list)
    for idx, vs in R.items(): groups[stratum(rows[idx])].append((rows[idx], vs))
    for g, items in sorted(groups.items()):
        print(f'  stratum {g}: n={len(items)}')
        for v in variants:
            c = collections.Counter(bucket(vs.get(v) - r['pages'] if vs.get(v) else None) for r, vs in items)
            line = '  '.join(f'{k}:{c.get(k,0)}' for k in ('exact', '+1', '+2..', '-1', '-2..', 'error'))
            tr = ''
            if v != base:
                t = collections.Counter()
                for r, vs in items:
                    a, b = vs.get(base), vs.get(v)
                    if a is None or b is None: t['n/a'] += 1; continue
                    da, db = abs(a - r['pages']), abs(b - r['pages'])
                    t['closer' if db < da else 'farther' if db > da else 'same'] += 1
                tr = f'   vs {base}: ' + ' '.join(f'{k}:{t.get(k,0)}' for k in ('closer', 'same', 'farther', 'n/a'))
            print(f'    {v:9} {line}{tr}')
if (SP / 'corpus' / 'c1').exists():
    table('c1', ['v0', 'v1', 'v3gm', 'v4adj', 'v5minrow', 'v6docx'], 'v1', lambda r: r['grid'])
if (SP / 'corpus' / 'c2').exists():
    table('c2', ['v0', 'v1adj', 'v2gm', 'v3both'], 'v0', lambda r: ('grid' if r['grid'] != 'none' else 'nogrid') + '/' + r['fam'])
if (SP / 'corpus' / 'c1r').exists():
    R1 = load('c1')
    table('c1r', ['v1r', 'v3r', 'v4r', 'v5r', 'v6r'], 'v1r', lambda r: r['grid'])
    R = load('c1r'); agree = sum(1 for i in R if R[i].get('v1r') is not None and R1.get(i, {}).get('v1') == R[i]['v1r'])
    print(f'  sanity: resaved v1r page count equals non-resaved v1 in {agree}/{sum(1 for i in R if R[i].get("v1r") is not None)} documents')
if (SP / 'corpus' / 'c1n').exists():
    table('c1n', ['n0e', 'n0', 'n4', 'n4e'], 'n0e', lambda r: r['grid'])
    R1 = load('c1'); R = load('c1n')
    agree = sum(1 for i in R if R[i].get('n0e') is not None and R1.get(i, {}).get('v1') == R[i]['n0e'])
    print(f'  control: n0e (settings processed, ext leading true) equals c1 v1 in {agree}/{sum(1 for i in R if R[i].get("n0e") is not None)}')
