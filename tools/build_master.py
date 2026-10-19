"""master-src/ のファイル（とURL）から、アプリが取得する master.json / master-meta.json を作る。

対応ファイル: .xlsx .xls .csv .txt .tsv .zip（中身を展開して処理）
  - 厚生労働省「薬価基準収載品目リスト」(Excel)          → 品名・薬価基準収載医薬品コード・成分名・規格・メーカー・後発品区分
  - MEDIS「医薬品HOTコードマスター」(TXT/CSV)            → 販売名・YJコード・HOTコード・JAN
  - GS1（販売包装単位／調剤包装単位）コードの対応表      → HOTコード経由で薬品名と結びつける
列は見出し、または中身の形から自動で判定する（アプリ内の取り込みと同じ考え方）。

使い方: python tools/build_master.py [URL ...]
"""
import hashlib, io, json, re, sys, unicodedata, urllib.request, zipfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'master-src'
YJ_RE = re.compile(r'^\d{7}[A-Z]\d{3}[0-9A-Z]$')
GTIN_RE = re.compile(r'^(?:\d?)(?:45|49)\d{11}$')
FORM_RE = re.compile(r'(錠|カプセル|散|顆粒|液|軟膏|クリーム|テープ|注|mg|%|g)')
KATA_RE = re.compile(r'[ァ-ヴ]{2,}')


def cell(v):
    if v is None:
        return ''
    s = str(v)
    if s.lower() == 'nan':
        return ''
    if re.fullmatch(r'\d+\.0', s):  # Excelの数値化されたコード
        s = s[:-2]
    return unicodedata.normalize('NFKC', s).strip()


def is_gtin(v):
    return bool(GTIN_RE.fullmatch(v)) and len(v) in (13, 14)


def is_hot(v):
    return bool(re.fullmatch(r'\d{9}|\d{13}', v)) and not is_gtin(v)


def norm_key(name):
    s = unicodedata.normalize('NFKC', name)
    s = ''.join(chr(ord(c) + 0x60) if 'ぁ' <= c <= 'ゖ' else c for c in s)
    s = re.sub(r'【般】|\[般\]|\(般\)', '', s)
    return re.sub(r'[\s・「」『』"\'“”]', '', s).upper()


# ---------- 読み込み ----------
def decode(b):
    for enc in ('utf-8-sig', 'cp932'):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            pass
    return b.decode('cp932', errors='replace')


def tables_from_bytes(name, data):
    low = name.lower()
    if low.endswith('.zip'):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                inner = info.filename
                try:
                    inner = inner.encode('cp437').decode('cp932')
                except Exception:
                    pass
                yield from tables_from_bytes(inner, z.read(info))
        return
    if low.endswith(('.xlsx', '.xlsm', '.xls')):
        import pandas as pd
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str)
        for sheet in sheets.values():
            yield name, [[cell(v) for v in row] for row in sheet.itertuples(index=False)]
        return
    if low.endswith(('.csv', '.txt', '.tsv')):
        import csv
        text = decode(data)
        head = text[:4000]
        delim = '\t' if head.count('\t') > head.count(',') else ','
        yield name, [[cell(v) for v in row] for row in csv.reader(io.StringIO(text), delimiter=delim)]


# ---------- 列の判定 ----------
def detect(rows):
    for r, row in enumerate(rows[:15]):
        def find(pat, bad=None):
            for i, c in enumerate(row):
                if len(c) < 40 and re.search(pat, c) and not (bad and re.search(bad, c)):
                    return i
            return -1
        name = next((i for i in (find(r'販売名', r'カナ|コード'), find(r'医薬品名|製品名|品名|薬品名|商品名', r'カナ|コード'), find(r'告示名称')) if i >= 0), -1)
        yj = find(r'YJ|個別医薬品コード')
        yaka = find(r'薬価基準収載医薬品コード|薬価コード|薬価基準コード')
        gs1 = [i for i, c in enumerate(row) if len(c) < 40 and re.search(r'GS1|調剤包装単位コード|販売包装単位コード|元梱包装単位コード|JAN|GTIN', c)]
        hot = find(r'基準番号|HOT')
        if (name >= 0 and (yj >= 0 or yaka >= 0 or gs1)) or (hot >= 0 and gs1):
            return dict(start=r + 1, name=name, yj=yj if yj >= 0 else yaka, gs1=gs1, hot=hot,
                        ingredient=find(r'^成分名'), spec=find(r'^規格'), maker=find(r'メーカー|製造販売'),
                        ge=find(r'後発医薬品$|^後発品'), brand=find(r'^先発医薬品$'), price=find(r'^薬価$'))
    sample = rows[:400]
    width = max((len(r) for r in sample), default=0)

    def ratio(i, test):
        return sum(1 for r in sample if i < len(r) and r[i] and test(r[i])) / max(1, len(sample))
    col = dict(start=0, name=-1, yj=-1, gs1=[], hot=-1, ingredient=-1, spec=-1, maker=-1, ge=-1, brand=-1, price=-1)
    best = 0
    for i in range(width):
        if col['yj'] < 0 and ratio(i, lambda v: bool(YJ_RE.fullmatch(v))) > 0.5:
            col['yj'] = i
        if ratio(i, is_gtin) > 0.5:
            col['gs1'].append(i)
        if col['hot'] < 0 and ratio(i, lambda v: bool(re.fullmatch(r'\d{13}', v)) and not is_gtin(v)) > 0.5:
            col['hot'] = i
        ns = ratio(i, lambda v: bool(KATA_RE.search(v)) and bool(FORM_RE.search(v)))
        if ns > best + 0.05:
            best, col['name'] = ns, i
    if best < 0.3:
        col['name'] = -1
    return col


# ---------- 組み立て ----------
def build(sources):
    items, key_index, hot_index, pending = [], {}, {}, []
    rows_read = 0

    def add_unique(lst, v):
        if v and v not in lst:
            lst.append(v)

    for fname, data in sources:
        for tname, rows in tables_from_bytes(fname, data):
            col = detect(rows)
            if col['name'] < 0 and not (col['hot'] >= 0 and col['gs1']):
                print(f'  skip: {tname}（列を判定できません）')
                continue
            print(f'  read: {tname}  name={col["name"]} yj={col["yj"]} hot={col["hot"]} gs1={col["gs1"]}')
            for row in rows[col['start']:]:
                g = lambda i: row[i] if 0 <= i < len(row) else ''
                name = g(col['name'])
                yj = g(col['yj'])
                if not YJ_RE.fullmatch(yj):
                    yj = ''
                gs = [v.zfill(14) for v in (g(i) for i in col['gs1']) if re.fullmatch(r'\d{13,14}', v)]
                hot = g(col['hot']) if is_hot(g(col['hot'])) else ''
                rows_read += 1
                if name and re.search(r'[ァ-ヴA-Za-z一-龥]', name):
                    k = (yj or 'N') + '|' + norm_key(name)
                    idx = key_index.get(k)
                    if idx is None:
                        idx = len(items)
                        items.append({'n': name, 'y': yj, 'g': [], 'h': []})
                        key_index[k] = idx
                    it = items[idx]
                    for v in gs:
                        add_unique(it['g'], v)
                    if hot:
                        add_unique(it['h'], hot)
                        hot_index[hot] = idx
                        hot_index[hot[:9]] = idx
                    for key, c in (('c', 'ingredient'), ('s', 'spec'), ('m', 'maker'), ('p', 'price')):
                        v = g(col[c])
                        if v and key not in it:
                            it[key] = v
                    gev, brv = g(col['ge']), g(col['brand'])
                    if gev and gev not in ('-', '－'):
                        it['ge'] = 1
                    if brv and brv not in ('-', '－'):
                        it['br'] = 1
                elif hot and gs:
                    pending.append((hot, gs))
    linked = 0
    for hot, gs in pending:
        idx = hot_index.get(hot, hot_index.get(hot[:9]))
        if idx is None:
            continue
        for v in gs:
            add_unique(items[idx]['g'], v)
        linked += 1
    for it in items:
        if not it['g']:
            del it['g']
        if not it['h']:
            del it['h']
    return items, rows_read, linked, len(pending) - linked


def main():
    sources = []
    for p in sorted(SRC.glob('*')):
        if p.is_file() and not p.name.startswith('.') and p.suffix.lower() in ('.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.tsv', '.zip'):
            sources.append((p.name, p.read_bytes()))
    for url in [u for a in sys.argv[1:] for u in a.split()]:
        print('download:', url)
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=120) as r:
            sources.append((url.rsplit('/', 1)[-1].split('?')[0] or 'download', r.read()))
    if not sources:
        print('master-src/ にファイルがありません')
        return 1
    items, rows_read, linked, unlinked = build(sources)
    if not items:
        print('薬品を1件も読み取れませんでした')
        return 1
    items.sort(key=lambda it: (it.get('y') or 'ZZZ', it['n']))
    body = json.dumps({'items': items}, ensure_ascii=False, separators=(',', ':'))
    version = hashlib.sha1(body.encode()).hexdigest()[:12]
    jst = timezone(timedelta(hours=9))
    meta = {
        'version': version,
        'date': datetime.now(jst).strftime('%Y/%m/%d'),
        'count': len(items),
        'gs1': sum(len(it.get('g', [])) for it in items),
        'files': [n for n, _ in sources],
    }
    (ROOT / 'master.json').write_text(json.dumps({'meta': meta, 'items': items}, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    (ROOT / 'master-meta.json').write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding='utf-8')
    print(f'行 {rows_read:,} → 品目 {len(items):,}（GS1 {meta["gs1"]:,}件, 対応表 {linked:,}行を結合, 未結合 {unlinked:,}行） version={version}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
