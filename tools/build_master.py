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
YJ_RE = re.compile(r'^\d{4}[0-9A-Z]{3}[A-Z][0-9A-Z]{4}$')  # 漢方・生薬などは英字を含む（例: 510000CA1020）
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
def is_medhot(rows):
    """医薬品コード登録システム（medhot.medd.jp）の全件ファイル：30列目が調剤包装単位コード(14桁)"""
    sample = [r for r in rows[:50] if len(r) >= 34]
    return len(sample) >= 3 and sum(1 for r in sample if re.fullmatch(r'\d{14}', r[29])) >= len(sample) * 0.6


def build(sources):
    items, key_index, hot_index, pending = [], {}, {}, []
    rows_read = 0

    def add_unique(lst, v):
        if v and v not in lst:
            lst.append(v)

    yj_index = {}

    def item_for(name, yj):
        k = (yj or 'N') + '|' + norm_key(name)
        idx = key_index.get(k)
        if idx is None and yj:
            idx = yj_index.get(yj)
        if idx is None:
            idx = len(items)
            items.append({'n': name, 'y': yj, 'g': [], 'h': []})
            key_index[k] = idx
        if yj and yj not in yj_index:
            yj_index[yj] = idx
        return idx

    for fname, data in sources:
        for tname, rows in tables_from_bytes(fname, data):
            if is_medhot(rows):
                n = 0
                for row in rows:
                    if len(row) < 34 or not re.fullmatch(r'\d{14}', row[29]):
                        continue
                    yj = row[8].upper() if YJ_RE.fullmatch(row[8].upper()) else ''
                    if not row[3]:
                        continue
                    it = items[item_for(row[3], yj)]
                    unit_code = row[29]
                    add_unique(it['g'], unit_code)
                    for v in (row[32], row[33]):
                        if re.fullmatch(r'\d{14}', v):
                            add_unique(it['g'], v)
                    label = row[30] or ' '.join(x for x in (row[22], row[23] + row[24]) if x)
                    if label:
                        it.setdefault('gp', {})[unit_code] = label[:40]
                    n += 1
                    rows_read += 1
                print(f'  read: {tname}（医薬品コード登録システム） {n:,}行')
                continue
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
                    idx = item_for(name, yj)
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


# ---------- 厚生労働省のページから最新ファイルを探す ----------
MHLW = 'https://www.mhlw.go.jp'
PRICE_PAGE = MHLW + '/topics/{y}/04/tp{y}0401-01.html'       # 薬価基準収載品目リスト（年度ごとのページ）
SUPPLY_PAGE = MHLW + '/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html'  # 医療用医薬品供給状況
JST = timezone(timedelta(hours=9))


def http_get(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (jishiyaku master builder)'})
    with urllib.request.urlopen(req, timeout=180) as r:
        return r.read()


def html_text(url):
    return decode(http_get(url))


def discover_price_list():
    """最新の「薬価基準収載品目リスト」ページを辿り、内用・注射・外用・歯科のExcelのURLを返す"""
    from urllib.parse import urljoin
    now = datetime.now(JST)
    page_url, html = None, None
    for y in (now.year, now.year - 1, now.year - 2):
        try:
            u = PRICE_PAGE.format(y=y)
            html, page_url = html_text(u), u
            break
        except Exception as e:
            print('  なし:', PRICE_PAGE.format(y=y), e)
    if not html:
        raise RuntimeError('薬価基準収載品目リストのページが見つかりません')
    # 「最新の情報はこちら」のように新しいページへのリンクがあれば辿る
    for _ in range(5):
        cur = re.search(r'tp(\d{8})-01\.html', page_url).group(1)
        newer = sorted({(m.group(2), urljoin(page_url, m.group(1))) for m in re.finditer(r'href="([^"]*/topics/\d{4}/\d{2}/tp(\d{8})-01\.html)"', html) if m.group(2) > cur})
        if not newer:
            break
        page_url = newer[-1][1]
        html = html_text(page_url)
    title = re.search(r'<title>(.*?)</title>', html, re.S)
    print('  薬価基準収載品目リスト:', page_url, title.group(1).strip() if title else '')
    best = {}
    for m in re.finditer(r'href="([^"]*tp(\d{8})-01_0([1-4])\.xlsx)"', html):
        kind, date = m.group(3), m.group(2)
        if kind not in best or date > best[kind][0]:
            best[kind] = (date, urljoin(page_url, m.group(1)))
    if not best:
        raise RuntimeError('Excelのリンクが見つかりません')
    return [best[k][1] for k in sorted(best)], (title.group(1).strip() if title else page_url)


MEDHOT_PAGE = 'https://medhot.medd.jp/view_download'


def discover_medhot():
    """医薬品コード登録システムの「調剤包装単位・販売包装単位・元梱包装単位コードファイル」(ZIP)のURL。
    ページ内のファイル名（A_YYYYMMDD_2）を探し、見つからなければ最近の月末日の名前を順に試す"""
    from urllib.parse import urljoin
    base = 'https://medhot.medd.jp/csv/'
    names = []
    try:
        html = html_text(MEDHOT_PAGE)
        names = sorted(set(re.findall(r'A_(\d{8})_2\.(?:zip|txt)', html)))
        if not names:
            print('  medhotページにファイル名がありません:', re.sub(r'\s+', ' ', html)[:300])
    except Exception as e:
        print('  medhotページを取得できません:', e)
    if names:
        url = base + f'A_{names[-1]}_2.zip'
        print('  医薬品コード登録システム:', url)
        return url
    # 推測：直近の月末日（と15日）を新しい順に試す
    today = datetime.now(JST).date()
    cands = []
    d = today
    for _ in range(4):
        first = d.replace(day=1)
        last_prev = first - timedelta(days=1)
        cands += [d.replace(day=15)] if d.day >= 15 else []
        cands.append(last_prev)
        d = last_prev
    for c in sorted(set(cands), reverse=True):
        url = base + f'A_{c:%Y%m%d}_2.zip'
        try:
            req = urllib.request.Request(url, method='HEAD', headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=60) as r:
                if r.status == 200:
                    print('  医薬品コード登録システム（推測）:', url)
                    return url
        except Exception:
            pass
    raise RuntimeError('GS1コードファイルが見つかりません')


def discover_supply():
    """「医療用医薬品供給状況」ページから最新のExcelのURLと「◯月◯日現在」の表記を返す"""
    from urllib.parse import urljoin
    html = html_text(SUPPLY_PAGE)
    links = []
    for m in re.finditer(r'<a[^>]+href="([^"]+\.xlsx)"[^>]*>(.*?)</a>', html, re.S):
        text = re.sub(r'<[^>]+>|\s', '', m.group(2))
        if '供給状況' in text or 'kyoukyu' in m.group(1):
            links.append((urljoin(SUPPLY_PAGE, m.group(1)), text))
    if not links:
        raise RuntimeError('供給状況のExcelのリンクが見つかりません')
    print('  医療用医薬品供給状況:', links[0][0], links[0][1])
    return links[0]


# ---------- 供給状況 ----------
SHIP = {1: '通常出荷', 2: '限定出荷（自社の事情）', 3: '限定出荷（他社品の影響）', 4: '限定出荷（その他）', 5: '供給停止'}


def ship_code(t):
    t = re.sub(r'\s', '', t)
    if not t:
        return 0
    if '供給停止' in t:
        return 5
    if '限定' in t:
        return 2 if '自社' in t else 3 if '他社' in t else 4
    if '通常' in t:
        return 1
    return 0


def qty_code(t):
    t = re.sub(r'\s', '', t)
    m = re.match(r'^[①-⑤]?([A-DＡ-Ｄ])(プラス|\+|＋)?', t)
    if not m:
        return ''
    return unicodedata.normalize('NFKC', m.group(1)) + ('+' if m.group(2) else '')


def to_date(v):
    """Excelの日付（シリアル値・日時文字列）を YYYY-MM-DD にする"""
    v = (v or '').strip()
    if re.fullmatch(r'\d{5}(\.\d+)?', v):
        return (datetime(1899, 12, 30) + timedelta(days=int(float(v)))).strftime('%Y-%m-%d')
    m = re.match(r'(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})', v)
    return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}' if m else v[:10]


SKIPPED = []


def parse_supply(tables):
    out = {}
    for tname, rows in tables:
        hdr = None
        for r, row in enumerate(rows[:30]):
            h = [re.sub(r'\s', '', c) for c in row]
            if any('YJ' in c for c in h) and any('出荷' in c for c in h):
                hdr = r
                break
        if hdr is None:
            continue
        H = [re.sub(r'\s', '', c) for c in rows[hdr]]

        def col(pred):
            return next((i for i, c in enumerate(H) if pred(c)), -1)
        yj = col(lambda c: 'YJ' in c)
        ship = col(lambda c: '出荷対応' in c and not re.search(r'日|理由|見込', c))
        qty = col(lambda c: '出荷量' in c and not re.search(r'日|理由|見込|改善', c))
        upd = col(lambda c: '更新' in c and '日' in c)
        reason = col(lambda c: '理由' in c)
        expect = col(lambda c: '見込' in c)
        print(f'  read: {tname}  yj={yj} 出荷対応={ship} 出荷量={qty} 更新日={upd} 理由={reason} 見込み={expect}')
        for row in rows[hdr + 1:]:
            g = lambda i: row[i] if 0 <= i < len(row) else ''
            code = g(yj).upper()
            if not YJ_RE.fullmatch(code):
                if code and len(SKIPPED) < 8:
                    SKIPPED.append(code)
                continue
            s, q = ship_code(g(ship)), qty_code(g(qty))
            if not s and not q:
                if len(SKIPPED) < 8:
                    SKIPPED.append(f'{code}:{g(ship)[:20]}/{g(qty)[:10]}')
                continue
            rec = [s, q, to_date(g(upd))]
            r_, e_ = g(reason), g(expect)
            if s and s != 1 and (r_ or e_):
                rec += [r_[:80], e_[:40]]
            out[code] = rec
    return out


# ---------- 採用薬リスト（棚表） ----------
FORM_SRC = ROOT / 'formulary-src'
FIXES = ROOT / 'formulary-fixes.csv'
FOLD = str.maketrans({'パ': 'バ', 'ピ': 'ビ', 'プ': 'ブ', 'ペ': 'ベ', 'ポ': 'ボ', '一': 'ー', '－': 'ー', '‐': 'ー', '-': 'ー',
                      '口': 'ロ', '力': 'カ', '工': 'エ', '夕': 'タ', '卜': 'ト'})
NUM_RE = re.compile(r'\d+(?:\.\d+)?')
STRENGTH_COL = re.compile(r'^\d+(?:\.\d+)?\s*(mg|g|%|μg|mL|ml|単位|mEq|IU)')


def nkey(s):
    s = unicodedata.normalize('NFKC', s or '')
    s = ''.join(chr(ord(c) + 0x60) if 'ぁ' <= c <= 'ゖ' else c for c in s)
    s = re.sub(r'【般】|\[般\]|\(般\)|（般）', '', s)
    s = re.sub(r'[\s・「」『』"\'“”()（）]', '', s).translate(FOLD)
    s = s.replace('ナトリウム', 'NA').replace('カリウム', 'K').replace('カルシウム', 'CA').replace('マグネシウム', 'MG')
    return s.upper()


def formulary_rows(tables):
    """棚表から「薬品名（＋規格の列）」と、あればコードを取り出す"""
    out = []
    for tname, rows in tables:
        sample = [r for r in rows[:300] if any(r)]
        width = max((len(r) for r in sample), default=0)
        if not width:
            continue

        def ratio(i, test):
            vals = [r[i] for r in sample if i < len(r) and r[i]]
            return sum(1 for v in vals if test(v)) / max(1, len(sample))
        name_col = max(range(width), key=lambda i: ratio(i, lambda v: bool(re.search(r'[ァ-ヴｦ-ﾟ]{2,}', v))))
        if ratio(name_col, lambda v: bool(re.search(r'[ァ-ヴｦ-ﾟ]{2,}', v))) < 0.2:
            continue
        spec_cols = [i for i in range(width) if i != name_col and ratio(i, lambda v: bool(STRENGTH_COL.match(v))) > 0.2]
        code_col = next((i for i in range(width) if ratio(i, lambda v: bool(YJ_RE.fullmatch(v.upper()))) > 0.2), -1)
        print(f'  棚表: {tname}  名前={name_col} 規格={spec_cols} コード={code_col}')
        for r in rows:
            g = lambda i: r[i] if 0 <= i < len(r) else ''
            name = g(name_col)
            if not name or not re.search(r'[ァ-ヴｦ-ﾟA-Za-z一-龥]{2,}', name) or re.search(r'薬品名|品名|医薬品名', name):
                continue
            spec = ' '.join(g(i) for i in spec_cols if g(i) and g(i) not in name)
            out.append({'r': (name + (' ' + spec if spec else '')).strip(), 'code': g(code_col).upper() if code_col >= 0 else ''})
    return out


def build_formulary(items):
    from rapidfuzz import fuzz, process
    tables = []
    for p in sorted(FORM_SRC.glob('*')):
        if p.is_file() and p.suffix.lower() in ('.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.tsv'):
            tables += list(tables_from_bytes(p.name, p.read_bytes()))
    if not tables:
        return None
    rows = formulary_rows(tables)
    fixes = {}
    if FIXES.exists():
        import csv
        for r in csv.reader(io.StringIO(decode(FIXES.read_bytes()))):
            if len(r) >= 2 and r[0].strip() and not r[0].startswith('#') and r[0].strip() != '棚表の表記':
                fixes[nkey(r[0])] = r[1].strip().upper()
    keys = [nkey(it['n']) for it in items]
    by_y = {}
    for it in items:
        if it.get('y'):
            by_y.setdefault(it['y'], it)
    out, seen = [], set()
    for row in rows:
        raw = row['r']
        k = nkey(raw)
        if not k or k in seen:
            continue
        seen.add(k)
        rec = {'r': raw}
        fix = fixes.get(k)
        if fix == '除外':
            continue
        code = fix or (row['code'] if YJ_RE.fullmatch(row['code']) else '')
        if code:
            it = by_y.get(code)
            rec.update(y=code, n=it['n'] if it else raw, s=100, st='fix' if fix else 'ok')
            out.append(rec)
            continue
        nums = NUM_RE.findall(k)
        cands = []
        for name_key, score, idx in process.extract(k, keys, scorer=fuzz.WRatio, limit=25):
            cn = NUM_RE.findall(name_key)
            if nums and any(n not in cn for n in nums):
                score -= 25                         # 規格の数字が違うものは大きく下げる
            if not nums and cn:
                score -= 5
            cands.append((score, idx))
        cands.sort(key=lambda x: -x[0])
        uniq, used = [], set()
        for score, idx in cands:
            it = items[idx]
            if it['n'] in used:
                continue
            used.add(it['n'])
            uniq.append((round(score), it))
            if len(uniq) >= 4:
                break
        if not uniq:
            rec.update(st='none')
            out.append(rec)
            continue
        top_s, top = uniq[0]
        second = uniq[1] if len(uniq) > 1 else None
        same9 = second and top.get('y') and second[1].get('y') and top['y'][:9] == second[1]['y'][:9]
        if top_s >= 88 and (not second or top_s - second[0] >= 4):
            st = 'ok'
        elif top_s >= 85 and same9:
            st = 'amb'                              # 銘柄（メーカー）だけ曖昧。成分・剤形・規格は確定
        elif top_s >= 60:
            st = 'check'
        else:
            st = 'none'
        rec.update(y=top.get('y', ''), n=top['n'], s=top_s, st=st)
        if st != 'ok':
            rec['c'] = [[it['n'], it.get('y', ''), sc] for sc, it in uniq]
        out.append(rec)
    counts = {k: sum(1 for r in out if r['st'] == k) for k in ('ok', 'fix', 'amb', 'check', 'none')}
    ver = hashlib.sha1(json.dumps(out, ensure_ascii=False).encode()).hexdigest()[:12]
    meta = {'version': ver, 'date': datetime.now(JST).strftime('%Y/%m/%d'), 'count': len(out), 'counts': counts,
            'files': [p.name for p in sorted(FORM_SRC.glob('*')) if p.is_file() and not p.name.startswith('.') and p.name != 'README.md']}
    write_json(ROOT / 'formulary.json', {'meta': meta, 'items': out})
    print(f'棚表 {len(out):,}品目 確定{counts["ok"] + counts["fix"]:,} 銘柄あいまい{counts["amb"]:,} 要確認{counts["check"]:,} 該当なし{counts["none"]:,}')
    return meta


def write_json(path, obj):
    path.write_text(json.dumps(obj, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')


def main():
    args = [u for a in sys.argv[1:] for u in a.split()]
    auto = '--auto' in args
    urls = [a for a in args if a.startswith('http')]
    meta_path = ROOT / 'master-meta.json'
    meta = json.loads(meta_path.read_text(encoding='utf-8')) if meta_path.exists() else {}
    jst_today = datetime.now(JST).strftime('%Y/%m/%d')
    failed = []

    # 1) 薬品マスター：master-src/ のファイル ＋ 指定URL ＋（--auto なら）厚労省の最新リスト
    sources = []
    for p in sorted(SRC.glob('*')):
        if p.is_file() and not p.name.startswith('.') and p.suffix.lower() in ('.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.tsv', '.zip'):
            sources.append((p.name, p.read_bytes()))
    price_title = ''
    if auto:
        try:
            price_urls, price_title = discover_price_list()
            urls = price_urls + urls
        except Exception as e:
            failed.append(f'薬価基準収載品目リスト: {e}')
        try:
            urls.append(discover_medhot())
        except Exception as e:
            failed.append(f'GS1コード（医薬品コード登録システム）: {e}')
    for url in urls:
        print('download:', url)
        try:
            sources.append((url.rsplit('/', 1)[-1].split('?')[0] or 'download', http_get(url)))
        except Exception as e:
            failed.append(f'{url}: {e}')
    if sources:
        items, rows_read, linked, unlinked = build(sources)
        if items:
            items.sort(key=lambda it: (it.get('y') or 'ZZZ', it['n']))
            version = hashlib.sha1(json.dumps(items, ensure_ascii=False).encode()).hexdigest()[:12]
            if meta.get('version') != version:
                meta.update(version=version, date=jst_today)
            meta.update(count=len(items), gs1=sum(len(it.get('g', [])) for it in items),
                        files=[n for n, _ in sources], source=price_title)
            write_json(ROOT / 'master.json', {'meta': {k: meta[k] for k in ('version', 'date', 'count')}, 'items': items})
            print(f'行 {rows_read:,} → 品目 {len(items):,}（GS1 {meta["gs1"]:,}件, 対応表 {linked:,}行を結合, 未結合 {unlinked:,}行） version={version}')
        else:
            failed.append('薬品を1件も読み取れませんでした')

    # 2) 供給状況（--auto のとき）
    if auto:
        try:
            url, label = discover_supply()
            data = http_get(url)
            supply = parse_supply(list(tables_from_bytes(url.rsplit('/', 1)[-1], data)))
            if not supply:
                raise RuntimeError('供給状況の表を読み取れませんでした（列名が変わった可能性があります）')
            asof = re.search(r'[（(](.*?現在)[）)]', label)
            sv = hashlib.sha1(json.dumps(supply, ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:12]
            counts = {SHIP[k]: sum(1 for v in supply.values() if v[0] == k) for k in SHIP}
            meta['supply'] = {'version': sv, 'asof': asof.group(1) if asof else label, 'count': len(supply), 'counts': counts}
            if SKIPPED:
                meta['supply']['skipped_samples'] = SKIPPED
            write_json(ROOT / 'supply.json', {'meta': meta['supply'], 'codes': SHIP, 'items': supply})
            print(f'供給状況 {len(supply):,}品目 ' + ' / '.join(f'{k}{v:,}' for k, v in counts.items()))
        except Exception as e:
            failed.append(f'供給状況: {e}')

    # 3) 採用薬リスト（棚表）
    try:
        mfile = ROOT / 'master.json'
        if mfile.exists() and FORM_SRC.exists():
            fm = build_formulary(json.loads(mfile.read_text(encoding='utf-8'))['items'])
            if fm:
                meta['formulary'] = fm
    except Exception as e:
        failed.append(f'採用薬リスト: {e}')

    if meta:
        if failed:
            meta['errors'] = failed
        else:
            meta.pop('errors', None)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=1), encoding='utf-8')
    for f in failed:
        print('::warning::' + f)
    if not meta.get('version') and not meta.get('supply'):
        print('作成できるデータがありませんでした')
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
