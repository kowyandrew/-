#!/usr/bin/env python3
"""契約図PDF → 案件属性JSON + サムネイル

使い方:
  python3 extract.py --cases <cases/*.json ...> --pdfdir <pdf dir> --out plans.json --thumbs thumbs/

cases JSON は Drive から集めた一覧（customer, fileId, title, modifiedTime, viewUrl, pdfPath, selected ...）。
PDF はベクター（CADから出力、文字が取れる）とスキャン（文字が取れない）が混在する。
ベクターは本文から属性を正規表現で抜く。スキャンは ocrText（Drive のOCR文字列）が与えられていればそれを使う。
"""
import argparse, json, os, re, sys, glob
import pymupdf

ROOM_WORDS = ('洋室', '和室', '寝室', '主寝室', '子供室', '子供部屋', '書斎', '納戸', 'ＬＤＫ', 'LDK', 'ＤＫ', 'DK',
              'リビング', 'ダイニング', 'キッチン', '玄関', 'ホール', '廊下', 'トイレ', '洗面', '脱衣', 'ＵＢ', 'UB',
              '浴室', 'ＷＩＣ', 'WIC', 'ウォークイン', 'クローゼット', 'ＣＬ', 'CL', 'パントリー', 'シューズ', 'ＳＣ', 'SIC',
              '物入', '収納', 'ファミリー', 'フリー', 'ロフト', '小屋裏', '土間', 'バルコニー', 'テラス', 'ポーチ', '階段',
              'ランドリー', '事務', '店舗', '倉庫', 'ガレージ', 'カーポート', '多目的', '趣味')
BEDROOM_RE = re.compile(r'^(洋室|和室|寝室|主寝室|子供室|子供部屋|居室)')

FEATURES = [
    ('平屋', r'平屋'), ('2階建', r'2階建|２階建|2 階 平面|２階平面|2階平面|床 面 積 表<2階>|2階 床面積'),
    ('太陽光', r'太陽光発電|太陽光パネル|太陽光ﾊﾟﾈﾙ|太陽光モジュール|ｿｰﾗｰﾊﾟﾈﾙ|ソーラーパネル'), ('吹抜', r'吹抜|吹き抜'), ('ロフト', r'ロフト'), ('小屋裏', r'小屋裏'),
    ('WIC', r'ＷＩＣ|WIC|ウォークイン|ｳｫｰｸｲﾝ'), ('パントリー', r'パントリー|ﾊﾟﾝﾄﾘｰ'), ('書斎', r'書斎'),
    ('シューズクローク', r'シューズ|ｼｭｰｽﾞ|ＳＩＣ|SIC|ＳＣ(?![A-Za-z])'), ('土間収納', r'土間収納|土間ｼｭｰｽﾞ|土間(?!ｺﾝ|コン|打)'),
    ('バルコニー', r'バルコニー|ﾊﾞﾙｺﾆｰ|ベランダ'), ('カーポート', r'カーポート|ｶｰﾎﾟｰﾄ'), ('勝手口', r'勝手口'),
    ('ランドリー', r'ランドリー|ﾗﾝﾄﾞﾘｰ|洗濯室'), ('和室', r'和室'), ('ファミリークローク', r'ファミリークロー|ﾌｧﾐﾘｰｸﾛｰ|ＦＣＬ|FCL'),
    ('シャッター', r'ｼｬｯﾀｰあり|シャッターあり|電動ｼｬｯﾀｰ'), ('ホスクリーン', r'ﾎｽｸﾘｰﾝ|ホスクリーン'), ('幹太くん', r'幹太くん'),
    ('宅配ボックス', r'宅配'), ('都市ガス', r'都市ガス'), ('プロパン', r'プロパン|ＬＰ|LPガス'),
    ('浄化槽', r'浄化槽'), ('下水道', r'下水道'), ('増築', r'増築'), ('店舗・事務所', r'店舗|事務所|事務室'), ('賃貸・共同住宅', r'共同住宅|アパート|ＡＰ|長屋'),
]

SHEET_WORDS = ['案内図', '配置図', '求積図', '平面詳細図', '平面図', '立面図', '断面図', '基礎伏図', '電気配線図', '換気計算', 'シックハウス',
               '矩計図', '屋根伏図', '床伏図', '小屋伏図', '展開図', '建具表', '仕上表', '仕様書', '給排水', '外構']

WAREKI = {'令和': 2018, '平成': 1988}


def z2h(s: str) -> str:
    """全角英数・記号を半角に。半角カナはそのまま。"""
    return s.translate(str.maketrans('０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ（）：．，－',
                                   '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz():.,-'))


def num(s):
    try:
        return float(s.replace(',', ''))
    except Exception:
        return None


def parse_wareki(text):
    m = re.search(r'(令和|平成)\s*(\d{1,2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', text)
    if m:
        y = WAREKI[m.group(1)] + int(m.group(2))
        return f'{y:04d}-{int(m.group(3)):02d}-{int(m.group(4)):02d}'
    m = re.search(r'(20\d\d)[/.年](\d{1,2})[/.月](\d{1,2})', text)
    if m:
        return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    return None


def extract_rooms(lines):
    """『洋室(2)』の次行に『（9.11㎡ 5.5帖）』が来る形式を拾う。"""
    rooms = []
    seen = set()
    for i, ln in enumerate(lines):
        m = re.search(r'\(?\s*([\d.]+)\s*㎡\s*([\d.]+)\s*[帖畳]\s*\)?', ln)
        if not m:
            continue
        name = None
        for j in range(i - 1, max(i - 4, -1), -1):
            cand = lines[j].strip()
            if not cand or re.match(r'^[\d.,()（）\s㎡帖畳CHFL±+\-]+$', cand):
                continue
            if any(w in cand for w in ROOM_WORDS) or re.match(r'^[^\d]{1,12}$', cand):
                name = cand
            break
        if not name:
            continue
        key = (name, m.group(1))
        if key in seen:
            continue
        seen.add(key)
        rooms.append({'name': name, 'm2': num(m.group(1)), 'jo': num(m.group(2))})
    return rooms


def extract_areas(text):
    """トータル面積表から敷地・建築・床・延床・建蔽率・容積率を取る。"""
    out = {}
    t = z2h(text)
    # 「トータル面積表」ブロック: 見出し群のあとに数値が並ぶ
    m = re.search(r'ト\s*ー\s*タ\s*ル\s*面\s*積\s*表(.{0,600}?)建築可能容積率(.{0,200})', t, re.S)
    if m:
        head, body = m.group(1), m.group(2)
        nums = re.findall(r'(?<=\n)\s*([\d,]+\.\d+)\s*(?=\n)', '\n' + body + '\n')
        floors = re.findall(r'(?<=\n)\s*(\d)階\s*(?=\n)', '\n' + head + '\n')
        # 並び: 敷地, 建築, 1階, (2階,) 延床
        if len(nums) >= 3:
            out['siteArea'] = num(nums[0])
            out['buildingArea'] = num(nums[1])
            nf = max(1, len(floors))
            fl = nums[2:2 + nf]
            if len(fl) >= 1:
                out['floorArea1'] = num(fl[0])
            if len(fl) >= 2:
                out['floorArea2'] = num(fl[1])
            if len(nums) >= 2 + nf + 1:
                out['totalFloorArea'] = num(nums[2 + nf])
        pc = re.findall(r'([\d.]+)%', body)
        if len(pc) >= 2:
            out['bcr'] = num(pc[0])
            out['far'] = num(pc[1])
        if len(pc) >= 4:
            out['bcrLimit'] = num(pc[2])
            out['farLimit'] = num(pc[3])
    # 個別表記のフォールバック
    for key, pat in [('siteArea', r'敷\s*地\s*面\s*積\s*[:：]?\s*([\d,]+\.\d+)'),
                     ('buildingArea', r'建\s*築\s*面\s*積\s*[:：]?\s*([\d,]+\.\d+)'),
                     ('totalFloorArea', r'延\s*べ?\s*床\s*面\s*積\s*[:：]?\s*([\d,]+\.\d+)')]:
        if key not in out:
            m = re.search(pat, t)
            if m:
                out[key] = num(m.group(1))
    # 床面積表<1階> 合計（坪, ㎡）
    for fl in ('1', '2'):
        m = re.search(r'床\s*面\s*積\s*表\s*<' + fl + r'階>(.{0,1500}?)合計', t, re.S)
        if m and f'floorArea{fl}' not in out:
            pass
    m = re.search(r'計\(坪\)', t)
    tsubo = re.findall(r'\n\s*(\d{1,3}\.\d{2})\s{2,}\n', t)
    return out


def extract_from_text(text):
    lines = [l.strip() for l in text.split('\n')]
    t = z2h(text)
    d = {}
    m = re.search(r'([^\n]{1,40}?(?:様邸|様|殿|御中)?[^\n]{0,20}?新築工事|[^\n]{1,40}?増築工事|[^\n]{1,40}?改修工事)', text)
    if m:
        d['workName'] = m.group(1).strip()
    m = re.search(r'建築地\s*[:：]?\s*([^\n]+)', text)
    if m:
        d['site'] = m.group(1).strip()
    else:
        m = re.search(r'(千葉県[^\n]{3,40})', text)
        if m:
            d['site'] = m.group(1).strip()
    if d.get('site'):
        m = re.search(r'(千葉県|東京都|埼玉県|茨城県|神奈川県)?\s*([^\s\d]+?[市郡])\s*([^\s\d]+?[区町村])?', d['site'])
        if m:
            d['city'] = (m.group(2) or '') + (m.group(3) or '')
    d['contractDate'] = parse_wareki(text)
    d.update(extract_areas(text))
    rooms = extract_rooms(lines)
    d['rooms'] = rooms
    beds = sorted({r['name'] for r in rooms if BEDROOM_RE.match(z2h(r['name']))})
    ldk = [r for r in rooms if re.search(r'LDK|ＬＤＫ|DK|ＤＫ', r['name'])]
    d['bedrooms'] = len(beds)
    if ldk:
        d['ldkJo'] = max(r['jo'] for r in ldk if r['jo']) if any(r['jo'] for r in ldk) else None
        d['layout'] = f"{len(beds)}{'LDK' if re.search('LDK|ＬＤＫ', ldk[0]['name']) else 'DK'}" if beds else ('LDK' if re.search('LDK|ＬＤＫ', ldk[0]['name']) else 'DK')
    elif beds:
        d['layout'] = f'{len(beds)}R'
    m = re.search(r'最高高さ\s*[:：]?\s*([\d.]+)\s*m', t)
    if m:
        d['maxHeight'] = num(m.group(1))
    two = re.search(r'2\s*階\s*平面|2階\s*床面積|床\s*面\s*積\s*表\s*<2階>|2階平面', t)
    d['floors'] = 2 if two else 1
    if re.search(r'3\s*階\s*平面|床\s*面\s*積\s*表\s*<3階>', t):
        d['floors'] = 3
    feats = []
    for name, pat in FEATURES:
        if name in ('平屋', '2階建'):
            continue
        if re.search(pat, text) or re.search(pat, t):
            feats.append(name)
    feats.insert(0, '平屋' if d['floors'] == 1 else f"{d['floors']}階建")
    d['features'] = feats
    ops = sorted({z2h(x.strip()) for x in re.findall(r'OP\s*[:：]\s*([^\n]{2,40})', text)})
    d['options'] = ops[:40]
    sheets = [s for s in SHEET_WORDS if s in text]
    d['sheets'] = sheets
    m = re.search(r'(法第?4?2条[^\n]{0,20}道路)', text)
    if m:
        d['road'] = z2h(m.group(1))
    m = re.search(r'道路幅員[^\n]*?\n?\s*([\d,]{3,6})', t)
    if m:
        d['roadWidthMm'] = num(m.group(1))
    return d


def scanned(doc):
    chars = sum(len(p.get_text()) for p in doc)
    return chars < 200


def thumbnail(doc, out_path, prefer=('平面詳細図', '平面図'), width=900):
    """平面図ページを優先してJPEGサムネイルを作る。"""
    idx = 0
    best = None
    for i, p in enumerate(doc):
        t = p.get_text()
        for k, w in enumerate(prefer):
            if w in t and ('1 階' in t or '１階' in t or '1階' in t or k == 0):
                if best is None or k < best[0]:
                    best = (k, i)
    if best:
        idx = best[1]
    elif len(doc) > 2:
        idx = 2
    page = doc[idx]
    zoom = width / page.rect.width
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    pix.save(out_path, jpg_quality=70)
    return idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cases', nargs='+', required=True)
    ap.add_argument('--pdfdir', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--thumbs', required=True)
    ap.add_argument('--ocr', help='JSON {fileId: ocrText} for scanned PDFs')
    ap.add_argument('--manual', help='JSON {fileId: {field: value}} 手動補正')
    a = ap.parse_args()
    os.makedirs(a.thumbs, exist_ok=True)
    ocr = json.load(open(a.ocr)) if a.ocr and os.path.exists(a.ocr) else {}
    manual = json.load(open(a.manual)) if a.manual and os.path.exists(a.manual) else {}
    plans = []
    seen = set()
    for cf in a.cases:
        for c in json.load(open(cf)):
            if not c.get('selected'):
                continue
            fid = c['fileId']
            if fid in seen:
                continue
            seen.add(fid)
            pdf = os.path.join(a.pdfdir, f'{fid}.pdf')
            rec = {'id': fid, 'customer': c.get('customer'), 'title': c.get('title'), 'folder': c.get('folder', ''),
                   'docType': c.get('docType', ''), 'modifiedTime': c.get('modifiedTime'), 'viewUrl': c.get('viewUrl'),
                   'owner': c.get('owner'), 'fileSize': c.get('fileSize'), 'scanned': None, 'pages': None}
            if not os.path.exists(pdf):
                rec['error'] = 'pdf missing'
                plans.append(rec)
                continue
            try:
                doc = pymupdf.open(pdf)
            except Exception as e:
                rec['error'] = f'open failed: {e}'
                plans.append(rec)
                continue
            rec['pages'] = len(doc)
            rec['scanned'] = scanned(doc)
            text = ocr.get(fid, '') if rec['scanned'] else '\n'.join(p.get_text() for p in doc)
            rec['textSource'] = 'ocr' if rec['scanned'] else 'pdf'
            if text:
                rec.update(extract_from_text(text))
            try:
                rec['thumbPage'] = thumbnail(doc, os.path.join(a.thumbs, f'{fid}.jpg'))
                rec['thumb'] = f'thumbs/{fid}.jpg'
            except Exception as e:
                rec['thumbError'] = str(e)
            rec.update(manual.get(fid, {}))
            if not rec.get('contractDate') and rec.get('modifiedTime'):
                rec['contractDate'] = rec['modifiedTime'][:10]
                rec['contractDateSource'] = 'drive'
            plans.append(rec)
    plans.sort(key=lambda r: r.get('contractDate') or '', reverse=True)
    json.dump(plans, open(a.out, 'w'), ensure_ascii=False, indent=1)
    ok = sum(1 for p in plans if not p.get('error'))
    print(f'{len(plans)} plans, {ok} parsed, {sum(1 for p in plans if p.get("scanned"))} scanned')


if __name__ == '__main__':
    main()
