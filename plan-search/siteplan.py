#!/usr/bin/env python3
"""測量図DXF → 配置図のたたき台（DXF）

敷地の測量図から敷地形状を読み取り、建物と駐車場を法規・社内ルールの範囲に置いて、
Jw_cad で開ける配置図DXFを書き出す。設計担当が手直しする前の「たたき台」を作るのが目的。

  # 1) 敷地候補を見る
  python3 siteplan.py --dxf 測量図.dxf --list

  # 2) 配置図を作る
  python3 siteplan.py --dxf 測量図.dxf --parcel 1173-2 --road-edge 3 \\
      --road-width 4.0 --building 5.46x9.555 --parking 2 --out 配置図.dxf

建物は --building W×D（m）で直接指定するか、--from-plan 施主名 で過去の契約図から持ってくる。

出力はあくまで下書き。最終的な図面は建築士が確認・修正すること。
"""
import argparse, json, math, os, re, sys
from collections import defaultdict

import ezdxf
from shapely.geometry import Polygon, Point, box
from shapely.affinity import rotate, translate

TSUBO = 3.305785
# 敷地境界線が入っている可能性のあるレイヤ名（測量事務所によって違う）
# 測量事務所によってレイヤ名が日本語だったりローマ字だったりする
BOUNDARY_LAYERS = ('求積線分', '境界線', '敷地', '画地', '求積', '外郭線',
                   'KYUSEKI', 'GAIKUSEN', 'KYOUKAI', 'KYOKAI', 'SIKICHI', 'SHIKICHI', 'PLOT')
# 求積「表」の罫線など、図形でないレイヤは除く
EXCLUDE_LAYERS = ('表', '罫線', '凡例', '文字', '寸法', 'Defpoints',
                  'KYORI', 'MIDASHI', 'TTL', 'MARK', 'ME')
POINT_NAME_LAYERS = ('求積地点名', '点名', 'プロット点名', 'SOKUTENME', 'TENME')
LOT_NAME_LAYERS = ('地番名', '地番', 'CHIBANME', 'CHIBAN')


# ---------------------------------------------------------------- 敷地の読み取り

def _k(x, y, q=3):
    return (round(x, q), round(y, q))


def collect_segments(msp, layer_hint=None):
    """境界線らしいレイヤから線分を集める。見つからなければ全LINEを使う。"""
    def want(name):
        if any(x in name for x in EXCLUDE_LAYERS):
            return False
        if layer_hint:
            return layer_hint in name
        return any(h in name for h in BOUNDARY_LAYERS)

    segs, used = set(), set()
    for e in msp:
        if e.dxftype() not in ('LINE', 'LWPOLYLINE', 'POLYLINE'):
            continue
        if not want(e.dxf.layer):
            continue
        used.add(e.dxf.layer)
        if e.dxftype() == 'LINE':
            a = _k(e.dxf.start.x, e.dxf.start.y)
            b = _k(e.dxf.end.x, e.dxf.end.y)
            if a != b:
                segs.add((a, b))
        else:
            pts = [_k(p[0], p[1]) for p in e.get_points('xy')]
            if getattr(e, 'closed', False) or e.dxf.get('flags', 0) & 1:
                pts.append(pts[0])
            for a, b in zip(pts, pts[1:]):
                if a != b:
                    segs.add((a, b))
    return segs, sorted(used)


def planar_faces(segs):
    """線分の集合を平面グラフとみなし、囲まれた面をすべて取り出す。

    区画が辺を共有していると単純な追跡では拾えないので、各頂点で角度順に
    並べた半辺をたどる（次の半辺 = 入ってきた辺の時計回り隣）。
    """
    adj = defaultdict(set)
    for a, b in segs:
        adj[a].add(b)
        adj[b].add(a)
    order = {v: sorted(ns, key=lambda w: math.atan2(w[1] - v[1], w[0] - v[0]))
             for v, ns in adj.items()}

    faces, used = [], set()
    for u in adj:
        for v in adj[u]:
            if (u, v) in used:
                continue
            face, cur = [], (u, v)
            while cur not in used:
                used.add(cur)
                face.append(cur[0])
                a, b = cur
                ns = order[b]
                cur = (b, ns[(ns.index(a) - 1) % len(ns)])
            if len(face) >= 3:
                faces.append(face)
    return faces


def signed_area(ring):
    s = 0.0
    for i in range(len(ring)):
        x1, y1 = ring[i]
        x2, y2 = ring[(i + 1) % len(ring)]
        s += x1 * y2 - x2 * y1
    return s / 2


def read_parcels(path, layer_hint=None, min_area=20.0):
    """DXFから敷地候補のポリゴンを取り出す。面積の大きい順。"""
    doc = ezdxf.readfile(path, errors='ignore')
    msp = doc.modelspace()
    segs, layers = collect_segments(msp, layer_hint)
    if not segs:
        segs, layers = collect_segments(msp, layer_hint='')  # 最後の手段: 全レイヤ

    # 地番名のラベル位置（どのポリゴンがどの地番かの判定に使う）
    labels = []
    for e in msp:
        if e.dxftype() not in ('TEXT', 'MTEXT'):
            continue
        if not any(h in e.dxf.layer for h in LOT_NAME_LAYERS):
            continue
        t = (e.dxf.text if e.dxftype() == 'TEXT' else e.text).strip()
        labels.append((t, Point(e.dxf.insert.x, e.dxf.insert.y)))

    seen, parcels = set(), []
    for face in planar_faces(segs):
        a = signed_area(face)
        ring = face if a > 0 else face[::-1]      # 反時計回りに揃える
        if abs(a) < min_area:
            continue
        key = tuple(sorted(ring))
        if key in seen:
            continue
        seen.add(key)
        poly = Polygon(ring)
        if not poly.is_valid:
            poly = poly.buffer(0)
        if poly.is_empty or poly.area < min_area:
            continue
        names = sorted({t for t, p in labels if poly.contains(p)})
        parcels.append({'ring': ring, 'poly': poly, 'area': poly.area, 'lots': names})

    # 地番のついたポリゴンを優先（図郭＝図面の枠を敷地と間違えないため）。次に面積順
    parcels.sort(key=lambda p: (0 if p['lots'] else 1, -p['area']))
    uniq, shapes = [], set()
    for p in parcels:
        sig = (round(p['area'], 1),
               tuple(sorted(round(edge_len(e), 2) for e in edges_of(p['poly']))))
        if sig in shapes:
            continue
        shapes.add(sig)
        uniq.append(p)
    return uniq, layers


# ---------------------------------------------------------------- 配置

def edges_of(poly):
    c = list(poly.exterior.coords)[:-1]
    return [(c[i], c[(i + 1) % len(c)]) for i in range(len(c))]


def edge_len(e):
    return math.dist(e[0], e[1])


def inward_normal(poly, e):
    """辺 e の敷地内側を向く単位法線。"""
    (x1, y1), (x2, y2) = e
    n = (-(y2 - y1), x2 - x1)
    L = math.hypot(*n) or 1.0
    n = (n[0] / L, n[1] / L)
    mid = ((x1 + x2) / 2, (y1 + y2) / 2)
    probe = Point(mid[0] + n[0] * 1e-4, mid[1] + n[1] * 1e-4)
    return n if poly.contains(probe) else (-n[0], -n[1])


def half_plane(poly, e, dist):
    """辺 e から内側に dist だけ離れた側の領域（敷地の外接矩形で切る）。"""
    if dist <= 0:
        return poly
    n = inward_normal(poly, e)
    (x1, y1), (x2, y2) = e
    big = max(poly.bounds[2] - poly.bounds[0], poly.bounds[3] - poly.bounds[1]) * 3 + 50
    d = (x2 - x1, y2 - y1)
    L = math.hypot(*d) or 1.0
    d = (d[0] / L, d[1] / L)
    ox, oy = x1 + n[0] * dist, y1 + n[1] * dist
    pts = [(ox - d[0] * big, oy - d[1] * big),
           (ox + d[0] * big, oy + d[1] * big),
           (ox + d[0] * big + n[0] * big, oy + d[1] * big + n[1] * big),
           (ox - d[0] * big + n[0] * big, oy - d[1] * big + n[1] * big)]
    return Polygon(pts)


def buildable_area(poly, road_idx, setback_road, setback_side):
    """建物を置ける範囲。隣地側は一律、道路側だけ別の後退距離。"""
    region = poly.buffer(-setback_side, join_style=2)
    if region.is_empty:
        return region
    road_e = edges_of(poly)[road_idx]
    return region.intersection(half_plane(poly, road_e, setback_road))


def building_candidates(poly, region, road_idx, w, d, step=0.25, limit=60):
    """建物 w×d を region 内に置ける候補を、道路に近い順に返す。

    道路と平行／直交の2通りを試す。返り値は (建物ポリゴン, 向き, 道路からの距離)。
    """
    if region.is_empty:
        return []
    road_e = edges_of(poly)[road_idx]
    (x1, y1), (x2, y2) = road_e
    theta = math.degrees(math.atan2(y2 - y1, x2 - x1))
    n = inward_normal(poly, road_e)
    pivot = (x1, y1)
    rot_ny = n[0] * math.sin(math.radians(-theta)) + n[1] * math.cos(math.radians(-theta))
    sign = 1.0 if rot_ny > 0 else -1.0

    cands = []
    for (bw, bd), tag in (((w, d), '道路と平行'), ((d, w), '道路と直交')):
        reg = rotate(region, -theta, origin=pivot, use_radians=False)
        minx, miny, maxx, maxy = reg.bounds
        if maxx - minx < bw - 1e-9 or maxy - miny < bd - 1e-9:
            continue
        nx = max(1, int((maxx - minx - bw) / step) + 1)
        ny = max(1, int((maxy - miny - bd) / step) + 1)
        cx = (minx + maxx) / 2
        for j in range(ny):
            yy = miny + j * step
            row = []
            for i in range(nx):
                xx = minx + i * step
                r = box(xx, yy, xx + bw, yy + bd)
                if reg.contains(r):
                    row.append(r)
            if not row:
                continue
            r = min(row, key=lambda t: abs(t.centroid.x - cx))   # 左右は中央寄せ
            cand = rotate(r, theta, origin=pivot, use_radians=False)
            # 道路辺からの最短距離（奥に置くほど大きい）
            gap = min(math.dist(p, _closest_on_edge(road_e, p))
                      for p in list(cand.exterior.coords))
            cands.append((cand, tag, gap))
    # 道路に近い順。駐車場が入る一番手前の位置を採るため（奥に寄せすぎない）
    cands.sort(key=lambda c: c[2])
    return cands[:limit]


def place_parking(poly, building, road_idx, count, pw=2.5, pd=5.0,
                  setback_side=0.5, step=0.1):
    """建物の残りの空地（道路側）に駐車スペースを並べる。"""
    if count <= 0:
        return []
    road_e = edges_of(poly)[road_idx]
    (x1, y1), (x2, y2) = road_e
    theta = math.degrees(math.atan2(y2 - y1, x2 - x1))
    pivot = (x1, y1)
    free = poly.buffer(-setback_side, join_style=2).difference(building.buffer(0.5))
    if free.is_empty:
        return []
    reg = rotate(free, -theta, origin=pivot)
    minx, miny, maxx, maxy = reg.bounds
    spots, placed = [], []
    ny = max(1, int((maxy - miny - pd) / step) + 1)
    nx = max(1, int((maxx - minx - pw) / step) + 1)
    n = inward_normal(poly, road_e)
    rot_ny = (n[0] * math.sin(math.radians(-theta)) + n[1] * math.cos(math.radians(-theta)))
    ys = sorted([miny + i * step for i in range(ny)], key=lambda v: v * (1 if rot_ny > 0 else -1))
    # 道路に近い位置から順に埋める
    for yy in ys:
        for i in range(nx):
            xx = minx + i * step
            r = box(xx, yy, xx + pw, yy + pd)
            if not reg.contains(r):
                continue
            if any(r.intersects(q) for q in placed):
                continue
            placed.append(r)
            spots.append(rotate(r, theta, origin=pivot))
            if len(spots) >= count:
                return spots
        if len(spots) >= count:
            break
    return spots


# ---------------------------------------------------------------- 法規チェック

def code_checks(site_area, building_area, total_floor, bcr_limit, far_limit,
                road_width, setback_from_road, max_height, north_slope_base=5.0):
    """建蔽率・容積率・道路斜線・北側斜線のざっくり判定。"""
    out = []
    if site_area:
        bcr = building_area / site_area * 100
        out.append(('建蔽率', f'{bcr:.2f}%', f'{bcr_limit:.0f}%', bcr <= bcr_limit))
        if total_floor:
            far = total_floor / site_area * 100
            out.append(('容積率', f'{far:.2f}%', f'{far_limit:.0f}%', far <= far_limit))
    if road_width and max_height:
        lim = (setback_from_road + road_width) * 1.25
        out.append(('道路斜線', f'建物 {max_height:.3f}m',
                    f'制限 {lim:.2f}m', max_height <= lim))
    if max_height:
        lim = north_slope_base + 1.25 * setback_from_road
        out.append(('北側斜線(目安)', f'建物 {max_height:.3f}m',
                    f'制限 {lim:.2f}m', max_height <= lim))
    return out


# ---------------------------------------------------------------- DXF 書き出し

LAYERS = [
    ('敷地境界線', 5, 'CONTINUOUS'),
    ('道路境界線', 1, 'CONTINUOUS'),
    ('壁面後退線', 8, 'DASHED'),
    ('建物', 3, 'CONTINUOUS'),
    ('外壁線', 4, 'DASHED'),
    ('駐車場', 2, 'CONTINUOUS'),
    ('寸法', 6, 'CONTINUOUS'),
    ('文字', 7, 'CONTINUOUS'),
    ('方位', 7, 'CONTINUOUS'),
]


def write_dxf(path, poly, road_idx, building, eaves, parking, region,
              info, checks, point_names=None, north=90.0):
    doc = ezdxf.new('R2000', setup=True)
    doc.encoding = 'cp932'                    # Jw_cad 向け（日本語をSHIFT-JISで書く）
    doc.header['$INSUNITS'] = 6               # メートル
    msp = doc.modelspace()
    for name, color, ltype in LAYERS:
        if ltype not in doc.linetypes:
            ltype = 'CONTINUOUS'
        doc.layers.add(name, color=color, linetype=ltype)

    txt = {'height': 0.25, 'style': 'Standard'}
    ring = list(poly.exterior.coords)
    es = edges_of(poly)

    # 敷地境界線と辺長
    for i, e in enumerate(es):
        lay = '道路境界線' if i == road_idx else '敷地境界線'
        msp.add_line(e[0], e[1], dxfattribs={'layer': lay})
        mx, my = (e[0][0] + e[1][0]) / 2, (e[0][1] + e[1][1]) / 2
        n = inward_normal(poly, e)
        ang = math.degrees(math.atan2(e[1][1] - e[0][1], e[1][0] - e[0][0]))
        if ang > 90 or ang <= -90:
            ang += 180
        t = msp.add_text(f'{edge_len(e):.2f}', height=0.22,
                         dxfattribs={'layer': '寸法', 'rotation': ang})
        t.set_placement((mx - n[0] * 0.35, my - n[1] * 0.35), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)
        lbl = '道路境界線' if i == road_idx else '隣地境界線'
        t2 = msp.add_text(lbl, height=0.2, dxfattribs={'layer': '文字', 'rotation': ang})
        t2.set_placement((mx + n[0] * 0.45, my + n[1] * 0.45), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)

    # 境界点名
    for name, (px, py) in (point_names or []):
        msp.add_circle((px, py), 0.12, dxfattribs={'layer': '敷地境界線'})
        t = msp.add_text(name, height=0.22, dxfattribs={'layer': '文字'})
        t.set_placement((px + 0.25, py + 0.2))

    # 壁面後退線
    if region is not None and not region.is_empty:
        geoms = region.geoms if region.geom_type == 'MultiPolygon' else [region]
        for g in geoms:
            msp.add_lwpolyline(list(g.exterior.coords), close=True,
                               dxfattribs={'layer': '壁面後退線'})

    # 建物・軒先
    if building is not None:
        msp.add_lwpolyline(list(building.exterior.coords), close=True,
                           dxfattribs={'layer': '建物'})
        if eaves is not None and not eaves.is_empty:
            msp.add_lwpolyline(list(eaves.exterior.coords), close=True,
                               dxfattribs={'layer': '外壁線'})
        c = building.centroid
        t = msp.add_text('申請建物', height=0.3, dxfattribs={'layer': '文字'})
        t.set_placement((c.x, c.y), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)

        # 配置基準寸法（建物の各辺から敷地境界までの最短距離）
        for e in edges_of(building):
            mx, my = (e[0][0] + e[1][0]) / 2, (e[0][1] + e[1][1]) / 2
            n = inward_normal(building, e)
            out = (-n[0], -n[1])
            far = Point(mx + out[0] * 200, my + out[1] * 200)
            from shapely.geometry import LineString
            ray = LineString([(mx, my), (far.x, far.y)])
            hit = ray.intersection(poly.exterior)
            if hit.is_empty:
                continue
            pts = [hit] if hit.geom_type == 'Point' else list(getattr(hit, 'geoms', []))
            pts = [p for p in pts if p.geom_type == 'Point']
            if not pts:
                continue
            p = min(pts, key=lambda q: math.dist((mx, my), (q.x, q.y)))
            dist = math.dist((mx, my), (p.x, p.y))
            msp.add_line((mx, my), (p.x, p.y), dxfattribs={'layer': '寸法'})
            t = msp.add_text(f'{dist:.2f}', height=0.2, dxfattribs={'layer': '寸法'})
            t.set_placement(((mx + p.x) / 2, (my + p.y) / 2),
                            align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)

    # 駐車場
    for i, s in enumerate(parking, 1):
        msp.add_lwpolyline(list(s.exterior.coords), close=True, dxfattribs={'layer': '駐車場'})
        c = s.centroid
        t = msp.add_text(f'駐車 {i}', height=0.25, dxfattribs={'layer': '駐車場'})
        t.set_placement((c.x, c.y), align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)

    # 方位（北）
    minx, miny, maxx, maxy = poly.bounds
    ox, oy = maxx + 2.0, maxy - 1.0
    a = math.radians(north)
    msp.add_line((ox, oy), (ox + math.cos(a) * 1.5, oy + math.sin(a) * 1.5),
                 dxfattribs={'layer': '方位'})
    t = msp.add_text('N', height=0.4, dxfattribs={'layer': '方位'})
    t.set_placement((ox + math.cos(a) * 1.9, oy + math.sin(a) * 1.9),
                    align=ezdxf.enums.TextEntityAlignment.MIDDLE_CENTER)

    # 面積表・チェック結果
    lines = [f'{k}：{v}' for k, v in info]
    lines.append('')
    lines.append('［法規チェック（目安）］')
    for name, got, lim, ok in checks:
        lines.append(f'  {name}　{got}　/　{lim}　{"適合" if ok else "要検討"}')
    lines.append('')
    lines.append('※ 自動生成のたたき台です。建築士が確認・修正してください。')
    ty = miny - 1.2
    for ln in lines:
        t = msp.add_text(ln, height=0.28, dxfattribs={'layer': '文字'})
        t.set_placement((minx, ty))
        ty -= 0.45

    doc.saveas(path)
    return path


# ---------------------------------------------------------------- CLI

def parse_building(s):
    m = re.match(r'^\s*([\d.]+)\s*[x×*]\s*([\d.]+)\s*$', s)
    if not m:
        raise argparse.ArgumentTypeError('建物は 5.46x9.555 の形で指定してください（単位 m）')
    return float(m.group(1)), float(m.group(2))


def from_plan(name, plans_path='plans.json'):
    """過去の契約図から建物の大きさを持ってくる。"""
    if not os.path.exists(plans_path):
        sys.exit(f'{plans_path} が見つかりません（先に extract.py を実行してください）')
    P = json.load(open(plans_path))
    hit = [p for p in P if p.get('customer') and name in p['customer']]
    if not hit:
        sys.exit(f'「{name}」に一致する過去案件がありません')
    p = sorted(hit, key=lambda r: -(r.get('buildingArea') or 0))[0]
    fp = p.get('footprint') or {}
    if not fp.get('w'):
        sys.exit(f'「{p["customer"]}」は建築面積求積表から寸法が取れていません。'
                 f'--building W×D で指定してください')
    if not fp.get('exact'):
        print(f'※ 「{p["customer"]}」は単純な矩形ではありません（L字など）。'
              f'一番大きい矩形 {fp["w"]}×{fp["d"]}m を使います。'
              f'実際の建築面積は {p.get("buildingArea")}㎡ です。')
    return (fp['w'], fp['d']), p


def main():
    ap = argparse.ArgumentParser(description='測量図DXFから配置図のたたき台を作る')
    ap.add_argument('--dxf', required=True, help='測量図のDXF')
    ap.add_argument('--list', action='store_true', help='敷地候補を一覧して終了')
    ap.add_argument('--parcel', help='地番（例 1173-2）または候補番号（0始まり）')
    ap.add_argument('--layer', help='境界線のレイヤ名の一部（自動で見つからないとき）')
    ap.add_argument('--road-edge', type=int, help='道路に面する辺の番号（--list で確認）')
    ap.add_argument('--road-width', type=float, default=4.0, help='前面道路の幅員 m')
    ap.add_argument('--building', type=parse_building, help='建物の大きさ W×D（m）')
    ap.add_argument('--from-plan', help='過去案件の施主名から建物の大きさを取る')
    ap.add_argument('--total-floor', type=float, help='延床面積 ㎡（容積率の判定用）')
    ap.add_argument('--max-height', type=float, default=None, help='最高高さ m（斜線の判定用）')
    ap.add_argument('--setback-road', type=float, default=1.0, help='道路境界からの後退 m')
    ap.add_argument('--setback-side', type=float, default=0.5, help='隣地境界からの後退 m')
    ap.add_argument('--eaves', type=float, default=0.455,
                    help='軒の出 m（建築面積の矩形から内側に外壁線を描く）')
    ap.add_argument('--parking', type=int, default=2, help='駐車台数')
    ap.add_argument('--bcr', type=float, default=60.0, help='建蔽率の上限 %%')
    ap.add_argument('--far', type=float, default=200.0, help='容積率の上限 %%')
    ap.add_argument('--north', type=float, default=90.0, help='真北の向き（度、+X軸から反時計回り）')
    ap.add_argument('--out', default='配置図.dxf')
    a = ap.parse_args()

    parcels, layers = read_parcels(a.dxf, a.layer)
    if not parcels:
        sys.exit('敷地らしいポリゴンが見つかりませんでした。--layer で境界線のレイヤを指定してください。')

    if a.list or not a.road_edge and a.road_edge != 0:
        print(f'境界線に使ったレイヤ: {", ".join(layers) or "(全レイヤ)"}\n')
        for i, p in enumerate(parcels):
            lots = '/'.join(p['lots']) or '地番なし（図郭や区画線かもしれません）'
            print(f'[{i}] {lots}  {p["area"]:.2f}㎡ ({p["area"]/TSUBO:.2f}坪)  頂点{len(p["ring"])}')
            for j, e in enumerate(edges_of(p['poly'])):
                print(f'      辺{j}: {edge_len(e):6.2f}m  {e[0]} → {e[1]}')
        if not a.list:
            print('\n--road-edge で道路に面する辺の番号を指定してください。')
        return

    # 敷地を選ぶ
    site = parcels[0]
    if a.parcel:
        if a.parcel.isdigit() and int(a.parcel) < len(parcels):
            site = parcels[int(a.parcel)]
        else:
            exact = [p for p in parcels if p['lots'] == [a.parcel]]
            part = [p for p in parcels if any(a.parcel in l for l in p['lots'])]
            if not part:
                sys.exit(f'地番「{a.parcel}」の区画が見つかりません（--list で確認）')
            site = (exact or part)[0]
            if not exact:
                print(f'※ 地番「{a.parcel}」単独の区画が無いため、'
                      f'{"/".join(site["lots"])} の {site["area"]:.2f}㎡ を使います')
    poly = site['poly']
    es = edges_of(poly)
    if not 0 <= a.road_edge < len(es):
        sys.exit(f'--road-edge は 0〜{len(es)-1} の範囲で指定してください')

    # 建物の大きさ
    src_plan = None
    if a.from_plan:
        (w, dep), src_plan = from_plan(a.from_plan)
    elif a.building:
        w, dep = a.building
    else:
        sys.exit('--building か --from-plan のどちらかを指定してください')

    region = buildable_area(poly, a.road_edge, a.setback_road, a.setback_side)
    cands = building_candidates(poly, region, a.road_edge, w, dep)
    if not cands:
        sys.exit(f'建物 {w}×{dep}m は後退距離を守ると入りません。'
                 f'（建てられる範囲 {region.area:.1f}㎡）建物を小さくするか後退を見直してください。')
    # 道路に近い候補から順に見て、駐車場が必要台数入る一番手前の位置を採る
    building, orient, gap = cands[0]
    parking = place_parking(poly, building, a.road_edge, a.parking,
                            setback_side=a.setback_side)
    if a.parking > 0 and len(parking) < a.parking:
        best = None
        for c, t, g in cands:
            ps = place_parking(poly, c, a.road_edge, a.parking, setback_side=a.setback_side)
            if len(ps) >= a.parking:
                building, orient, gap, parking = c, t, g, ps
                best = True
                break
            if best is None or len(ps) > len(parking):
                building, orient, gap, parking = c, t, g, ps
        if best is None:
            print(f'※ 駐車 {a.parking}台は入りませんでした。'
                  f'置ける最大 {len(parking)}台の配置にしています。')
    # --building は建築面積（＝軒先の水平投影）。外壁線はその内側に描く
    eaves = building.buffer(-a.eaves, join_style=2)

    b_area = w * dep
    setback_from_road = gap

    checks = code_checks(poly.area, b_area, a.total_floor, a.bcr, a.far,
                         a.road_width, setback_from_road, a.max_height)

    info = [('敷地面積', f'{poly.area:.2f}㎡ ({poly.area/TSUBO:.2f}坪)'),
            ('地番', '/'.join(site['lots']) or '—'),
            ('建築面積', f'{b_area:.2f}㎡ ({b_area/TSUBO:.2f}坪)  {w:.3f}×{dep:.3f}m'),
            ('建物の向き', orient),
            ('道路境界からの後退', f'{setback_from_road:.2f}m'),
            ('前面道路', f'幅員 {a.road_width:.2f}m'),
            ('駐車スペース', f'{len(parking)}台（計画 {a.parking}台）')]
    if a.total_floor:
        info.insert(3, ('延床面積', f'{a.total_floor:.2f}㎡ ({a.total_floor/TSUBO:.2f}坪)'))
    if src_plan:
        info.append(('参照した過去案件', f'{src_plan["customer"]}（{src_plan.get("contractDate","")}）'))

    pnames = []
    write_dxf(a.out, poly, a.road_edge, building, eaves, parking, region,
              info, checks, pnames, a.north)

    for k, v in info:
        print(f'{k}：{v}')
    print()
    for name, got, lim, ok in checks:
        print(f'{name}　{got} / {lim}　{"適合" if ok else "要検討"}')
    if len(parking) < a.parking:
        print(f'\n※ 駐車スペースは {len(parking)}台しか置けませんでした。')
    print(f'\n→ {a.out} に書き出しました（Jw_cad で開けます）')


def _closest_on_edge(e, p):
    (x1, y1), (x2, y2) = e
    dx, dy = x2 - x1, y2 - y1
    L2 = dx * dx + dy * dy or 1.0
    t = max(0.0, min(1.0, ((p[0] - x1) * dx + (p[1] - y1) * dy) / L2))
    return (x1 + t * dx, y1 + t * dy)


if __name__ == '__main__':
    main()
