"""種子撮影トレイ写真の処理エンジン（GUI・コマンドラインの両方から使う）
  1. ArUco マーカーで真上視点に補正
  2. 封筒の窓からバーコード（Code128）を読む
  3. シャーレ内の種子を1粒ずつ分割して数える
  4. 輪郭を描いた確認画像を作る
"""
import itertools, cv2, numpy as np
from scipy import ndimage as ndi
from skimage.segmentation import watershed
from skimage.feature import peak_local_max, blob_log

# ---------------------------------------------------------------- トレイ定義
# key: (幅, 奥行) mm  / marker: (縁からマーカー中心, 黒部分の半分) mm
# regions: トレイ座標 mm（左下原点・上向き y）での dish 外接矩形と envelope 窓
TRAYS = {
    (206, 140):     dict(marker=(13.2, 7.5), regions=None),
    (214, 146):     dict(marker=(13.2, 7.5), regions=None),
    (214, 180):     dict(marker=(13.2, 7.5), regions={'dish': (2, 29, 124, 151), 'envelope': (129, 61, 185, 180)}),
    (180, 178):     dict(marker=(13.2, 7.5), regions={'dish': (29, 52, 151, 174), 'envelope': (59, 0, 121, 52)}),
    (188.8, 141.6): dict(marker=(10.2, 6.0), regions={'dish': (9.8, 9.8, 131.8, 131.8), 'envelope': (131.8, 39.8, 188.8, 101.8)}),
    (141.6, 188.8): dict(marker=(10.2, 6.0), regions={'dish': (9.8, 9.8, 131.8, 131.8), 'envelope': (39.8, 131.8, 101.8, 188.8)}),
}
LAYOUTS = [{0: 'FL', 1: 'FR', 2: 'BL', 3: 'BR'}, {0: 'FL', 1: 'FR', 3: 'BL', 2: 'BR'}]


# ---------------------------------------------------------------- マーカー
def detect_markers(img, work_max=2400):
    """大きな写真（12〜48MP）は長辺 work_max px に縮小して探し、座標を元の写真に戻す"""
    g = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    s0 = min(1.0, work_max / max(g.shape))
    if s0 < 1: g = cv2.resize(g, None, fx=s0, fy=s0, interpolation=cv2.INTER_AREA)
    d = cv2.aruco.getPredefinedDictionary(cv2.aruco.DICT_4X4_50)
    found = {}
    for v in (g, cv2.createCLAHE(3.0, (8, 8)).apply(g), cv2.GaussianBlur(g, (5, 5), 0)):
        for sc in (1.0, 0.5, 0.75, 1.5):
            im2 = cv2.resize(v, None, fx=sc, fy=sc, interpolation=cv2.INTER_AREA if sc < 1 else cv2.INTER_CUBIC)
            for win in ((3, 23, 10), (3, 53, 4), (5, 95, 10)):
                p = cv2.aruco.DetectorParameters()
                p.adaptiveThreshWinSizeMin, p.adaptiveThreshWinSizeMax, p.adaptiveThreshWinSizeStep = win
                p.cornerRefinementMethod = cv2.aruco.CORNER_REFINE_SUBPIX
                c, ids, _ = cv2.aruco.ArucoDetector(d, p).detectMarkers(im2)
                if ids is None: continue
                for i, cc in zip(ids.flatten(), c):
                    if int(i) in (0, 1, 2, 3): found.setdefault(int(i), []).append(cc[0] / sc / s0)
        if len(found) == 4: break
    out = []
    for i, cands in found.items():
        cen = np.array([c.mean(0) for c in cands])
        k = max(range(len(cands)), key=lambda k: (np.linalg.norm(cen - cen[k], axis=1) < 10).sum())
        out.append((i, cands[k]))
    return out


def _fit(found, W, D, layout):
    inset, half = TRAYS[(W, D)]['marker']
    cen = {'FL': (inset, inset), 'FR': (W - inset, inset), 'BL': (inset, D - inset), 'BR': (W - inset, D - inset)}
    best = None
    for rots in itertools.product(range(4), repeat=len(found)):
        src, dst = [], []
        for (i, cc), r in zip(found, rots):
            x, y = cen[layout[i]]
            sq = [(x - half, y + half), (x + half, y + half), (x + half, y - half), (x - half, y - half)]
            src += sq[r:] + sq[:r]; dst += cc.tolist()
        src, dst = np.float32(src), np.float32(dst)
        H, _ = cv2.findHomography(src, dst, 0)
        if H is None: continue
        err = float(np.linalg.norm(cv2.perspectiveTransform(src[None], H)[0] - dst, axis=1).mean())
        if best is None or err < best[0]: best = (err, H)
    if best and len(found) == 4:     # 4個そろえば中心4点で最終補正
        H = cv2.getPerspectiveTransform(np.float32([cen[layout[i]] for i, _ in found]), np.float32([c.mean(0) for _, c in found]))
        best = (best[0], H)
    return best


def locate_tray(img):
    """戻り値: dict(H=トレイmm→画像px, size=(W,D), err, markers) / 失敗時 None"""
    found = detect_markers(img)
    if len(found) < 2: return None
    cands = []
    for k in range(len(found), 1, -1):
        for sub in itertools.combinations(found, k):
            for size in TRAYS:
                for L in LAYOUTS:
                    r = _fit(list(sub), *size, L)
                    if r: cands.append((k, r[0], r[1], size, [i for i, _ in sub]))
        if any(c[0] == k and c[1] < 6 for c in cands) and k >= 3: break
    good = [c for c in cands if c[1] < 6] or cands
    k, err, H, size, used = min(good, key=lambda c: (-c[0], c[1]))
    return dict(H=H, size=size, err=err, markers=sorted(used))


def warp_region(img, H, box, ppm):
    """トレイ座標の矩形 box=(x0,y0,x1,y1) を ppm px/mm で切り出す（上が奥）"""
    x0, y0, x1, y1 = box
    M = H @ np.array([[1 / ppm, 0, x0], [0, -1 / ppm, y1], [0, 0, 1]])
    return cv2.warpPerspective(img, M, (int((x1 - x0) * ppm), int((y1 - y0) * ppm)),
                               flags=cv2.WARP_INVERSE_MAP | cv2.INTER_CUBIC, borderValue=(255, 255, 255))


def native_ppm(H, size):
    """元画像でのおおよその px/mm"""
    W, D = size
    p = cv2.perspectiveTransform(np.float32([[[0, 0], [W, 0], [0, D]]]), H)[0]
    return float(np.mean([np.linalg.norm(p[1] - p[0]) / W, np.linalg.norm(p[2] - p[0]) / D]))


# ---------------------------------------------------------------- Code128
_P = ("212222 222122 222221 121223 121322 131222 122213 122312 132212 221213 221312 231212 112232 122132 122231 "
      "113222 123122 123221 223211 221132 221231 213212 223112 312131 311222 321122 321221 312212 322112 322211 "
      "212123 212321 232121 111323 131123 131321 112313 132113 132311 211313 231113 231311 112133 112331 132131 "
      "113123 113321 133121 313121 211331 231131 213113 213311 213131 311123 311321 331121 312113 312311 332111 "
      "314111 221411 431111 111224 111422 121124 121421 141122 141221 112214 112412 122114 122411 142112 142211 "
      "241211 221114 413111 241112 134111 111242 121142 121241 114212 124112 124211 411212 421112 421211 212141 "
      "214121 412121 111143 111341 131141 114113 114311 411113 411311 113141 114131 311141 411131 211412 211214 "
      "211232").split()
_SIG = np.array([[int(p[0]) + int(p[1]), int(p[1]) + int(p[2]), int(p[2]) + int(p[3]), int(p[3]) + int(p[4])] for p in _P], float)


def _text(body):
    code = {103: 'A', 104: 'B', 105: 'C'}[body[0]]; t = ''
    for v in body[1:-1]:
        if code == 'C':
            if v < 100: t += f'{v:02d}'
            elif v == 100: code = 'B'
            elif v == 101: code = 'A'
        elif code == 'B':
            if v < 96: t += chr(v + 32)
            elif v == 99: code = 'C'
            elif v == 101: code = 'A'
        else:
            if v < 64: t += chr(v + 32)
            elif v < 96: t += chr(v - 64)
            elif v == 99: code = 'C'
            elif v == 100: code = 'B'
    return t


def _decode_runs(r):
    if len(r) < 19 or (len(r) - 7) % 6: return None
    n = (len(r) - 7) // 6; syms = []
    for k in range(n):
        g = np.asarray(r[6 * k:6 * k + 6], float); g = g / g.sum() * 11
        s = np.array([g[0] + g[1], g[1] + g[2], g[2] + g[3], g[3] + g[4]])
        syms.append(int(np.argmin(np.abs(_SIG - s).sum(1))))
    if syms[0] not in (103, 104, 105): return None
    if (syms[0] + sum(i * v for i, v in enumerate(syms[1:-1], 1))) % 103 != syms[-1]: return None
    t = _text(syms)
    return t if t and t.isprintable() else None


def _scan_lines(gray):
    """横向きのバーコード画像 gray を何本も横切って読む。{文字列: 票数}"""
    res = {}; h, w = gray.shape; g = gray.astype(np.float32)
    for y in np.linspace(0.2 * h, 0.8 * h, 9):
        for ang in np.arange(-3, 3.01, 0.75):
            xs = np.arange(0, w, 0.5); ys = y + (xs - w / 2) * np.tan(np.radians(ang))
            line = cv2.remap(g, xs.astype(np.float32)[None], ys.astype(np.float32)[None], cv2.INTER_LINEAR)[0]
            for pct in (5, 15):
                th = (np.percentile(line, pct) + np.percentile(line, 100 - pct)) / 2
                b = (line < th).astype(np.int8)
                idx = np.flatnonzero(np.diff(b)) + 1
                runs = np.diff(np.r_[0, idx, len(b)]); vals = b[np.r_[0, idx]]
                if len(runs) < 21: continue
                # 長い白（クワイエットゾーン）で区切り、いちばん本数の多い区間をバーコードとみなす
                quiet = (vals == 0) & (runs > 6 * np.median(runs))
                cuts = [-1] + list(np.flatnonzero(quiet)) + [len(runs)]
                a, z = max(zip(cuts[:-1], cuts[1:]), key=lambda p: p[1] - p[0])
                seg = runs[a + 1:z]; sv = vals[a + 1:z]
                if len(seg) and sv[0] == 0: seg, sv = seg[1:], sv[1:]
                if len(seg) and sv[-1] == 0: seg = seg[:-1]
                r = list(seg)
                for rr in (r, r[::-1]):
                    t = _decode_runs(rr)
                    if t: res[t] = res.get(t, 0) + 1
        if res and max(res.values()) >= 12: break
    return res


def _barcode_boxes(gray):
    """縦縞（横向きバーコード）の候補矩形を返す"""
    gx = cv2.Sobel(gray, cv2.CV_32F, 1, 0, ksize=3); gy = cv2.Sobel(gray, cv2.CV_32F, 0, 1, ksize=3)
    m = cv2.blur(np.abs(gx) - np.abs(gy), (9, 9))
    m = (m > max(20, np.percentile(m, 97) * 0.5)).astype(np.uint8)
    m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, cv2.getStructuringElement(cv2.MORPH_RECT, (25, 5)))
    m = cv2.morphologyEx(m, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (9, 3)))
    cnts = cv2.findContours(m, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)[0]
    boxes = []
    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)
        if w > 4 * h and w > 60: boxes.append((w * h, x, y, w, h))
    return [b[1:] for b in sorted(boxes, reverse=True)[:3]]


def _read_env(env):
    """封筒の窓画像から読めた文字列ごとの票数"""
    res = {}
    for rot in (None, cv2.ROTATE_90_CLOCKWISE):
        e = env if rot is None else cv2.rotate(env, rot)
        g = cv2.cvtColor(e, cv2.COLOR_BGR2GRAY)
        for (x, y, w, h) in _barcode_boxes(g):
            pad = int(0.15 * w)
            crop = g[max(0, y - 2):y + h + 2, max(0, x - pad):x + w + pad]
            if crop.size == 0: continue
            for t, v in _scan_lines(crop).items(): res[t] = res.get(t, 0) + v
        if res and max(res.values()) >= 12: break
    return res


def read_barcode(img, H, size, region):
    """封筒の窓を切り出してバーコードを読む。
    戻り値 (文字列 or None, 票数, 確認用画像, 確か=True/False)
    確か=False: 票が少ない、または別の読み方が拮抗（読み違いの可能性）→ 手で確認してもらう"""
    nat = native_ppm(H, size)
    p0 = float(np.clip(2.0 * nat, 16, 24))          # 写真の解像度によらず 16〜24 px/mm で読む
    total, env0 = {}, None
    for ppm in [p0] + [q for q in (20, 16, 24, 12, 28) if abs(q - p0) > 2]:
        env = warp_region(img, H, region, ppm)
        if env0 is None: env0 = env
        for t, v in _read_env(env).items(): total[t] = total.get(t, 0) + v
        if total and max(total.values()) >= 6: break
    if not total: return None, 0, env0, False
    ranked = sorted(total.items(), key=lambda kv: -kv[1])
    code, votes = ranked[0]; second = ranked[1][1] if len(ranked) > 1 else 0
    return code, votes, env0, bool(votes >= 3 and second * 3 < votes)


# ---------------------------------------------------------------- 種子計数
# 調整できるしきい値（GUI の「しきい値調整」で変更・保存できる）
DEFAULT_PARAMS = dict(
    dark_ratio=0.75,      # 暗さの基準。上げると色の薄い粒も拾う／下げると濃い部分だけ（0.75 が標準）
    min_radius_mm=0.55,   # これより細いものは種子ではない（文字・影・ゴミ）
    seg_max_mean=0.55,    # 1粒の平均の明るさがこれより明るければ種子ではない（影・凹み）
    rmm=60.5,             # 数える範囲の半径 mm（シャーレの内側）
)


def count_seeds(dish, ppm, params=None):
    """dish: シャーレ外接矩形の画像（ppm px/mm）。
    戻り値 (粒数, 輪郭つき画像, 各粒リスト)。各粒 = dict(x,y,area,dia,length,width,measured)
    measured=False は重なりが強く面積から粒数を推定した塊（大きさの統計には使わない）"""
    P = dict(DEFAULT_PARAMS, **(params or {}))
    rmm = P['rmm']
    h, w = dish.shape[:2]
    g = cv2.cvtColor(dish, cv2.COLOR_BGR2GRAY).astype(np.float32)
    k = int(6 * ppm) | 1                                  # 種子より大きい窓で「明るい側」を取り、背景とする
    bg = cv2.GaussianBlur(cv2.dilate(g, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (k, k))), (0, 0), 2 * ppm)
    norm = np.clip(g / (bg + 1) * 180, 0, 255).astype(np.uint8)
    yy, xx = np.mgrid[:h, :w]; rr = np.hypot(xx - w / 2, yy - h / 2)
    inside = rr < rmm * ppm; edge = inside & (rr >= (rmm - 0.5) * ppm)
    otsu, _ = cv2.threshold(norm, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    shift = (P['dark_ratio'] - 0.75) * 180                # 標準 0.75 からのずらし量（画素・粒の両方の基準をずらす）
    bw = norm < min(otsu + 1, 0.75 * 180) + shift
    bw = ndi.binary_opening(bw & inside, iterations=1)
    holes = ndi.binary_fill_holes(bw) & ~bw                  # 粒の中の小さな穴（光の反射）だけ埋める。
    hl, nh = ndi.label(holes)                                 # 水滴の跡などの輪の内側は埋めない（中の種子が消えるため）
    if nh:
        ha = ndi.sum(holes, hl, range(1, nh + 1))
        bw |= np.isin(hl, np.flatnonzero(ha < (1.5 * ppm) ** 2 * np.pi) + 1)
    lab0, n0 = ndi.label(bw)                              # 縁の帯にかかる小さな断片（縁の影）を落とす
    for i in set(np.unique(lab0[edge])) - {0}:
        comp = lab0 == i
        if (comp & ~edge).sum() < 3 * (ppm ** 2): bw &= ~comp
    bw &= ~edge
    out = dish.copy()
    cv2.circle(out, (w // 2, h // 2), int(rmm * ppm), (255, 0, 0), 2)
    lab, n = ndi.label(bw)
    if n == 0: return 0, out, []
    areas = ndi.sum(bw, lab, range(1, n + 1))
    maxr = ndi.maximum(ndi.distance_transform_edt(bw), lab, range(1, n + 1))
    bad = {i for i, A in enumerate(areas, 1) if A < (0.6 * ppm) ** 2} \
          | {i for i, m in enumerate(maxr, 1) if m < P['min_radius_mm'] * ppm}
    bw &= ~np.isin(lab, list(bad))
    lab, n = ndi.label(bw)
    if n == 0: return 0, out, []
    areas = ndi.sum(bw, lab, range(1, n + 1))
    single = float(np.median(areas[areas <= np.percentile(areas, 60)]))
    r = np.sqrt(single / np.pi)
    dist = ndi.distance_transform_edt(bw)
    ds = cv2.GaussianBlur(dist.astype(np.float32), (0, 0), max(0.5, 0.15 * r))
    inv = cv2.GaussianBlur(255 - norm, (0, 0), 1).astype(np.float32) / 255.0
    blobs = blob_log(inv, min_sigma=0.6 * r / 1.414, max_sigma=1.25 * r / 1.414, num_sigma=5, threshold=0.04, overlap=0.5)
    markers = np.zeros_like(lab); k = 0
    for y, x, _ in blobs:
        y, x = int(y), int(x)
        if bw[y, x] and dist[y, x] > 0.3 * r: k += 1; markers[y, x] = k
    for i in range(1, n + 1):
        comp = lab == i
        if not markers[comp].any():
            y, x = np.unravel_index(np.argmax(np.where(comp, dist, 0)), dist.shape); k += 1; markers[y, x] = k
    seg = watershed(-ds, markers, mask=bw, compactness=1.0)
    seeds = []
    objs = ndi.find_objects(seg)
    for s, sl in enumerate(objs, 1):
        if sl is None: continue
        y0, x0 = sl[0].start, sl[1].start
        m = seg[sl] == s; A = int(m.sum())
        if A == 0: continue
        if A < 0.3 * single or norm[sl][m].mean() > P['seg_max_mean'] * 180 + shift: continue
        kk = max(1, int(round(A / single))) if A > 1.7 * single else 1
        ys, xs = np.nonzero(m); cy, cx = ys.mean() + y0, xs.mean() + x0
        cnt = cv2.findContours(m.astype(np.uint8), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)[0]
        c = max(cnt, key=len)
        if kk >= 3 and A / (np.pi * cv2.minEnclosingCircle(c)[1] ** 2 + 1e-9) < 0.15:
            continue                     # 細長い弧（シャーレの縁の影など）を「何粒もの塊」と数えない
        eq = 2 * np.sqrt(A / kk / np.pi) / ppm                 # 同じ面積の円の直径
        L = Wd = eq
        if kk == 1 and len(c) >= 5:
            (_, _), (d1, d2), _ = cv2.fitEllipse(c)            # 長径・短径
            L, Wd = max(d1, d2) / ppm, min(d1, d2) / ppm
            if not (np.isfinite(L) and np.isfinite(Wd)) or Wd <= 0 or L > 3 * eq or Wd < 0.25 * eq: L = Wd = eq
        a_mm = A / kk / ppm ** 2
        for _ in range(kk):
            seeds.append(dict(x=cx / ppm, y=(h - cy) / ppm, area=a_mm, dia=2 * np.sqrt(a_mm / np.pi),
                              length=L, width=Wd, measured=(kk == 1)))
        cv2.drawContours(out, [c + [x0, y0]], -1, (0, 200, 0) if kk == 1 else (0, 140, 255), 2)
        cv2.circle(out, (int(cx), int(cy)), 2, (0, 0, 255), -1)
        if kk > 1: cv2.putText(out, str(kk), (int(cx) + 4, int(cy) - 4), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
    return len(seeds), out, seeds


def seed_stats(seeds, density=None):
    """大きさの統計。体積は回転楕円体（長径×短径×短径×π/6）で近似。
    density（g/cm³）を与えると推定千粒重 g も返す"""
    m = [s for s in seeds if s['measured']]
    st = dict(n_measured=len(m))
    if not m: return st
    A = np.array([s['area'] for s in m]); D = np.array([s['dia'] for s in m])
    L = np.array([s['length'] for s in m]); Wd = np.array([s['width'] for s in m])
    V = np.pi / 6 * L * Wd * Wd                                   # mm³
    st.update(area_mean=A.mean(), area_sd=A.std(ddof=1) if len(A) > 1 else 0.0,
              dia_mean=D.mean(), dia_sd=D.std(ddof=1) if len(D) > 1 else 0.0,
              length_mean=L.mean(), width_mean=Wd.mean(), aspect_mean=(L / np.maximum(Wd, 1e-6)).mean(),
              vol_mean=V.mean(), vol1000_cm3=V.mean() * 1000 / 1000.0)
    st['area_cv'] = 100 * st['area_sd'] / st['area_mean']
    if density: st['tgw_est'] = st['vol1000_cm3'] * density           # g / 1000粒
    return st


HEAD_H = 70


def _header(width, n, code):
    """確認画像の見出し。文字は黒、バーコードが読めなかったときだけ赤"""
    head = np.full((HEAD_H, width, 3), 255, np.uint8)
    cv2.putText(head, f"Seeds: {n}", (15, 48), cv2.FONT_HERSHEY_SIMPLEX, 1.2, (35, 35, 35), 2, cv2.LINE_AA)
    x = 15 + cv2.getTextSize(f"Seeds: {n}", cv2.FONT_HERSHEY_SIMPLEX, 1.2, 2)[0][0] + 60
    cv2.putText(head, f"Barcode: {code}" if code else "Barcode: not read", (x, 48), cv2.FONT_HERSHEY_SIMPLEX, 1.2,
                (35, 35, 35) if code else (30, 30, 200), 2, cv2.LINE_AA)
    return head


def relabel_check(check_img, n, code):
    """保存済みの確認画像の見出し（種子数・バーコード）だけを書き換える"""
    out = check_img.copy(); out[:HEAD_H] = _header(out.shape[1], n, code)
    return out


def compose_check(ann, env, n, code):
    """確認画像: 左にシャーレ（輪郭つき）、右に封筒の窓、上に結果"""
    e = cv2.resize(env, (int(env.shape[1] * ann.shape[0] / env.shape[0]), ann.shape[0])) if env.shape[0] else env
    if e.shape[1] > ann.shape[1] * 0.8:
        e = cv2.resize(e, (int(ann.shape[1] * 0.8), int(e.shape[0] * ann.shape[1] * 0.8 / e.shape[1])))
        e = cv2.copyMakeBorder(e, 0, ann.shape[0] - e.shape[0], 0, 0, cv2.BORDER_CONSTANT, value=(255, 255, 255))
    canvas = np.hstack([ann, np.full((ann.shape[0], 10, 3), 255, np.uint8), e])
    return np.vstack([_header(canvas.shape[1], n, code), canvas])


# ---------------------------------------------------------------- 画像の読み込み（HEIC 対応）
HEIC_EXTS = ('.heic', '.heif')


def _open_heic(path):
    try:
        from pillow_heif import register_heif_opener
    except ImportError:
        raise RuntimeError('HEICを読むにはpillow-heifが必要です（pip install pillow-heif）')
    from PIL import Image, ImageOps
    register_heif_opener()
    im = Image.open(path)
    exif = im.info.get('exif')
    if exif:                                   # 画素はこのあと正しい向きにするので、EXIF の向きは「回転なし」に
        try:
            ex = Image.Exif(); ex.load(exif); ex[0x0112] = 1; exif = ex.tobytes()
        except Exception:
            exif = None
    return ImageOps.exif_transpose(im).convert('RGB'), exif


def load_image(path):
    """JPEG/PNG/HEIC を BGR 画像で返す（日本語パス対応）"""
    if path.lower().endswith(HEIC_EXTS):
        im, _ = _open_heic(path)
        return cv2.cvtColor(np.asarray(im), cv2.COLOR_RGB2BGR)
    return cv2.imdecode(np.fromfile(path, np.uint8), cv2.IMREAD_COLOR)


def heic_to_jpeg(src, dst, quality=95):
    """HEIC を JPEG に変換して保存（撮影日時などの EXIF も引き継ぐ）"""
    im, exif = _open_heic(src)
    kw = dict(quality=quality)
    if exif: kw['exif'] = exif
    im.save(dst, 'JPEG', **kw)


# ---------------------------------------------------------------- 1枚まるごと
def prepare(path, ppm=10):
    """補正・バーコード読取・シャーレ切り出しまで（しきい値を変えての再計数はこの結果を使い回す）"""
    try:
        img = load_image(path)
    except Exception as e:
        return dict(ok=False, msg=str(e))
    if img is None: return dict(ok=False, msg='画像を開けません')
    loc = locate_tray(img)
    if loc is None: return dict(ok=False, msg='マーカーが見つかりません')
    regs = TRAYS[loc['size']]['regions']
    if regs is None: return dict(ok=False, msg=f"未登録のトレイです（{loc['size'][0]}×{loc['size'][1]}mm）")
    code, votes, env, sure = read_barcode(img, loc['H'], loc['size'], regs['envelope'])
    dish = warp_region(img, loc['H'], regs['dish'], ppm)
    return dict(ok=True, barcode=code, votes=votes, barcode_sure=sure, env=env, dish=dish, ppm=ppm,
                tray=f"{loc['size'][0]}x{loc['size'][1]}", marker_err=round(loc['err'], 2), markers=loc['markers'])


def process(path, ppm=10, params=None, density=None):
    """戻り値 dict: ok, msg, barcode, votes, barcode_sure, seeds, seed_list, stats, tray, marker_err, annotated, ann_dish, env, dish"""
    r = prepare(path, ppm)
    if not r['ok']: return r
    n, ann, seeds = count_seeds(r['dish'], ppm, params)
    r.update(msg='' if r['barcode'] else 'バーコードを読み取れません', seeds=n, seed_list=seeds, ann_dish=ann,
             stats=seed_stats(seeds, density), annotated=compose_check(ann, r['env'], n, r['barcode']))
    return r
