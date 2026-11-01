#!/usr/bin/env python3
"""種子カウンター

トレイで撮ったシャーレの写真から種子を数え、封筒のバーコードを読み取り、結果をExcelに1枚1行で記録する。
写真のコピーはバーコード（対応表があれば系統名）の名前で保存し、輪郭を描いた確認画像も保存する。
バーコードを読めなかった写真は、ラベルの画像を見ながら手で入力できる。

コマンドラインでの使い方:
    python seed_counter_app.py --cli 出力フォルダ 結果.xlsx [--map 対応表.xlsx] 写真1.jpg 写真2.jpg ...
"""
import os, sys, csv, json, math, shutil, threading, queue, datetime, base64
from collections import Counter
import cv2, numpy as np
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import seedtray_core as core

EXTS = ('.jpg', '.jpeg', '.png', '.heic', '.heif', '.tif', '.tiff', '.bmp')
SETTINGS = os.path.join(os.path.expanduser('~'), '.seed_counter_settings.json')
STATS = [('平均面積（mm²）', 'area_mean', 2), ('面積SD（mm²）', 'area_sd', 2), ('面積CV（%）', 'area_cv', 1),
         ('平均直径（mm）', 'dia_mean', 2), ('直径SD（mm）', 'dia_sd', 2), ('平均長径（mm）', 'length_mean', 2),
         ('平均短径（mm）', 'width_mean', 2), ('長短比', 'aspect_mean', 2), ('平均体積（mm³）', 'vol_mean', 2),
         ('千粒体積（cm³）', 'vol1000_cm3', 3), ('推定千粒重（g）', 'tgw_est', 3)]
H_DENS, H_NMEAS, H_SRC, H_NEW, H_CHECK, H_INPUT = '密度（g/cm³）', '測定粒数', '元のファイル名', 'ファイル名', '確認画像', 'バーコードの入力'
H_TGW = '推定千粒重（g）'
HEADER = (['No', 'バーコード', '系統名', '種子数'] + [h for h, _, _ in STATS] +
          [H_DENS, H_NMEAS, H_SRC, H_NEW, H_CHECK, H_INPUT, '状態', '処理日時', 'しきい値', '元のパス'])
OLD_NAMES = {'平均面積（mm²）': '平均面積 mm²', '面積SD（mm²）': '面積SD', '面積CV（%）': '面積CV %', '平均直径（mm）': '平均直径 mm',
             '直径SD（mm）': '直径SD', '平均長径（mm）': '平均長径 mm', '平均短径（mm）': '平均短径 mm', '平均体積（mm³）': '平均体積 mm³',
             '千粒体積（cm³）': '千粒体積 cm³', H_TGW: '推定千粒重 g', H_DENS: '密度 g/cm³', H_NMEAS: '大きさを測った粒数',
             H_NEW: '新しいファイル名', H_INPUT: 'バーコード入力'}           # 前の版で作った Excel の見出し
WIDTHS = {'No': 6, 'バーコード': 16, '系統名': 18, '種子数': 8, H_SRC: 24, H_NEW: 20, H_CHECK: 24,
          H_INPUT: 14, '状態': 26, '処理日時': 17, 'しきい値': 30, '元のパス': 45}
RECOUNT_FIELDS = ['種子数'] + [h for h, _, _ in STATS] + [H_DENS, H_NMEAS, '処理日時', 'しきい値']
ST_OK, ST_MAN = 'OK', '手入力'
ST_NEED, ST_CHECK = '要確認（読み取り不可）', '要確認（読み取りが不確か）'
ST_NOMAP = '要確認（対応表にない）'
ST_FAIL = '失敗'
EXCEL_BUSY = 'Excelファイルに書き込めません。Excelで開いている場合は閉じてから、もう一度実行してください。'
LOCK = threading.RLock()        # Excel の読み書きとファイル名の決定は同時に1つだけ（処理中の手入力でも壊れないように）
WIN_RESERVED = {'CON', 'PRN', 'AUX', 'NUL'} | {f'COM{i}' for i in range(1, 10)} | {f'LPT{i}' for i in range(1, 10)}


def needs_input(rec):
    return rec['status'].startswith(('要入力', '要確認'))


# ------------------------------------------------------------ バーコード → 系統名の対応表
def _cell_text(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): v = int(v)       # Excel で数値になった番号
    return str(v).strip()


def _read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        try: return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
        finally: wb.close()
    if ext in ('.csv', '.txt', '.tsv'):
        for enc in ('utf-8-sig', 'cp932'):                       # Excel で保存した CSV（Shift-JIS）も読む
            try:
                with open(path, encoding=enc, newline='') as f: text = f.read()
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError('CSVの文字コードを判別できません。UTF-8かShift_JISで保存してください。')
        delim = '\t' if ext == '.tsv' or text.count('\t') > text.count(',') else ','
        return list(csv.reader(text.splitlines(), delimiter=delim))
    raise ValueError('対応表は.xlsxか.csv形式で保存してください。')


class NameMap:
    """1列目（または見出し「バーコード」）= バーコード、2列目（または見出し「系統名」など）= 系統名"""
    CODE_KEYS = ('バーコード', 'barcode', 'コード', 'code')
    NAME_KEYS = ('系統', '品種', 'サンプル', 'sample', 'name', '名前', '名称', 'line', 'accession', 'id')

    def __init__(self):
        self.clear()

    def clear(self):
        self.path, self.by_code, self.by_name, self.dups = '', {}, {}, []

    @property
    def loaded(self):
        return bool(self.by_code)

    @staticmethod
    def norm(v):
        return _cell_text(v).upper()

    def load(self, path):
        rows = [r for r in _read_table(path) if any(_cell_text(c) for c in r)]
        if not rows: raise ValueError('対応表にデータがありません。')
        hdr = [_cell_text(c).lower() for c in rows[0]]
        def find(keys, skip=None):
            for j, h in enumerate(hdr):
                if j != skip and h and any(k.lower() in h for k in keys): return j
            return None
        ci = find(self.CODE_KEYS); ni = find(self.NAME_KEYS, skip=ci)
        if ci is not None and ni is not None: data = rows[1:]
        else: ci, ni, data = 0, 1, rows[1:] if any(k in ' '.join(hdr) for k in ('バーコード', 'barcode', '系統')) else rows
        by_code, by_name, dups = {}, {}, []
        for r in data:
            r = list(r) + [None, None]
            code, name = _cell_text(r[ci]), _cell_text(r[ni])
            if not code or not name: continue
            k = code.upper()
            if k in by_code and by_code[k][1] != name: dups.append(code)
            by_code[k] = (code, name); by_name.setdefault(name.upper(), (code, name))
        if not by_code:
            raise ValueError('バーコードと系統名を読み取れませんでした。1列目にバーコード、2列目に系統名を入力してください。')
        self.path, self.by_code, self.by_name, self.dups = path, by_code, by_name, dups
        return len(by_code)

    def lookup(self, code):
        hit = self.by_code.get(self.norm(code)) if code else None
        return hit[1] if hit else ''

    def resolve(self, text):
        """手入力の値 → (バーコード, 系統名, 'code'|'name'|'')。系統名を打ってもバーコードに戻す"""
        t = text.strip(); k = self.norm(t)
        if k in self.by_code: return t, self.by_code[k][1], 'code'
        if k in self.by_name: return self.by_name[k][0], self.by_name[k][1], 'name'
        return t, '', ''


MAP = NameMap()


def make_template(path, rows=None):
    """対応表の見本を作る。rows=[(バーコード, 系統名)] があればそれを入れる（無ければ記入例）"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook(); ws = wb.active; ws.title = '対応表'
    ws.append(['バーコード', '系統名'])
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496'); c.alignment = Alignment(horizontal='center')
    for code, name in rows or []: ws.append([code, name])
    for r in range(2, max(1000, ws.max_row + 200)): ws.cell(r, 1).number_format = '@'   # 先頭の 0 が消えないよう文字列に
    ws.column_dimensions['A'].width = 22; ws.column_dimensions['B'].width = 34; ws.freeze_panes = 'A2'
    note = wb.create_sheet('使い方')
    for line in ['1列目にバーコード、2列目に系統名を入力します。1行に1袋です。',
                 'バーコードの列は文字列の書式にしてあるので、先頭の0は消えません。',
                 '系統名に含まれる / \\ : * ? " < > | は、ファイル名では _ に置き換わります。',
                 '同じバーコードが2行あると、下の行を使います。',
                 '保存したら、種子カウンターの［ツール］メニューの［対応表を開く］で読み込みます。', '',
                 '記入例（「対応表」シートに、このように書きます）']:
        note.append([line])
    for ex in (['バーコード', '系統名'], ['3960B2BF1E', 'SG301_26_6133'], ['369EFE4F09', 'SG307_26_0300']): note.append(ex)
    note.column_dimensions['A'].width = 22; note.column_dimensions['B'].width = 24
    wb.save(path)


def head_label(barcode, name):
    """確認画像の見出し用（画像には日本語を描けないので、英数字の系統名だけ併記）"""
    if name and name.isascii() and name != barcode: return f'{barcode} ({name})' if barcode else name
    return barcode


# ------------------------------------------------------------ 設定
def default_outdir():
    desk = None
    if sys.platform.startswith('win'):
        try:                    # OneDrive でデスクトップが移動していても本当のデスクトップを使う
            import ctypes
            buf = ctypes.create_unicode_buffer(1024)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf) == 0: desk = buf.value
        except Exception:
            pass
    if not desk: desk = os.path.join(os.path.expanduser('~'), 'Desktop')
    if not os.path.isdir(desk): desk = os.path.expanduser('~')
    return os.path.join(desk, 'seed_results')


def load_settings():
    s = dict(params=dict(core.DEFAULT_PARAMS), density='', outdir='', xlsx='', mapfile='', geometry='')
    try:
        with open(SETTINGS, encoding='utf-8') as f: d = json.load(f)
        for k in core.DEFAULT_PARAMS:
            try: s['params'][k] = float(d.get('params', {})[k])
            except (KeyError, TypeError, ValueError): pass
        for k in ('density', 'outdir', 'xlsx', 'mapfile', 'geometry'):
            if isinstance(d.get(k), (str, int, float)): s[k] = str(d[k])
    except Exception:
        pass
    return s


def save_settings(s):
    try:
        with open(SETTINGS, 'w', encoding='utf-8') as f: json.dump(s, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def parse_density(text):
    """空欄 → None。数字でない・0以下 → ValueError"""
    text = str(text).strip()
    if not text: return None
    v = float(text)
    if not math.isfinite(v) or v <= 0: raise ValueError(text)
    return v


def params_text(p):
    return f"暗さ{p['dark_ratio']:.2f}/太さ{p['min_radius_mm']:.2f}/影{p['seg_max_mean']:.2f}/半径{p['rmm']:.1f}"


# ------------------------------------------------------------ ファイル
def safe_name(code):
    s = ''.join(ch if ch.isalnum() or ch in '-_.' else '_' for ch in str(code).strip()).strip('. ')[:100]
    if not s: s = '_'
    if s.split('.')[0].upper() in WIN_RESERVED: s = '_' + s
    return s


def _key(p):
    return os.path.normcase(os.path.abspath(p))


def unique_pair(folder, stem, ext, exclude=()):
    """folder/stem.ext と folder/stem_check.jpg が両方とも空いている名前（exclude のファイルは空き扱い）"""
    k = 1
    while True:
        st = stem if k == 1 else f'{stem}_{k}'
        p, c = os.path.join(folder, st + ext), os.path.join(folder, st + '_check.jpg')
        if all(not os.path.exists(x) or _key(x) in exclude for x in (p, c)): return p, c
        k += 1


def imread(path):
    try:
        return cv2.imdecode(np.fromfile(path, np.uint8), cv2.IMREAD_COLOR)
    except Exception:
        return None


def imwrite(path, img, q=92):
    ok, buf = cv2.imencode('.jpg', img, [cv2.IMWRITE_JPEG_QUALITY, q])
    if not ok: raise IOError('画像を保存できません: ' + path)
    buf.tofile(path)                  # 日本語パス対応


def jpg_bytes(img, maxside=1400, q=88):
    """メモリ節約: 手入力画面用の封筒画像は縮小 JPEG で持つ"""
    h, w = img.shape[:2]; s = min(1.0, maxside / max(h, w))
    if s < 1: img = cv2.resize(img, (int(w * s), int(h * s)), interpolation=cv2.INTER_AREA)
    return cv2.imencode('.jpg', img, [cv2.IMWRITE_JPEG_QUALITY, q])[1].tobytes()


def from_jpg(b):
    return cv2.imdecode(np.frombuffer(b, np.uint8), cv2.IMREAD_COLOR) if b else None


def new_rec(path, outdir, xlsx, params):
    return dict(path=path, src=os.path.basename(path), outdir=outdir, xlsx=xlsx, barcode='', name='', seeds='', stats={},
                density=None, new='', check='', input='', status='', params=params_text(params),
                excel_row=None, env_jpg=None)


def process_one(path, outdir, xlsx, params, density):
    """1枚処理して rec（Excel 1行分＋手入力用の情報）を返す"""
    rec = new_rec(path, outdir, xlsx, params)
    r = core.process(path, params=params, density=density)
    if not r.get('ok'):
        rec['status'] = f"{ST_FAIL}（{r.get('msg', '')}）"; return rec
    heic = rec['src'].lower().endswith(core.HEIC_EXTS)
    ext = '.jpg' if heic else (os.path.splitext(rec['src'])[1].lower() or '.jpg')
    name = MAP.lookup(r['barcode'])
    stem = safe_name(name or r['barcode']) if r['barcode'] else 'NG_' + safe_name(os.path.splitext(rec['src'])[0])
    ann = core.relabel_check(r['annotated'], r['seeds'], head_label(r['barcode'], name)) if name else r['annotated']
    with LOCK:
        newp, chk = unique_pair(outdir, stem, ext)
        if heic: core.heic_to_jpeg(path, newp)
        else: shutil.copy2(path, newp)
        imwrite(chk, ann)
    if not r['barcode']: status, how = ST_NEED, ''
    elif not r.get('barcode_sure', True): status, how = ST_CHECK, '自動（要確認）'
    elif MAP.loaded and not name: status, how = ST_NOMAP, '自動'
    else: status, how = ST_OK, '自動'
    rec.update(barcode=r['barcode'] or '', name=name, seeds=r['seeds'], stats=r['stats'], density=density,
               new=os.path.basename(newp), check=os.path.basename(chk), input=how, status=status,
               env_jpg=jpg_bytes(r['env']))
    return rec


# ------------------------------------------------------------ Excel
def _num(v, d=None):
    if v is None or v == '': return None
    try: v = float(v)
    except (TypeError, ValueError): return None
    if not math.isfinite(v): return None          # NaN/inf を書くと Excel が「修復」を求める
    return round(v, d) if d is not None else v


def row_values(rec, now):
    st = rec['stats']; v = {'バーコード': rec['barcode'] or None, '系統名': rec.get('name') or None,
                            '種子数': rec['seeds'] if isinstance(rec['seeds'], (int, np.integer)) else None}
    for h, k, d in STATS: v[h] = _num(st.get(k), d)
    v[H_DENS] = _num(rec.get('density'), 3)
    v[H_NMEAS] = st.get('n_measured')
    v.update({H_SRC: rec['src'], H_NEW: rec['new'] or None, H_CHECK: rec['check'] or None,
              H_INPUT: rec['input'] or None, '状態': rec['status'], '処理日時': now,
              'しきい値': rec['params'], '元のパス': rec['path']})
    return v


def _open_wb(xlsx):
    """戻り値 (wb, ws, 列名→列番号)。見出しが足りなければ右端に足す（前の版で作った Excel にも追記できる）"""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    if os.path.exists(xlsx):
        wb = load_workbook(xlsx); ws = wb['種子数'] if '種子数' in wb.sheetnames else wb.create_sheet('種子数')
    else:
        wb = Workbook(); ws = wb.active; ws.title = '種子数'
    cols = {str(c.value): c.column for c in ws[1] if c.value is not None}
    if not cols: ws.freeze_panes = 'E2'
    nxt = max(cols.values(), default=0) + 1
    for h in HEADER:
        if h in cols: continue
        if OLD_NAMES.get(h) in cols:                 # 前の版の見出しは新しい名前に書き換えて使う
            cols[h] = cols.pop(OLD_NAMES[h]); ws.cell(1, cols[h], h); continue
        c = ws.cell(1, nxt, h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.column_dimensions[c.column_letter].width = WIDTHS.get(h, 10)
        cols[h] = nxt; nxt += 1
    ws.row_dimensions[1].height = 32
    return wb, ws, cols


def _last_row(ws):
    r = ws.max_row                       # 書式だけ残った空行は数えない
    while r > 1 and all(c.value is None for c in ws[r]): r -= 1
    return r


def _write(ws, rr, cols, vals):
    for h, v in vals.items():
        if h in cols: ws.cell(rr, cols[h], v)


def _style_row(ws, rr, cols, rec):
    from openpyxl.styles import Font, PatternFill
    base = os.path.dirname(os.path.abspath(rec['xlsx']))
    for key, h in (('new', H_NEW), ('check', H_CHECK)):
        c = ws.cell(rr, cols[h])
        if rec[key]:
            target = os.path.join(rec['outdir'], rec[key])
            try: link = os.path.relpath(target, base).replace(os.sep, '/')   # 相対リンク: フォルダごと移動しても切れない
            except ValueError: link = target                                  # Excel と写真が別ドライブ
            c.hyperlink = link; c.font = Font(color='0563C1', underline='single')
        else:
            c.hyperlink = None; c.font = Font()
    st = rec['status']                       # 色は対応が必要な行だけ（要確認=黄、失敗=赤）
    fill = (PatternFill('solid', fgColor='FDE7E9') if st.startswith(ST_FAIL) else
            PatternFill('solid', fgColor='FFF4CE') if needs_input(rec) else PatternFill(fill_type=None))
    for j in range(1, max(cols.values()) + 1): ws.cell(rr, j).fill = fill


def _save_wb(wb, xlsx):
    """一時ファイルに保存してから置き換える（保存中に止まっても Excel が壊れない）"""
    d = os.path.dirname(os.path.abspath(xlsx)); os.makedirs(d, exist_ok=True)
    tmp = os.path.join(d, '.~saving_' + os.path.basename(xlsx))
    wb.save(tmp)
    try:
        os.replace(tmp, xlsx)
    except Exception:
        try: os.remove(tmp)
        except OSError: pass
        raise


def excel_writable(xlsx):
    if os.path.exists(xlsx):
        try:
            with open(xlsx, 'r+b'): return True     # Windows では Excel で開いていると失敗する
        except OSError:
            return False
    d = os.path.dirname(os.path.abspath(xlsx))
    try: os.makedirs(d, exist_ok=True)
    except OSError: return False
    return os.access(d, os.W_OK)


def append_rows(xlsx, recs):
    with LOCK:
        wb, ws, cols = _open_wb(xlsx)
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M'); last = _last_row(ws)
        for i, rec in enumerate(recs, 1):
            _write(ws, last + i, cols, dict(row_values(rec, now), No=last - 1 + i)); _style_row(ws, last + i, cols, rec)
        _save_wb(wb, xlsx)
        for i, rec in enumerate(recs, 1): rec['excel_row'] = last + i


def _find_row(ws, cols, rec, name):
    """Excel の行を「元のパス＋新しいファイル名」で探す（並べ替え・行削除をしても別の行を書き換えない）"""
    def match(r):
        return (ws.cell(r, cols['元のパス']).value == rec['path'] and
                ws.cell(r, cols[H_NEW]).value == (name or None))
    r0 = rec.get('excel_row')
    if r0 and 1 < r0 <= ws.max_row and match(r0): return r0
    for r in range(_last_row(ws), 1, -1):
        if match(r): return r
    return None


def update_row(rec, old_name, fields=None):
    with LOCK:
        wb, ws, cols = _open_wb(rec['xlsx'])
        now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
        rr = _find_row(ws, cols, rec, old_name); vals = row_values(rec, now)
        if rr is None:                       # 行が消されていたら新しい行として足す
            rr = _last_row(ws) + 1; vals['No'] = rr - 1
        elif fields:
            vals = {k: vals[k] for k in fields}
        _write(ws, rr, cols, vals); _style_row(ws, rr, cols, rec); _save_wb(wb, rec['xlsx'])
        rec['excel_row'] = rr


def _rename(rec, barcode, name, how, status):
    """バーコード・系統名を変えて、写真・確認画像の名前と Excel の行を書き換える（途中で失敗したら元に戻す）"""
    if not excel_writable(rec['xlsx']): raise PermissionError(EXCEL_BUSY)
    od, old_name, old_chk = rec['outdir'], rec['new'], rec['check']
    oldp, oldc = os.path.join(od, old_name), os.path.join(od, old_chk)
    with LOCK:
        newp, newc = unique_pair(od, safe_name(name or barcode), os.path.splitext(old_name)[1] or '.jpg',
                                 exclude={_key(oldp), _key(oldc)})
        img = imread(oldc) if os.path.exists(oldc) else None
        saved = {k: rec[k] for k in ('barcode', 'name', 'new', 'check', 'input', 'status')}
        moved = wrote = False
        try:
            if os.path.exists(oldp) and _key(oldp) != _key(newp): os.replace(oldp, newp); moved = True
            if img is not None: imwrite(newc, core.relabel_check(img, rec['seeds'], head_label(barcode, name))); wrote = True
            rec.update(barcode=barcode, name=name, new=os.path.basename(newp), check=os.path.basename(newc),
                       input=how, status=status)
            update_row(rec, old_name)
        except Exception:
            rec.update(saved)
            if wrote:
                if _key(oldc) == _key(newc): imwrite(oldc, img)
                else:
                    try: os.remove(newc)
                    except OSError: pass
            if moved: os.replace(newp, oldp)
            raise
        if img is not None and _key(oldc) != _key(newc) and os.path.exists(oldc):
            try: os.remove(oldc)
            except OSError: pass
    return rec


def apply_manual_code(rec, text):
    """手入力した値（バーコード、または対応表の系統名）で名前と Excel を書き換える"""
    if not text.strip(): raise ValueError('バーコードまたは系統名を入力してください。')
    barcode, name, kind = MAP.resolve(text)
    same = barcode.upper() == rec['barcode'].upper() and rec['input'].startswith('自動')
    if same and (name or not MAP.loaded):
        how, status = '自動（確認済み）', ST_OK                 # 読み取った値をそのまま確認
    elif MAP.loaded and not name:
        how, status = '手入力（対応表になし）', ST_MAN
    else:
        how, status = '手入力（系統名）' if kind == 'name' else '手入力', ST_MAN
    return _rename(rec, barcode, name, how, status)


def remap_rec(rec):
    """対応表を読み込み直したあと、処理済みの写真に系統名を反映。変わったら True"""
    if not rec['check'] or not rec['barcode']: return False
    name = MAP.lookup(rec['barcode'])
    status = rec['status']
    if status in (ST_OK, ST_NOMAP): status = ST_OK if (name or not MAP.loaded) else ST_NOMAP
    if name == rec['name'] and status == rec['status']: return False
    _rename(rec, rec['barcode'], name, rec['input'], status)
    return True


def recount_one(rec, params, density):
    """今のしきい値で数え直し、確認画像と Excel の同じ行を書き換える"""
    if not rec['check']: raise RuntimeError('トレイを認識できなかった写真は数え直せません')
    if not excel_writable(rec['xlsx']): raise PermissionError(EXCEL_BUSY)
    r = core.process(rec['path'], params=params, density=density)
    if not r.get('ok'): raise RuntimeError(r.get('msg', ''))
    saved = {k: rec[k] for k in ('seeds', 'stats', 'params', 'density')}
    rec.update(seeds=r['seeds'], stats=r['stats'], params=params_text(params), density=density)
    try:
        imwrite(os.path.join(rec['outdir'], rec['check']),
                core.compose_check(r['ann_dish'], r['env'], r['seeds'], head_label(rec['barcode'], rec['name']) or None))
        update_row(rec, rec['new'], fields=RECOUNT_FIELDS)
    except Exception:
        rec.update(saved); raise
    if rec['env_jpg'] is None: rec['env_jpg'] = jpg_bytes(r['env'])
    return rec


# ------------------------------------------------------------ CLI
def cli(argv):
    outdir, xlsx, files = argv[0], argv[1], argv[2:]
    s = load_settings(); os.makedirs(outdir, exist_ok=True)
    mapfile = s['mapfile']
    if len(files) >= 2 and files[0] == '--map': mapfile, files = files[1], files[2:]
    MAP.clear()
    if mapfile:
        try: print('対応表:', mapfile, MAP.load(mapfile), '件')
        except Exception as e: MAP.clear(); print(f'対応表を読み込めないため、対応表なしで処理します（{e}）')
    try: dens = parse_density(s['density'])
    except ValueError: dens = None
    recs = []
    for f in files:
        try:
            rec = process_one(f, outdir, xlsx, s['params'], dens)
        except Exception as e:
            rec = new_rec(f, outdir, xlsx, s['params']); rec['status'] = f'{ST_FAIL}（{e}）'
        recs.append(rec); st = rec['stats']
        print(f"{rec['src']} -> {rec['new'] or '-'}  seeds={rec['seeds']}  barcode={rec['barcode'] or '-'}  name={rec['name'] or '-'}  "
              f"dia={st.get('dia_mean', 0):.2f}mm  {rec['status']}")
    append_rows(xlsx, recs); print('Excel:', xlsx)
    return recs


# ------------------------------------------------------------ GUI
APP_NAME = '種子カウンター'
APP_VERSION = '1.4'
C_WARN, C_FAIL, C_ERRTEXT = '#FFF4CE', '#FDE7E9', '#A4262C'      # 要確認・失敗の行と、注意の文字（Windows の通知の配色）


def gui():
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import ttk, filedialog, messagebox, simpledialog

    S = load_settings()
    root = tk.Tk(); root.title(APP_NAME); root.minsize(960, 600)
    try: root.geometry(S.get('geometry') or '1280x800')
    except tk.TclError: root.geometry('1280x800')
    style = ttk.Style(root)
    if style.theme_use() == 'default' and 'clam' in style.theme_names(): style.theme_use('clam')
    line = tkfont.nametofont('TkDefaultFont').metrics('linespace')
    style.configure('Treeview', rowheight=max(20, int(line * 1.5)))          # 行の高さを文字に合わせる（下線が欠けない）
    MAC = sys.platform == 'darwin'
    MOD, MODN = ('Command', 'Cmd') if MAC else ('Control', 'Ctrl')

    files, recs, previews, prepared, jobs = [], {}, {}, {}, {}
    q = queue.Queue(); state = dict(running=False, stop=False, closing=False, photo=None, shown=None)
    outdir = tk.StringVar(value=S['outdir'] or default_outdir())
    xlsx = tk.StringVar(value=S['xlsx'] or os.path.join(outdir.get(), '種子数.xlsx'))
    density = tk.StringVar(value=S['density'])
    mapvar = tk.StringVar(value='')
    EMPTY_PREVIEW = '写真を選ぶと、確認画像をここに表示します。'

    def info(m, parent=None): messagebox.showinfo(APP_NAME, m, parent=parent or root)
    def error(m, parent=None): messagebox.showerror(APP_NAME, m, parent=parent or root)
    def ask(m, parent=None): return messagebox.askyesno(APP_NAME, m, parent=parent or root)
    def confirm(m, parent=None): return messagebox.askokcancel(APP_NAME, m, parent=parent or root)
    def status(text): msg.config(text=text)

    def to_photo(img, W, H):
        h, w = img.shape[:2]; s = min(W / w, H / h)
        im = cv2.resize(img, (max(1, int(w * s)), max(1, int(h * s))), interpolation=cv2.INTER_AREA)
        ok, buf = cv2.imencode('.png', im)
        return tk.PhotoImage(data=base64.b64encode(buf.tobytes()))

    def debounce(key, ms, fn):
        if jobs.get(key): root.after_cancel(jobs[key])
        jobs[key] = root.after(ms, lambda: (jobs.pop(key, None), fn()))

    class Tooltip:
        """マウスを重ねると説明を出す（画面に説明文を並べない）"""
        def __init__(self, widget, text, delay=600):
            self.w, self.text, self.delay, self.tip, self.job = widget, text, delay, None, None
            widget.bind('<Enter>', self.schedule, add='+'); widget.bind('<Leave>', self.hide, add='+')
            widget.bind('<ButtonPress>', self.hide, add='+')
        def schedule(self, _=None):
            self.hide(); self.job = self.w.after(self.delay, self.show)
        def show(self):
            text = self.text() if callable(self.text) else self.text
            x, y = self.w.winfo_rootx() + 8, self.w.winfo_rooty() + self.w.winfo_height() + 4
            self.tip = tk.Toplevel(self.w); self.tip.wm_overrideredirect(True); self.tip.wm_geometry(f'+{x}+{y}')
            tk.Label(self.tip, text=text, justify='left', background='#FFFFE1', foreground='#000000', relief='solid',
                     borderwidth=1, padx=6, pady=3, wraplength=380).pack()
        def hide(self, _=None):
            if self.job: self.w.after_cancel(self.job); self.job = None
            if self.tip: self.tip.destroy(); self.tip = None

    # --- メニュー
    menubar = tk.Menu(root); root.config(menu=menubar)
    def add_menu(label, key):
        m = tk.Menu(menubar, tearoff=0); menubar.add_cascade(label=f'{label}({key})', underline=len(label) + 1, menu=m)
        return m
    m_file, m_run, m_tool, m_help = add_menu('ファイル', 'F'), add_menu('処理', 'P'), add_menu('ツール', 'T'), add_menu('ヘルプ', 'H')
    def item(menu, label, cmd, key='', bind=None):
        menu.add_command(label=label, command=cmd, accelerator=key)
        if bind: root.bind(bind, lambda e: (cmd(), 'break')[1])
        return menu.index('end')
    item(m_file, '写真を追加...', lambda: cmd_add_photos(), f'{MODN}+O', f'<{MOD}-o>')
    item(m_file, 'フォルダを追加...', lambda: cmd_add_folder(), f'{MODN}+Shift+O', f'<{MOD}-O>')
    m_file.add_separator()
    item(m_file, '選んだ写真をリストから外す', lambda: remove_selected(), 'Delete')
    item(m_file, 'リストを空にする', lambda: clear())
    m_file.add_separator()
    item(m_file, '保存先フォルダを開く', lambda: open_path(outdir.get()))
    item(m_file, 'Excelファイルを開く', lambda: open_path(xlsx.get()))
    m_file.add_separator()
    item(m_file, '終了', lambda: on_close(), f'{MODN}+Q', f'<{MOD}-q>')
    i_start = item(m_run, '開始', lambda: start(), 'F5', '<F5>')
    i_stop = item(m_run, '中止', lambda: stop(), 'Esc', '<Escape>')
    m_run.add_separator()
    item(m_run, 'バーコードを入力...', lambda: manual_selected(), 'Enter')
    i_check = item(m_run, '要確認の写真を順に入力...', lambda: manual_all(), f'{MODN}+E', f'<{MOD}-e>')
    item(m_run, '選んだ写真を数え直す', lambda: recount_selected(), f'{MODN}+R', f'<{MOD}-r>')
    item(m_tool, 'しきい値の調整', lambda: threshold_window(), f'{MODN}+T', f'<{MOD}-t>')
    item(m_tool, '千粒重から密度を計算...', lambda: calibrate())
    m_tool.add_separator()
    item(m_tool, '対応表を開く...', lambda: choose_map())
    item(m_tool, '対応表を解除', lambda: clear_map())
    item(m_tool, '対応表の見本を保存...', lambda: save_template())
    item(m_help, '使い方', lambda: help_window(), 'F1', '<F1>')
    item(m_help, 'バージョン情報', lambda: info(f'{APP_NAME} {APP_VERSION}\nシャーレの種子を数え、封筒のバーコードと合わせてExcelに記録します。'))
    m_run.entryconfigure(i_stop, state='disabled'); m_run.entryconfigure(i_check, state='disabled')

    # --- ツールバー（よく使う操作だけ）
    tb = ttk.Frame(root, padding=(8, 6, 8, 2)); tb.pack(fill='x')
    def tbtn(text, cmd, side='left', padx=(0, 4)):
        b = ttk.Button(tb, text=text, command=cmd); b.pack(side=side, padx=padx); return b
    tbtn('写真を追加...', lambda: cmd_add_photos()); tbtn('フォルダを追加...', lambda: cmd_add_folder())
    ttk.Separator(tb, orient='vertical').pack(side='left', fill='y', padx=8, pady=3)
    b_start = tbtn('開始', lambda: start()); b_stop = tbtn('中止', lambda: stop()); b_stop.state(['disabled'])
    ttk.Separator(tb, orient='vertical').pack(side='left', fill='y', padx=8, pady=3)
    b_check = tbtn('要確認の写真を入力...', lambda: manual_all()); b_check.state(['disabled'])
    b_thr = tbtn('しきい値の調整', lambda: threshold_window(), side='right', padx=(4, 0))
    Tooltip(b_thr, lambda: '現在の設定：' + params_text(S['params']))

    # --- 設定（ラベルと入力欄を1列にそろえる。補足はマウスを重ねたときだけ）
    form = ttk.Frame(root, padding=(10, 6, 10, 8)); form.pack(fill='x')
    def field(row, label, var, tip, readonly=False):
        l = ttk.Label(form, text=label); l.grid(row=row, column=0, sticky='w', padx=(0, 12), pady=2)
        e = ttk.Entry(form, textvariable=var, state='readonly' if readonly else 'normal'); e.grid(row=row, column=1, sticky='we', pady=2)
        Tooltip(l, tip); Tooltip(e, tip)
        return e
    def choose_outdir():
        d = filedialog.askdirectory(parent=root, title='保存先フォルダ')
        if d: outdir.set(d); xlsx.set(os.path.join(d, '種子数.xlsx'))
    def choose_xlsx():
        f = filedialog.asksaveasfilename(parent=root, title='Excelファイル', defaultextension='.xlsx',
                                         filetypes=[('Excelブック', '*.xlsx')], confirmoverwrite=False)
        if f: xlsx.set(f)
    field(0, '保存先フォルダ', outdir, '写真のコピーと確認画像を保存するフォルダです。')
    ttk.Button(form, text='参照...', command=choose_outdir).grid(row=0, column=2, sticky='w', padx=(6, 0))
    field(1, 'Excelファイル', xlsx, '結果を1枚1行で記録します。既存のファイルを選ぶと、末尾に追記します。')
    ttk.Button(form, text='参照...', command=choose_xlsx).grid(row=1, column=2, sticky='w', padx=(6, 0))
    field(2, '系統名の対応表', mapvar, 'バーコードと系統名を並べた表（ExcelまたはCSV）です。読み込むと、ファイル名とExcelに系統名を使います。', True)
    mf = ttk.Frame(form); mf.grid(row=2, column=2, sticky='w', padx=(6, 0))
    ttk.Button(mf, text='参照...', command=lambda: choose_map()).pack(side='left')
    ttk.Button(mf, text='解除', command=lambda: clear_map()).pack(side='left', padx=(4, 0))
    maplab = ttk.Label(mf, text=''); maplab.pack(side='left', padx=(8, 0))
    dl = ttk.Label(form, text='種子の密度'); dl.grid(row=3, column=0, sticky='w', padx=(0, 12), pady=2)
    df = ttk.Frame(form); df.grid(row=3, column=1, sticky='w', pady=2)
    de = ttk.Entry(df, textvariable=density, width=8); de.pack(side='left')
    ttk.Label(df, text='g/cm³').pack(side='left', padx=(4, 12))
    ttk.Button(df, text='千粒重から計算...', command=lambda: calibrate()).pack(side='left')
    for w in (dl, de): Tooltip(w, '入力すると推定千粒重を計算します。空欄なら計算しません。')
    form.columnconfigure(1, weight=1)

    # --- 状態表示（下端）
    sbar = ttk.Frame(root); sbar.pack(side='bottom', fill='x')
    ttk.Separator(root, orient='horizontal').pack(side='bottom', fill='x')
    msg = ttk.Label(sbar, text='写真を追加してください。', anchor='w', padding=(10, 3)); msg.pack(side='left', fill='x', expand=True)
    pb = ttk.Progressbar(sbar, length=160, mode='determinate'); pb.pack(side='right', padx=(4, 10), pady=3)
    counts = ttk.Label(sbar, text='', padding=(10, 3)); counts.pack(side='right')

    # --- 一覧と確認画像
    mid = ttk.Panedwindow(root, orient='horizontal'); mid.pack(fill='both', expand=True, padx=8, pady=(0, 6))
    lf = ttk.Frame(mid); mid.add(lf, weight=3)
    fm = tkfont.nametofont('TkDefaultFont'); W = lambda text: fm.measure(text) + 28
    COLS = [('src', '写真', W('関係ない画像.jpg'), 'w', True), ('barcode', 'バーコード', W('3960B2BF1E'), 'w', False),
            ('name', '系統名', W('SG301_26_6133'), 'w', False), ('seeds', '種子数', W('種子数'), 'e', False),
            ('dia', '平均直径（mm）', W('平均直径（mm）'), 'e', False), ('status', '状態', W('要確認（読み取り不可）'), 'w', True)]
    tv = ttk.Treeview(lf, columns=[c[0] for c in COLS], show='headings', selectmode='extended')
    def sort_by(col, rev=False):
        def key(k):
            v = tv.set(k, col)
            try: return (0, float(v), '')
            except ValueError: return (1, 0.0, v)
        for i, k in enumerate(sorted(tv.get_children(''), key=key, reverse=rev)): tv.move(k, '', i)
        tv.heading(col, command=lambda: sort_by(col, not rev))
    for c, t, w, anc, stretch in COLS:
        tv.heading(c, text=t, anchor=anc, command=lambda c=c: sort_by(c))
        tv.column(c, width=w, minwidth=w if stretch else 40, anchor=anc, stretch=stretch)   # 狭いときは横スクロール
    tv.tag_configure('warn', background=C_WARN); tv.tag_configure('fail', background=C_FAIL)
    sb = ttk.Scrollbar(lf, orient='vertical', command=tv.yview); xsb = ttk.Scrollbar(lf, orient='horizontal', command=tv.xview)
    tv.configure(yscrollcommand=sb.set, xscrollcommand=xsb.set)
    tv.grid(row=0, column=0, sticky='nsew'); sb.grid(row=0, column=1, sticky='ns'); xsb.grid(row=1, column=0, sticky='ew')
    lf.rowconfigure(0, weight=1); lf.columnconfigure(0, weight=1)
    rf = ttk.Frame(mid); mid.add(rf, weight=2)
    prev = ttk.Label(rf, text=EMPTY_PREVIEW, anchor='center', justify='center'); prev.pack(fill='both', expand=True)
    root.after(300, lambda: mid.sashpos(0, int(mid.winfo_width() * 0.6)))

    ctx = tk.Menu(root, tearoff=0)                     # 右クリックのメニュー
    ctx.add_command(label='バーコードを入力...', command=lambda: manual_selected())
    ctx.add_command(label='数え直す', command=lambda: recount_selected())
    ctx.add_separator()
    ctx.add_command(label='元の写真を開く', command=lambda: open_selected('path'))
    ctx.add_command(label='確認画像を開く', command=lambda: open_selected('check'))
    ctx.add_separator()
    ctx.add_command(label='リストから外す', command=lambda: remove_selected())
    def popup(e):
        iid = tv.identify_row(e.y)
        if iid and iid not in tv.selection(): tv.selection_set(iid)
        if tv.selection(): ctx.tk_popup(e.x_root, e.y_root)
    tv.bind('<Button-2>' if MAC else '<Button-3>', popup)
    if MAC: tv.bind('<Control-Button-1>', popup)

    # --- 表示の更新
    def update_title(*_):
        root.title(f'{os.path.basename(xlsx.get()) or "種子数.xlsx"} - {APP_NAME}')
    xlsx.trace_add('write', update_title); update_title()

    def update_counts():
        done = [recs[p] for p in files if p in recs]
        need = sum(needs_input(r) for r in done); fail = sum(r['status'].startswith(ST_FAIL) for r in done)
        parts = [f'写真{len(files)}枚'] + ([f'処理済み{len(done)}'] if done else []) + \
                ([f'要確認{need}'] if need else []) + ([f'失敗{fail}'] if fail else [])
        counts.config(text='　'.join(parts))
        b_check.state(['!disabled' if need and not state['running'] else 'disabled'])
        m_run.entryconfigure(i_check, state='normal' if need else 'disabled')

    def show(path=None):
        path = path or state['shown']
        img = imread(path) if path else None
        if img is None: return
        state['shown'] = path
        state['photo'] = to_photo(img, max(200, rf.winfo_width() - 10), max(200, rf.winfo_height() - 10))
        prev.config(image=state['photo'], text='')
    rf.bind('<Configure>', lambda e: debounce('prev', 250, show))

    def refresh_row(p):
        rec = recs[p]; st = rec['stats']
        tag = ('fail',) if rec['status'].startswith(ST_FAIL) else ('warn',) if needs_input(rec) else ()
        tv.item(p, values=(rec['src'], rec['barcode'], rec['name'], rec['seeds'],
                           f"{st['dia_mean']:.2f}" if 'dia_mean' in st else '', rec['status']), tags=tag)

    def on_select(_):
        sel = tv.selection()
        if sel and sel[0] in previews: show(previews[sel[0]])
    tv.bind('<<TreeviewSelect>>', on_select)
    tv.bind('<Double-1>', lambda e: (lambda iid: iid in recs and manual_entry(iid))(tv.identify_row(e.y)))
    tv.bind('<Return>', lambda e: manual_selected()); tv.bind('<KP_Enter>', lambda e: manual_selected())
    tv.bind('<Delete>', lambda e: remove_selected())
    if MAC: tv.bind('<BackSpace>', lambda e: remove_selected())

    def open_path(p):
        if not p or not os.path.exists(p): status(f'見つかりません：{p}'); return
        if sys.platform.startswith('win'): os.startfile(p)
        else:
            import subprocess
            try: subprocess.Popen(['open' if MAC else 'xdg-open', p])
            except OSError: pass

    def open_selected(kind):
        for p in tv.selection()[:5]:
            r = recs.get(p)
            if kind == 'path': open_path(p)
            elif r and r['check']: open_path(os.path.join(r['outdir'], r['check']))

    # --- リスト
    def cmd_add_photos():
        add(filedialog.askopenfilenames(parent=root, title='写真を追加', filetypes=[
            ('画像', ' '.join('*' + e for e in EXTS + tuple(x.upper() for x in EXTS))), ('すべてのファイル', '*.*')]))

    def cmd_add_folder():
        add_folder(filedialog.askdirectory(parent=root, title='写真のフォルダを追加'))

    def add(paths):
        n0 = len(files)
        for p in paths or []:
            p = os.path.abspath(p)
            if p.lower().endswith(EXTS) and p not in files:
                files.append(p); tv.insert('', 'end', iid=p, values=(os.path.basename(p), '', '', '', '', '未処理'))
        update_counts()
        if len(files) > n0: status(f'{len(files) - n0}枚を追加しました。［開始］を押すと処理します。')

    def add_folder(d):
        if not d: return
        if _key(d) == _key(outdir.get()):
            error('保存先フォルダは追加できません。元の写真があるフォルダを選んでください。'); return
        paths = sorted(os.path.join(d, f) for f in os.listdir(d)
                       if f.lower().endswith(EXTS) and '_check' not in f and not f.startswith('.'))
        if not paths: status(f'「{os.path.basename(d)}」に写真がありません。'); return
        add(paths)

    def remove_selected():
        if state['running']: status('処理中はリストから外せません。'); return
        sel = tv.selection()
        for p in sel:
            tv.delete(p); files.remove(p); recs.pop(p, None); previews.pop(p, None)
        if sel:
            update_counts(); status(f'{len(sel)}枚をリストから外しました。保存したファイルとExcelはそのままです。')

    def clear():
        if state['running']: return
        files.clear(); recs.clear(); previews.clear(); prepared.clear(); tv.delete(*tv.get_children())
        state.update(shown=None, photo=None); prev.config(image='', text=EMPTY_PREVIEW)
        update_counts(); status('リストを空にしました。')

    # --- 対応表
    def map_status():
        if MAP.loaded:
            d = len(MAP.dups)
            maplab.config(text=f'{len(MAP.by_code)}件' + (f'（重複{d}件）' if d else ''), foreground=C_ERRTEXT if d else '')
        else:
            maplab.config(text='', foreground='')

    def load_map(path, quiet=False):
        try:
            MAP.load(path)
        except Exception as e:
            MAP.clear(); map_status()
            if not quiet: error(f'対応表を読み込めません。\n{path}\n\n{e}')
            return False
        mapvar.set(path); S['mapfile'] = path; save_settings(S); map_status()
        return True

    def choose_map():
        f = filedialog.askopenfilename(parent=root, title='対応表を開く',
                                       filetypes=[('Excel・CSV', '*.xlsx *.csv *.txt *.tsv'), ('すべてのファイル', '*.*')])
        if not f or not load_map(f): return
        text = f'対応表を読み込みました（{len(MAP.by_code)}件）。'
        if MAP.dups:
            text += f'\n同じバーコードに異なる系統名があります。下の行を使います：' + '、'.join(MAP.dups[:6])
        done = [r for r in recs.values() if r['check'] and r['barcode']]
        if done and not state['running'] and ask(text + f'\n\n処理済みの{len(done)}枚のファイル名と系統名を、この対応表で書き換えますか？'):
            remap_all(done)
        elif MAP.dups:
            info(text)
        else:
            status(text)

    def remap_all(done):
        if not all(excel_writable(x) for x in {r['xlsx'] for r in done}): error(EXCEL_BUSY); return
        n, errs = 0, []
        for r in done:
            try: n += remap_rec(r)
            except Exception as e: errs.append(f"{r['src']}：{e}")
            if tv.exists(r['path']): refresh_row(r['path'])
            if r['check']: previews[r['path']] = os.path.join(r['outdir'], r['check'])
        sel = tv.selection()
        if sel and sel[0] in previews: show(previews[sel[0]])
        update_counts()
        nomap = sum(r['status'] == ST_NOMAP for r in done)
        status(f'{n}枚のファイル名を書き換えました。' + (f'対応表にないバーコードが{nomap}枚あります。' if nomap else ''))
        if errs: error('次の写真は書き換えられませんでした。\n' + '\n'.join(errs[:10]))

    def clear_map():
        MAP.clear(); mapvar.set(''); S['mapfile'] = ''; save_settings(S); map_status(); status('対応表を解除しました。')

    def save_template():
        codes = {}
        for r in recs.values():                     # 処理済みの写真があれば、読み取ったバーコードを入れておく
            if r['barcode'] and r['barcode'] not in codes: codes[r['barcode']] = r.get('name') or ''
        f = filedialog.asksaveasfilename(parent=root, title='対応表の見本を保存', initialfile='対応表.xlsx',
                                         defaultextension='.xlsx', filetypes=[('Excelブック', '*.xlsx')])
        if not f: return
        try:
            make_template(f, sorted(codes.items()) or None)
        except PermissionError:
            error('保存できません。同じ名前のファイルをExcelで開いている場合は閉じてください。'); return
        except Exception as e:
            error(f'保存できません。\n{e}'); return
        status(f'見本を保存しました（読み取ったバーコード{len(codes)}件入り）：{f}' if codes else f'見本を保存しました：{f}')
        open_path(f)

    if S['mapfile'] and not load_map(S['mapfile'], quiet=True):     # 前回の対応表が無くなっていたら、なしで動く
        S['mapfile'] = ''; save_settings(S)
        root.after(500, lambda: status('前回の対応表が見つからないため、対応表を使わずに処理します。'))
    elif MAP.loaded:
        mapvar.set(S['mapfile'])

    # --- 処理
    def busy(on, n=0):
        state['running'] = on
        b_start.state(['disabled' if on else '!disabled']); b_stop.state(['!disabled' if on else 'disabled'])
        m_run.entryconfigure(i_start, state='disabled' if on else 'normal')
        m_run.entryconfigure(i_stop, state='normal' if on else 'disabled')
        if on: state['stop'] = False; pb.config(maximum=max(1, n), value=0)
        update_counts()

    def get_density(parent=None):
        try: return True, parse_density(density.get())
        except ValueError:
            error('種子の密度には正の数を入力してください。空欄にすると推定千粒重を計算しません。', parent); return False, None

    def worker(todo, od, xl, params, dens):
        batch, pending, err = [], [], None
        for i, p in enumerate(todo, 1):
            if state['stop']: break
            q.put(('status', p, '処理中'))
            try:
                rec = process_one(p, od, xl, params, dens)
            except Exception as e:
                rec = new_rec(p, od, xl, params); rec['status'] = f'{ST_FAIL}（{e}）'
            batch.append(rec); pending.append(rec)
            try: append_rows(xl, pending); pending, err = [], None
            except Exception as e: err = e                 # 次の写真のときにまとめて再試行
            q.put(('row', p, rec, i, len(todo)))
        q.put(('done', batch, xl, pending, err))

    def start():
        if state['running']: return
        if not files: status('写真がありません。［写真を追加］で追加してください。'); return
        todo = [p for p in files if p not in recs]
        if not todo: status('未処理の写真はありません。'); return
        ok, dens = get_density()
        if not ok: return
        od, xl = outdir.get().strip(), xlsx.get().strip()
        if not od or not xl: error('保存先フォルダとExcelファイルを指定してください。'); return
        if not xl.lower().endswith('.xlsx'): xl += '.xlsx'; xlsx.set(xl)
        try: os.makedirs(od, exist_ok=True)
        except OSError as e: error(f'保存先フォルダを作成できません。\n{e}'); return
        if not excel_writable(xl): error(EXCEL_BUSY); return
        if S['mapfile'] and not load_map(S['mapfile'], quiet=True):   # 対応表は毎回読み直す（途中で書き足してもよい）
            if not ask(f'対応表を読み込めません。\n{S["mapfile"]}\n\n対応表を使わずに処理しますか？'): return
            clear_map()
        S.update(density=density.get().strip(), outdir=od, xlsx=xl); save_settings(S)
        busy(True, len(todo)); status(f'{len(todo)}枚を処理しています。')
        threading.Thread(target=worker, args=(todo, od, xl, dict(S['params']), dens), daemon=True).start()

    def stop():
        if state['running'] and not state['stop']:
            state['stop'] = True; b_stop.state(['disabled']); status('中止します。処理中の写真が終わると止まります。')

    def dup_groups():
        c = Counter(r['barcode'] for r in recs.values() if r['barcode'])
        return sum(1 for v in c.values() if v > 1)

    def quit_app():
        S['geometry'] = root.geometry(); save_settings(S); root.destroy()

    def finish(batch, xl, pending, err):
        busy(False)
        while pending:                       # Excel に書けなかった分: 再試行するか、別のファイルに保存する
            if messagebox.askretrycancel(APP_NAME, f'Excelに{len(pending)}枚分を書き込めませんでした。\n{err}\n\n'
                                         'Excelでファイルを開いている場合は閉じて［再試行］を押してください。'
                                         '［キャンセル］を押すと別のファイルに保存します。', parent=root):
                try: append_rows(xl, pending); pending = []
                except Exception as e: err = e
            else:
                bak = os.path.join(os.path.dirname(os.path.abspath(xl)), f'種子数_未保存分_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx')
                try:
                    for r in pending: r['xlsx'] = bak
                    append_rows(bak, pending); info(f'別のファイルに保存しました。\n{bak}')
                except Exception as e:
                    error(f'保存できませんでした。\n{e}')
                break
        if state['closing']: quit_app(); return
        need = sum(needs_input(r) for r in batch); fail = sum(r['status'].startswith(ST_FAIL) for r in batch)
        dups = dup_groups()
        text = f'{len(batch)}枚を処理しました。' + ('途中で中止しました。' if state['stop'] else '')
        extra = ([f'要確認{need}枚'] if need else []) + ([f'失敗{fail}枚'] if fail else []) + ([f'同じバーコード{dups}組'] if dups else [])
        if extra: text += '、'.join(extra) + '。'
        if need: text += '［要確認の写真を入力］で確認できます。'
        status(text); update_counts()
        if need: b_check.focus_set()

    def poll():
        try:
            while True:
                m = q.get_nowait()
                if m[0] == 'status':
                    if tv.exists(m[1]): tv.set(m[1], 'status', m[2]); tv.see(m[1])
                elif m[0] == 'row':
                    _, p, rec, i, n = m
                    recs[p] = rec; pb.config(value=i); status(f'{i}/{n}枚を処理しました（{rec["src"]}）。')
                    if tv.exists(p): refresh_row(p)
                    if rec['check']:
                        previews[p] = os.path.join(rec['outdir'], rec['check'])
                        sel = tv.selection()
                        if not sel or sel[0] == p: show(previews[p])
                    update_counts()
                elif m[0] == 'recounted':
                    _, p, i, n = m; pb.config(value=i); status(f'{i}/{n}枚を数え直しました。')
                    if tv.exists(p): refresh_row(p)
                    if state['shown'] == previews.get(p): show(previews[p])
                elif m[0] == 'done':
                    finish(*m[1:])
                elif m[0] == 'cb':
                    m[1]()
        except queue.Empty:
            pass
        root.after(120, poll)

    def on_close():
        if state['running']:
            if not confirm('処理中です。いま処理している写真が終わったら終了します。'): return
            state.update(stop=True, closing=True); status('終了しています。'); return
        quit_app()
    root.protocol('WM_DELETE_WINDOW', on_close)

    # --- 数え直し（ファイル名はそのまま、確認画像と Excel の同じ行を書き換える）
    def recount_selected():
        if state['running']: status('処理中は数え直せません。'); return
        sel = [p for p in tv.selection() if p in recs and recs[p]['check']]
        if not sel: info('数え直す写真をリストで選んでください。'); return
        ok, dens = get_density()
        if not ok: return
        if not all(excel_writable(x) for x in {recs[p]['xlsx'] for p in sel}): error(EXCEL_BUSY); return
        if not confirm(f'選んだ{len(sel)}枚を、現在のしきい値で数え直します。\n確認画像とExcelの行を上書きします。'): return
        busy(True, len(sel)); params = dict(S['params'])
        def work():
            errs = []
            for i, p in enumerate(sel, 1):
                if state['stop']: break
                q.put(('status', p, '数え直し中'))
                try: recount_one(recs[p], params, dens)
                except Exception as e: errs.append(f"{recs[p]['src']}：{e}")
                q.put(('recounted', p, i, len(sel)))
            q.put(('cb', lambda: recount_done(len(sel), errs)))
        threading.Thread(target=work, daemon=True).start()

    def recount_done(n, errs):
        busy(False)
        if state['closing']: quit_app(); return
        status(f'{n - len(errs)}枚を数え直しました。')
        if errs: error('次の写真は数え直せませんでした。\n' + '\n'.join(errs[:10]))

    # --- バーコードの入力
    def manual_selected():
        sel = [p for p in tv.selection() if p in recs]
        if sel: manual_entry(sel[0])

    def manual_entry(p, on_done=None, pos=None):
        rec = recs.get(p)
        if not rec or not rec['check']:
            status('この写真はトレイを認識できなかったため、入力できません。'); return False
        if rec['excel_row'] is None and state['running']:
            status('この写真はまだExcelに書き込まれていません。処理が終わってから入力してください。'); return False
        win = tk.Toplevel(root); win.transient(root); win.resizable(False, False)
        win.title('バーコードの入力' + (f'（{pos[0]}/{pos[1]}）' if pos else ''))
        body = ttk.Frame(win, padding=12); body.pack(fill='both', expand=True)
        env = from_jpg(rec['env_jpg'])
        if env is None: env = imread(os.path.join(rec['outdir'], rec['check']))
        if env is not None:
            ph = to_photo(env, 720, 470); lab = ttk.Label(body, image=ph); lab.image = ph; lab.pack(anchor='w')
        row = ttk.Frame(body); row.pack(fill='x', pady=(10, 0))
        ttk.Label(row, text=rec['src']).pack(side='left')
        ttk.Label(row, text=f"種子数：{rec['seeds']}").pack(side='right')
        reason = {ST_NEED: 'バーコードを読み取れませんでした。ラベルを見て入力してください。',
                  ST_CHECK: '読み取った値が正しいか、ラベルと見比べてください。',
                  ST_NOMAP: 'このバーコードは対応表にありません。'}.get(rec['status'])
        if reason: ttk.Label(body, text=reason).pack(anchor='w', pady=(4, 0))
        frm = ttk.Frame(body); frm.pack(fill='x', pady=(8, 0))
        ttk.Label(frm, text='バーコードまたは系統名' if MAP.loaded else 'バーコード').grid(row=0, column=0, sticky='w', padx=(0, 10))
        val = tk.StringVar(value=rec['barcode']); e = ttk.Entry(frm, textvariable=val, font=('', 13))
        e.grid(row=0, column=1, sticky='we'); frm.columnconfigure(1, weight=1)
        e.focus_set(); e.select_range(0, 'end')
        result = ttk.Label(frm, text=''); result.grid(row=1, column=1, sticky='w', pady=(4, 0))
        def update_result(*_):
            t = val.get().strip()
            if not MAP.loaded or not t: result.config(text=''); return
            b, n, kind = MAP.resolve(t)
            if kind == 'code': result.config(text=f'系統名：{n}', foreground='')
            elif kind == 'name': result.config(text=f'系統名：{n}（バーコード：{b}）', foreground='')
            else: result.config(text='対応表にありません', foreground=C_ERRTEXT)
        val.trace_add('write', update_result); update_result()

        def close():
            try: win.grab_release()
            except tk.TclError: pass
            win.destroy()

        def ok(_=None):
            t = val.get().strip()
            if MAP.loaded and t and not MAP.resolve(t)[1] and not ask(f'「{t}」は対応表にありません。この値で登録しますか？', win):
                return
            try:
                apply_manual_code(rec, val.get())
            except PermissionError as ex:
                m = str(ex)
                error(m if m == EXCEL_BUSY else f'ファイルを書き換えられません。写真を開いている場合は閉じてから、もう一度実行してください。\n{m}', win)
                return
            except Exception as ex:
                error(str(ex), win); return
            previews[p] = os.path.join(rec['outdir'], rec['check']); refresh_row(p); update_counts()
            if tv.selection() and tv.selection()[0] == p: show(previews[p])
            status(f"{rec['src']}を{rec['name'] or rec['barcode']}として登録しました。")
            close()
            if on_done: on_done()

        bf = ttk.Frame(body); bf.pack(fill='x', pady=(14, 0))
        ttk.Button(bf, text='キャンセル', command=close).pack(side='right')
        if on_done: ttk.Button(bf, text='スキップ', command=lambda: (close(), on_done())).pack(side='right', padx=(0, 6))
        ttk.Button(bf, text='OK', command=ok, default='active').pack(side='right', padx=(0, 6))
        win.bind('<Return>', ok); win.bind('<KP_Enter>', ok); win.bind('<Escape>', lambda _: close())
        win.protocol('WM_DELETE_WINDOW', close)
        def grab():
            try: win.grab_set()
            except tk.TclError: pass
        win.after(80, grab)
        return True

    def manual_all():
        todo = [p for p in files if p in recs and needs_input(recs[p])]
        if not todo: status('要確認の写真はありません。'); return
        it = iter(enumerate(todo, 1))
        def nxt():
            for k, p in it:
                tv.selection_set(p); tv.see(p)
                if manual_entry(p, on_done=nxt, pos=(k, len(todo))): return
            update_counts()
        nxt()

    # --- 千粒重から密度を計算
    def calibrate():
        sel = [p for p in tv.selection() if p in recs and 'vol1000_cm3' in recs[p]['stats']]
        if not sel: info('千粒重を実測した種子の写真を、リストで1枚選んでから実行してください。'); return
        rec = recs[sel[0]]
        t = simpledialog.askfloat('千粒重から密度を計算', f"{rec['src']}の千粒重の実測値（g）", parent=root, minvalue=0.0001)
        if not t: return
        d = t / rec['stats']['vol1000_cm3']
        density.set(f'{d:.3f}'); S['density'] = density.get(); save_settings(S)
        done = [r for r in recs.values() if 'vol1000_cm3' in r['stats'] and r['excel_row']]
        if done and ask(f'密度を{d:.3f}g/cm³にしました。\n処理済みの{len(done)}枚の推定千粒重も、この密度で計算し直しますか？\n'
                        '種類の違う種子が含まれている場合は［いいえ］を選んでください。'):
            if not all(excel_writable(x) for x in {r['xlsx'] for r in done}): error(EXCEL_BUSY); return
            for r in done:
                r['density'] = d; r['stats']['tgw_est'] = r['stats']['vol1000_cm3'] * d
                update_row(r, r['new'], fields=[H_TGW, H_DENS])
            status(f'密度を{d:.3f}g/cm³にし、{len(done)}枚の推定千粒重を更新しました。')
        else:
            status(f'密度を{d:.3f}g/cm³にしました。次に処理する写真から使います。')

    # --- しきい値の調整
    def threshold_window():
        sel = list(tv.selection()) or files[:1]
        if not sel: info('調整に使う写真を追加して、リストで選んでください。'); return
        p = sel[0]
        win = tk.Toplevel(root); win.transient(root); win.title(f'しきい値の調整（{os.path.basename(p)}）'); win.geometry('1100x760')
        P = dict(S['params'])
        left = ttk.Frame(win); left.pack(side='left', fill='both', expand=True)
        img_lab = ttk.Label(left, text='写真を読み込んでいます。', anchor='center'); img_lab.pack(fill='both', expand=True, padx=(8, 0), pady=8)
        right = ttk.Frame(win, padding=(16, 12)); right.pack(side='left', fill='y')
        res = ttk.Label(right, text=''); res.pack(anchor='w')
        busy_lab = ttk.Label(right, text=''); busy_lab.pack(anchor='w', pady=(0, 14))
        specs = [('dark_ratio', '暗さの基準', 0.55, 0.92, '上げると薄い色の粒も数えます。\n上げすぎると影も拾います。'),
                 ('min_radius_mm', '最小の太さ（mm）', 0.20, 1.50, 'これより細いもの（文字、ごみ）は\n数えません。'),
                 ('seg_max_mean', '影の除外', 0.30, 0.80, '下げると薄い影を除きます。\n上げると薄い色の粒が残ります。'),
                 ('rmm', '数える範囲の半径（mm）', 45.0, 61.0, 'シャーレの縁の影を拾うときは\n小さくします。')]
        vars_ = {}
        for key, name, lo, hi, tip in specs:
            box = ttk.Frame(right); box.pack(fill='x', pady=(0, 14))
            head = ttk.Frame(box); head.pack(fill='x')
            ttk.Label(head, text=name).pack(side='left')
            lv = ttk.Label(head, width=6, anchor='e'); lv.pack(side='right')
            v = tk.DoubleVar(value=P[key]); vars_[key] = v
            v.trace_add('write', lambda *_a, v=v, lv=lv: lv.config(text=f'{v.get():.2f}'))
            ttk.Scale(box, from_=lo, to=hi, variable=v, length=300, command=lambda _x: schedule()).pack(fill='x', pady=(2, 3))
            lv.config(text=f'{v.get():.2f}')
            ttk.Label(box, text=tip, justify='left').pack(anchor='w')
        def reset():
            for k, val in core.DEFAULT_PARAMS.items(): vars_[k].set(val)
            schedule()
        def save():
            S['params'] = {k: round(v.get(), 3) for k, v in vars_.items()}; save_settings(S); win.destroy()
            status('しきい値を保存しました。次に処理する写真から使います。処理済みの写真は［数え直す］で反映できます。')
        bf = ttk.Frame(right); bf.pack(side='bottom', fill='x')
        ttk.Button(bf, text='初期値に戻す', command=reset).pack(side='left')
        ttk.Button(bf, text='キャンセル', command=win.destroy).pack(side='right')
        ttk.Button(bf, text='保存', command=save).pack(side='right', padx=(0, 6))
        win.bind('<Escape>', lambda _: win.destroy())
        job = dict(id=None, busy=False, again=False)

        def schedule():
            if job['id']: win.after_cancel(job['id'])
            job['id'] = win.after(350, recount)

        def recount():
            job['id'] = None
            if p not in prepared: return
            if job['busy']: job['again'] = True; return
            job['busy'] = True; busy_lab.config(text='計算しています。')
            params = {k: v.get() for k, v in vars_.items()}; r = prepared[p]
            def run():
                try:
                    n, ann, seeds = core.count_seeds(r['dish'], r['ppm'], params); st = core.seed_stats(seeds)
                    q.put(('cb', lambda: done(n, ann, st)))
                except Exception as ex:
                    q.put(('cb', lambda: done(None, None, str(ex))))
            threading.Thread(target=run, daemon=True).start()

        def done(n, ann, st):
            job['busy'] = False
            if not win.winfo_exists(): return
            if n is None:
                busy_lab.config(text=f'計算できませんでした：{st}', foreground=C_ERRTEXT); return
            ph = to_photo(ann, max(300, left.winfo_width() - 10), max(300, left.winfo_height() - 20))
            img_lab.config(image=ph, text=''); img_lab.image = ph
            res.config(text=f'種子数：{n}' + (f"　平均直径：{st['dia_mean']:.2f}mm" if 'dia_mean' in st else ''))
            busy_lab.config(text='')
            if job['again']: job['again'] = False; recount()

        def prep():
            r = core.prepare(p)
            def put():
                if not r['ok']:
                    if win.winfo_exists(): img_lab.config(text=f"この写真は使えません：{r['msg']}")
                    return
                while len(prepared) >= 3: prepared.pop(next(iter(prepared)))     # メモリ節約: 最近の3枚だけ保持
                prepared[p] = dict(dish=r['dish'], ppm=r['ppm'])
                if win.winfo_exists(): recount()
            q.put(('cb', put))
        if p in prepared: recount()
        else: threading.Thread(target=prep, daemon=True).start()

    # --- 使い方
    def help_window():
        win = tk.Toplevel(root); win.transient(root); win.title('使い方'); win.geometry('600x440')
        t = tk.Text(win, wrap='word', padx=16, pady=12, relief='flat', borderwidth=0, font='TkDefaultFont',
                    background=style.lookup('TFrame', 'background') or None, spacing1=2, spacing3=2)
        t.insert('1.0', (
            '1. ［写真を追加］で、トレイを撮った写真を選びます。HEICも使えます。\n'
            '2. ［開始］を押すと、写真ごとに種子を数え、封筒のバーコードを読み取ります。\n'
            '3. 結果はExcelファイルに1枚1行で追記されます。写真のコピーと確認画像は保存先フォルダに保存されます。\n'
            '4. 黄色の行は確認が必要な写真です。［要確認の写真を入力］で、ラベルを見ながらバーコードを入力します。\n'
            '5. 数え方が合わないときは［しきい値の調整］で設定を変え、［数え直す］で処理済みの写真に反映します。\n\n'
            'キー操作\n'
            f'{MODN}+O　写真を追加\nF5　開始\nEsc　中止\nEnter　選んだ写真のバーコードを入力\n'
            f'{MODN}+R　選んだ写真を数え直す\nDelete　選んだ写真をリストから外す\n'))
        t.config(state='disabled'); t.pack(fill='both', expand=True)
        ttk.Button(win, text='閉じる', command=win.destroy).pack(anchor='e', padx=12, pady=(0, 12))
        win.bind('<Escape>', lambda _: win.destroy())

    update_counts()
    poll()
    root.mainloop()


if __name__ == '__main__':
    if sys.stdout is None:   # 画面アプリとしてビルドした exe では標準出力がない → ログファイルへ
        sys.stdout = sys.stderr = open(os.path.join(os.path.expanduser('~'), 'seed_counter_log.txt'), 'a', encoding='utf-8')
    if len(sys.argv) > 1 and sys.argv[1] == '--cli':
        cli(sys.argv[2:])
    elif len(sys.argv) > 1 and sys.argv[1] == '--selftest':   # ビルド確認用: 必要な部品が全部入っているか
        import scipy.ndimage, skimage.segmentation, skimage.feature, openpyxl, PIL.Image, tkinter
        from pillow_heif import register_heif_opener
        register_heif_opener()                                # HEIC の部品（ネイティブライブラリ）
        assert hasattr(cv2, 'aruco'), 'opencv-contrib (aruco) がありません'
        tcl = tkinter.Tcl().eval('info patchlevel')           # Tcl のライブラリ一式が入っているか
        print('selftest OK', cv2.__version__, 'Tcl', tcl)
        with open('selftest.txt', 'w') as f: f.write(f'selftest OK Tcl {tcl}\n')
    else:
        gui()
