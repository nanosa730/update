#!/usr/bin/env python3
"""前年度の年計（年間計画）シートを、次年度の暦に組み替える。

使い方:
  python shift_nenkei.py config.json

config.json の例は references/config_example.json を参照。
入力ファイルの「年計」シートの並び（月ごとに 日付・行事・会議 の3列、1日2行、4行目から）を前提にする。
"""
import json, re, sys, datetime, math
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl

D = datetime.date
td = datetime.timedelta
WD = '月火水木金土日'
MONTHS = [(4, 'B'), (5, 'E'), (6, 'H'), (7, 'K'), (8, 'N'), (9, 'Q'), (10, 'T'), (11, 'W'), (12, 'Z'), (1, 'AC'), (2, 'AF'), (3, 'AI')]


# ---------- 祝日（内閣府の公表で必ず確認すること） ----------
def nth_monday(y, m, n):
    d = D(y, m, 1)
    d += td(days=(0 - d.weekday()) % 7)
    return d + td(weeks=n - 1)


def holidays(y):
    """y年1月〜12月の祝日 {date: name}"""
    h = {D(y, 1, 1): '元日', nth_monday(y, 1, 2): '成人の日', D(y, 2, 11): '建国記念の日', D(y, 2, 23): '天皇誕生日',
         D(y, 4, 29): '昭和の日', D(y, 5, 3): '憲法記念日', D(y, 5, 4): 'みどりの日', D(y, 5, 5): 'こどもの日',
         nth_monday(y, 7, 3): '海の日', D(y, 8, 11): '山の日', nth_monday(y, 9, 3): '敬老の日',
         nth_monday(y, 10, 2): 'スポーツの日', D(y, 11, 3): '文化の日', D(y, 11, 23): '勤労感謝の日'}
    k = y - 1980
    h[D(y, 3, int(20.8431 + 0.242194 * k - math.floor(k / 4)))] = '春分の日'
    h[D(y, 9, int(23.2488 + 0.242194 * k - math.floor(k / 4)))] = '秋分の日'
    # 国民の休日（祝日にはさまれた平日）
    for d in sorted(h):
        n = d + td(days=2)
        if n in h and (d + td(days=1)) not in h and (d + td(days=1)).weekday() != 6:
            h[d + td(days=1)] = '国民の休日'
    # 振替休日
    for d in sorted(list(h)):
        if d.weekday() == 6:
            n = d + td(days=1)
            while n in h:
                n += td(days=1)
            h[n] = '振替休日'
    return h


HOLNAMES = ['元日', '成人の日', '建国記念の日', '天皇誕生日', '昭和の日', '憲法記念日', 'みどりの日', 'こどもの日', '海の日', '山の日',
            '敬老の日', 'スポーツの日', '文化の日', '勤労感謝の日', '春分の日', '秋分の日', '国民の休日', '振替休日', '日曜日祝日による休業日']
MARK = re.compile(r'^[｜│↓\s　]+$')


def clean(t):
    s = str(t)
    if any(h in s for h in HOLNAMES):
        return ''
    if MARK.match(s):
        return ''
    s = re.sub(r'^[｜│↓\s　]+', '', s)   # 前年に残った矢印は外す（行事は毎日書く方針）
    return s.strip()


def main(cfg):
    sy = cfg['source_year']          # 前年度（例 2027 = 令和9年度）
    ty = sy + 1
    shift = (D(ty, 4, 1) - D(sy, 4, 1)).days
    shift = shift - shift % 7 if shift % 7 else shift   # 同じ曜日になる日数（364 or 371）
    if shift not in (364, 371):
        shift = 364
    P = lambda s: D.fromisoformat(s)
    src = openpyxl.load_workbook(cfg['source'])
    ws = src[cfg.get('sheet', '年計')]

    # ---- 読み込み
    R0, COL = {}, {}
    for m, L in MONTHS:
        y = sy + 1 if m <= 3 else sy
        dc, ec, cc = L, gl(ci(L) + 1), gl(ci(L) + 2)
        for day in range(1, 32):
            try:
                d = D(y, m, day)
            except ValueError:
                continue
            r1 = 4 + 2 * (day - 1)
            ev, co = [], []
            for r in (r1, r1 + 1):
                v = ws[f'{ec}{r}'].value
                if v not in (None, ''):
                    for line in str(v).replace('\n', '・').split('・'):
                        t = clean(line)
                        if t:
                            ev.append(t)
                            col = ws[f'{ec}{r}'].font.color
                            COL[t] = col.rgb if (col is not None and isinstance(col.rgb, str)) else None
                v = ws[f'{cc}{r}'].value
                if v not in (None, ''):
                    co += [s for s in str(v).split('\n') if s and not re.match(r'^研\d+$', s)]
            R0[d] = dict(ev=ev, co=co)

    HOL = {**holidays(ty), **holidays(ty + 1)}
    # 休業日は「名古屋市立小学校及び中学校の管理運営に関する規則」で決まっている
    #   学年始 4/1〜4/5、夏季 7/21〜8/31、冬季 12/24〜1/6、学年末 3/25〜3/31
    default_vac = [(D(ty, 4, 1), D(ty, 4, 5)), (D(ty, 7, 21), D(ty, 8, 31)),
                   (D(ty, 12, 24), D(ty + 1, 1, 6)), (D(ty + 1, 3, 25), D(ty + 1, 3, 31))]
    vac = [(P(a), P(b)) for a, b in cfg['vacations']] if cfg.get('vacations') else default_vac
    if cfg.get('first_day'):
        first_day = P(cfg['first_day'])
    else:
        first_day = D(ty, 4, 6)
        while first_day.weekday() >= 5 or first_day in HOL:
            first_day += td(days=1)

    def school(d):
        return d.weekday() < 5 and d not in HOL and not any(a <= d <= b for a, b in vac) and d >= first_day

    # ---- 日付の対応（同じ週・同じ曜日。anchors は日付どおり）
    anchor = {P(a): P(b) for a, b in cfg.get('anchors', {}).items()}
    adst = {v: k for k, v in anchor.items()}
    R, flags = {}, []
    d = D(ty, 4, 1)
    while d <= D(ty + 1, 3, 31):
        s = adst.get(d)
        if s is None:
            s = d - td(days=shift)
            if s in anchor:
                s = None
        x = R0.get(s) if s else None
        R[d] = dict(ev=list(x['ev']) if x else [], co=list(x['co']) if x else [])
        d += td(days=1)
    for s in anchor:   # アンカーで押し出された日の行事は捨てずに一覧へ
        pass
    # 年度初め（始業式より前）の会議は、曜日ではなく順番で次年度の4/1以降の平日へ
    sfd = P(cfg['source_first_day']) if cfg.get('source_first_day') else None
    if sfd:
        pre = [x for x in sorted(R0) if x < sfd and x.month == 4 and x.weekday() < 5]
        tgt = [x for x in sorted(R) if x < first_day and x.weekday() < 5 and x not in HOL]
        for k in tgt:
            R[k] = dict(ev=[], co=[])
        for s_, t_ in zip(pre, tgt):
            R[t_] = dict(ev=list(R0[s_]['ev']), co=list(R0[s_]['co']))
        for s_ in pre[len(tgt):]:
            R[tgt[-1]]['ev'] += R0[s_]['ev']; R[tgt[-1]]['co'] += R0[s_]['co']
    for t in cfg.get('remove_texts', []):
        for k in R:
            R[k]['ev'] = [x for x in R[k]['ev'] if t not in x]
    for k, t in cfg.get('fixed_events', {}).items():
        R[P(k)]['ev'].insert(0, t)

    # ---- 祝日に当たった行事は次の登校日へ（赤字※）
    for d in sorted(R):
        if d in HOL:
            if R[d]['ev'] or R[d]['co']:
                n = d + td(days=1)
                while not school(n):
                    n += td(days=1)
                if R[d]['ev']:
                    R[n]['ev'] = ['※' + x for x in R[d]['ev']] + R[n]['ev']
                    flags.append((d, HOL[d], '・'.join(R[d]['ev']), f'{n.month}/{n.day}へ移動'))
                R[n]['co'] += R[d]['co']
            R[d]['ev'] = [HOL[d]]
            R[d]['co'] = []

    # ---- 会議は同じ週の木曜へ（B帯：木曜は5限後に研修→会議）
    MOVE = re.compile(cfg.get('move_codes_regex', r'^(職\d+|学\d+.*|運\d+|旧運|現\d+|努･採|不|予|教|教[･・]部活|教支|採|編\d|卒認|進\d|進認|部反.*|全反|部会|教科等研究部会)$'))
    for d in sorted(R):
        if not school(d) or d.weekday() == 3:
            continue
        if d < first_day:
            continue
        keep = []
        for c in R[d]['co']:
            th = d + td(days=3 - d.weekday())
            if MOVE.match(c) and school(th):
                R[th]['co'].append(c)
                continue
            keep.append(c)
        R[d]['co'] = keep

    # ---- 研修（木曜）に番号。研修できない木曜はキーワードで除く
    nk = cfg.get('kenshu_count', 35)
    # 研修ができない木曜：全校で先生が出られない行事の日（教育相談・保護者会の期間、修学旅行、体育大会当日）
    def no_kenshu(items):
        for it in items:
            it = it.lstrip('※')
            if '教育相談' in it or it in ('保護者会（全）', '保護者会(全)') or it.startswith('修学旅行(') or it == '体育大会':
                return True
        return False
    ths = [d for d in sorted(R) if d.weekday() == 3 and school(d)]
    bad = [d for d in ths if no_kenshu(R[d]['ev'])]
    ok = [d for d in ths if d not in bad]
    for i, d in enumerate(ok[:nk], 1):
        R[d]['co'].insert(0, f'研{i}')

    # ---- 書き出し（元のシートの書式を残す）
    for n in list(src.sheetnames):
        if n != ws.title:
            del src[n]
    for v in src.views:
        v.firstSheet = 0
        v.activeTab = 0
    for nm in list(src.defined_names.keys()):
        del src.defined_names[nm]
    for nm in list(ws.defined_names.keys()):
        if nm != '_xlnm.Print_Area':
            del ws.defined_names[nm]
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 4 and mr.max_row <= 65:
            ws.unmerge_cells(str(mr))
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and ('WEBSERVICE' in c.value or c.value.startswith('ERROR')):
                c.value = None
    green = PatternFill('solid', fgColor='92D050')
    nofill = PatternFill(fill_type=None)
    ws['C1'] = cfg.get('title', f'令和{ty - 2018}年度')
    ws['C2'] = f'（{ty}年）'
    ws['AD2'] = f'（{ty + 1}年）'
    for m, L in MONTHS:
        y = ty + 1 if m <= 3 else ty
        dc, ec, cc = L, gl(ci(L) + 1), gl(ci(L) + 2)
        for day in range(1, 32):
            r1 = 4 + 2 * (day - 1)
            for c in (dc, ec, cc):
                for r in (r1, r1 + 1):
                    ws[f'{c}{r}'].value = None
                    ws[f'{c}{r}'].fill = nofill
            try:
                d = D(y, m, day)
            except ValueError:
                continue
            ws[f'{dc}{r1}'] = day
            ws[f'{dc}{r1 + 1}'] = WD[d.weekday()]
            evs = list(dict.fromkeys(R[d]['ev']))
            if len(evs) > 2:
                evs = [evs[0], '・'.join(evs[1:])]
            for i, t in enumerate(evs):
                c = ws[f'{ec}{r1 + i}']
                c.value = t
                rgb = 'FFFF0000' if t.startswith('※') else COL.get(t.split('・')[0])
                c.font = Font(name=c.font.name, size=c.font.sz, color=rgb)
            cos = list(dict.fromkeys(R[d]['co']))
            if cos:
                h = (len(cos) + 1) // 2
                ws[f'{cc}{r1}'] = '\n'.join(cos[:h])
                ws[f'{cc}{r1 + 1}'] = '\n'.join(cos[h:]) or None
            if d.weekday() >= 5 or d in HOL:   # 緑は土日・祝日だけ
                for c in (dc, ec, cc):
                    for r in (r1, r1 + 1):
                        ws[f'{c}{r}'].fill = green

    memo = src.create_sheet('作成メモ・要確認')
    memo.append([f'{ws["C1"].value} 年間計画（案）  自動組み替えの記録'])
    memo.append([f'・同じ週・同じ曜日に移した（{shift}日後）。日付で決まるもの（anchors）は日付どおり。'])
    memo.append(['・祝日は計算で入れた。内閣府の公表で必ず確認すること。'])
    memo.append([f'・研修は木曜に{min(nk, len(ok))}回（木曜の登校日{len(ths)}日、研修できない木曜{len(bad)}日：' + '、'.join(f'{d.month}/{d.day}' for d in bad) + '）'])
    if len(ok) < nk:
        memo.append([f'※ 研修が{nk}回に足りない（{len(ok)}回）。行事の曜日を見直すこと。'])
    memo.append([])
    memo.append(['祝日に当たって移した行事', '', '', ''])
    for f in flags:
        memo.append([f'{f[0].month}/{f[0].day}', f[1], f[2], f[3]])
    memo.column_dimensions['A'].width = 16
    memo.column_dimensions['C'].width = 40
    src.save(cfg['output'])
    print('saved', cfg['output'], 'flags', len(flags), 'kenshu', min(nk, len(ok)), 'bad_thursdays', [str(x) for x in bad])


if __name__ == '__main__':
    main(json.load(open(sys.argv[1], encoding='utf-8')))
