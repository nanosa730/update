import openpyxl,json,datetime,re,copy
from openpyxl.styles import PatternFill,Font
from openpyxl.comments import Comment
D=datetime.date; td=datetime.timedelta
WD='月火水木金土日'; Z='０１２３４５６'
P=json.load(open('pos.json')); PF=json.load(open('pos_full.json')); L=json.load(open('plan.json')); SW=json.load(open('swap2.json'))
wb=openpyxl.load_workbook('/home/user/update/参照資料/Excel/R8_授業数カウント_9.1版.xlsx')
y=wb['年計']
for mr in list(y.merged_cells.ranges):
    if mr.min_col<=21 and mr.min_row>=3: y.unmerge_cells(str(mr))
mst=openpyxl.load_workbook('/home/user/update/令和9年度_年間計画（案）.xlsx')['年計']
from openpyxl.utils import get_column_letter as gl, column_index_from_string as ci
MC={4:'B',5:'E',6:'H',7:'K',8:'N',9:'Q',10:'T',11:'W',12:'Z',1:'AC',2:'AF',3:'AI'}
def evs(d):
    ec=gl(ci(MC[d.month])+1); r=4+2*(d.day-1)
    return '・'.join(str(mst[f'{ec}{rr}'].value) for rr in (r,r+1) if mst[f'{ec}{rr}'].value)
HOL={D(2027,4,29),D(2027,5,3),D(2027,5,4),D(2027,5,5),D(2027,7,19),D(2027,8,11),D(2027,9,20),D(2027,9,23),D(2027,10,11),D(2027,11,3),D(2027,11,23),D(2028,1,1),D(2028,1,10),D(2028,2,11),D(2028,2,23),D(2028,3,20)}
chg=PatternFill('solid',fgColor='FFC000')   # 時間割変更（その曜日と違うコマ）
blue=PatternFill('solid',fgColor='9BC2E6')  # 2年 総合→道徳・学活
none=PatternFill(fill_type=None)
row=3; d=D(2027,4,1); rows={}
while d<=D(2028,3,31):
    rows[d]=row
    y.cell(row,1,(d-D(1899,12,30)).days); y.cell(row,2,WD[d.weekday()])
    if row>6:   # 3〜6行のC・D列は凡例なので残す
        y.cell(row,3,evs(d) or None)
    for c in range(4,22):
        cell=y.cell(row,c); cell.value=None; cell.fill=none
    ds=str(d)
    for gi,g in enumerate((1,2,3)):
        base=4+gi*6
        pos=P.get(f'{ds}|{g}')
        full=PF.get(f'{ds}|{g}') or []
        if pos is None and not full: continue
        if d<D(2027,4,14):
            for p in full: y.cell(row,base+p-1,'学活')
            continue
        for p in (pos or []):
            x=L[ds][str(p)]; c=y.cell(row,base+p-1,WD[x[0]]+Z[x[1]])
            if not(x[0]==d.weekday() and x[1]==p): c.fill=chg
            if g==2:
                for s in SW:
                    if s['date']==ds and s['pos']==p:
                        c.fill=blue; c.comment=Comment(f"2年は総合のかわりに{s['to']}",'年計作成')
    row+=1; d+=td(days=1)
last=row-1
for r in range(last+1,y.max_row+1):
    for c in range(1,22): y.cell(r,c).value=None
def R(dt): return rows[dt]
# 合計シートの区切り（令和9年度）
seg=[(3,R(D(2027,5,16))),(R(D(2027,5,17)),R(D(2027,6,23))),(R(D(2027,6,24)),R(D(2027,9,26))),(R(D(2027,9,27)),R(D(2027,11,16))),(R(D(2027,11,17)),R(D(2027,12,31))),(R(D(2028,1,1)),R(D(2028,2,16))),(R(D(2028,2,17)),last)]
blocks=[(4,11),(12,19),(20,27),(28,35),(36,43),(44,51),(52,59)]
s=wb['合計']
pat=re.compile(r'年計!\$([A-Z]+)\$\d+:\$([A-Z]+)\$\d+')
for (b0,b1),(a,b) in zip(blocks,seg):
    for r in range(b0,b1+1):
        for c in s[r]:
            v=c.value
            if not(isinstance(v,str) and v.startswith('=')): continue
            col=c.column
            if 'MATCH' in v or '$C$' in v:
                c.value=re.sub(r'\$C\$2:\$C\$\d+',f'$C$2:$C${last}',v); continue
            if col<=16: c.value=pat.sub(lambda m:f'年計!${m.group(1)}${a}:${m.group(2)}${b}',v)
            else: c.value=pat.sub(lambda m:f'年計!${m.group(1)}$3:${m.group(2)}${b}',v)
s['A4']='＜確認テストまで（授業開始4/14）＞'
s['A44']='＜1月初め～１２年学年末テ＞'
for v in wb.views: v.firstSheet=0; v.activeTab=0
y['W2']='橙＝時間割変更したコマ　水色＝2年の総合を道徳・学活に替えるコマ　「学活」＝授業開始(4/14)前の学級の時間'
y['W2'].font=Font(size=10,bold=True)
out='/home/user/update/令和9年度_授業数カウント.xlsx'
wb.save(out); print(out,last,seg)
