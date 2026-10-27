import json,datetime,collections,itertools
D=datetime.date
WD='月火水木金'
POSJ=json.load(open('pos.json'))
POS={}; days=set()
for k,v in POSJ.items():
    ds,g=k.split('|'); d=D.fromisoformat(ds); POS[(d,int(g))]=v; days.add(d)
days=sorted(days); G=(1,2,3)
SLOTS=[(w,p) for w in range(5) for p in range(1,7) if not (w==3 and p==6)]
TANNIN={(0,6),(2,6),(4,6)}
ENDS=[D(2027,5,16),D(2027,6,23),D(2027,9,26),D(2027,11,16),D(2027,12,31),D(2028,2,16),D(2028,3,31)]
NAMES=['確認テストまで','確テ〜1期末','1期末〜中間','中間〜2期末','2期末〜12月末','1月〜学年末','学年末後']
def seg(d):
    for i,e in enumerate(ENDS):
        if d<=e: return i
    return len(ENDS)-1
def cands(p,w0,w2):
    c=[]
    for w in {w0,w2}:
        if w is None: continue
        if not(w==3 and p==6): c.append((w,p))
    for t in TANNIN:
        if t[0] in (w0,w2) and t not in c: c.append(t)
    return c
plan={}
C={(s,g):collections.Counter() for s in range(7) for g in G}
PEN=0.6
for d in days:
    w0=d.weekday(); si=seg(d)
    Pg={g:(POS[(d,g)] or []) for g in G}
    allp=sorted(set(p for g in G for p in Pg[g]))
    if not allp: plan[d]={}; continue
    best=None
    for w2 in [None]+[w for w in range(5) if w!=w0]:
        tmp={g:C[(si,g)].copy() for g in G}
        lab={}; used=set(); cost=0
        for p in allp:
            opts=[x for x in cands(p,w0,w2) if x not in used]
            if not opts: opts=[(w0,p)] if (w0,p) not in used else []
            bs=None
            for x in opts:
                inc=sum(2*tmp[g][x]+1 for g in G if p in Pg[g])
                if x!=(w0,p): inc+=PEN
                if bs is None or inc<bs[0]: bs=(inc,x)
            if bs is None: continue
            lab[p]=bs[1]; used.add(bs[1]); cost+=bs[0]
            for g in G:
                if p in Pg[g]: tmp[g][bs[1]]+=1
        ws={x[0] for x in lab.values()}
        if len(ws)>2 or len(lab)<len(allp): continue
        if best is None or cost<best[0]: best=(cost,lab)
    plan[d]=best[1]
    for p,x in best[1].items():
        for g in G:
            if p in Pg[g]: C[(si,g)][x]+=1
# 評価
for si in range(7):
    for g in G:
        c=C[(si,g)]
        vals=[c[s] for s in SLOTS]
        
        print(NAMES[si],f'{g}年','最少',min(vals),'最多',max(vals),'差',max(vals)-min(vals))
ch=sum(1 for d in days for p,x in plan[d].items() if x!=(d.weekday(),p))
print('入れ替えたコマ',ch)
json.dump({str(d):{str(p):[x[0],x[1]] for p,x in plan[d].items()} for d in days},open('plan.json','w'))
