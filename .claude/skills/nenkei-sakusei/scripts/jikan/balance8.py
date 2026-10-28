# 年間の各コマ回数の差を3以内（できれば2）に、区切りごとの差を2以内にそろえる局所探索
# 条件：1限には1限のコマか月6・水6・金6だけ／1日に入れる曜日は2つまで／同じ日に同じコマは1回
import json,datetime,collections,random
D=datetime.date
P=json.load(open('pos.json')); L=json.load(open('plan.json'))
G=(1,2,3)
SLOTS=[(w,p) for w in range(5) for p in range(1,7) if not (w==3 and p==6)]
TANNIN={(0,6),(2,6),(4,6)}
ENDS=[D(2027,5,16),D(2027,6,23),D(2027,9,26),D(2027,11,16),D(2027,12,31),D(2028,2,16),D(2028,3,31)]
def seg(d): return next(i for i,e in enumerate(ENDS) if d<=e)
days=sorted({D.fromisoformat(k.split('|')[0]) for k in P})
POS={(D.fromisoformat(k.split('|')[0]),int(k.split('|')[1])):(v or []) for k,v in P.items()}
plan={d:{int(p):tuple(x) for p,x in L[str(d)].items()} for d in days}
SEG=collections.Counter(); YR=collections.Counter()
for d in days:
    for p,x in plan[d].items():
        for g in G:
            if p in POS[(d,g)]: SEG[(seg(d),g,x)]+=1; YR[(g,x)]+=1
def ok(d,lab):
    if len(set(lab.values()))<len(lab): return False
    if len({x[0] for x in lab.values()})>2: return False
    if 1 in lab and not(lab[1][1]==1 or lab[1] in TANNIN): return False
    return d.weekday() in {x[0] for x in lab.values()} or not lab
def rng(vals): return max(vals)-min(vals)
def cost_seg(si,g):
    v=[SEG[(si,g,s)] for s in SLOTS]
    if sum(v)==0: return 0
    return sum(a*a for a in v)+1000*max(0,rng(v)-2)
def cost_yr(g):
    v=[YR[(g,s)] for s in SLOTS]
    return 3*sum(a*a for a in v)+300*max(0,rng(v)-2)+5000*max(0,rng(v)-3)
def movec(d,lab): return sum(0.4 for p,x in lab.items() if x!=(d.weekday(),p))
def apply(d,lab,sign):
    si=seg(d)
    for p,x in lab.items():
        for g in G:
            if p in POS[(d,g)]: SEG[(si,g,x)]+=sign; YR[(g,x)]+=sign
def local(d,lab): return sum(cost_seg(seg(d),g) for g in G)+sum(cost_yr(g) for g in G)+movec(d,lab)
random.seed(1)
for it in range(40):
    ch=0; order=days[:]; random.shuffle(order)
    for d in order:
        lab=plan[d]
        if not lab: continue
        base=local(d,lab)
        best=(base,lab)
        ps=sorted(lab)
        cand=[]
        # 1か所を別のコマに替える／2か所を入れ替える
        for p in ps:
            for s in SLOTS:
                if s==lab[p]: continue
                n=dict(lab); n[p]=s; cand.append(n)
        for i in range(len(ps)):
            for j in range(i+1,len(ps)):
                n=dict(lab); n[ps[i]],n[ps[j]]=lab[ps[j]],lab[ps[i]]; cand.append(n)
        for n in cand:
            if not ok(d,n): continue
            apply(d,lab,-1); apply(d,n,1)
            c=local(d,n)
            apply(d,n,-1); apply(d,lab,1)
            if c<best[0]-1e-9: best=(c,n)
        if best[1] is not lab:
            apply(d,lab,-1); apply(d,best[1],1); plan[d]=best[1]; ch+=1
    if not ch: break
NAMES=['確認テストまで','確テ〜1期末','1期末〜中間','中間〜2期末','2期末〜12月末','1月〜学年末','学年末後']
for si in range(7):
    print(NAMES[si],' / '.join(f'{g}年 {min(v)}〜{max(v)}' for g in G for v in [[SEG[(si,g,s)] for s in SLOTS]] if sum(v)))
for g in G:
    v=[YR[(g,s)] for s in SLOTS]; print(f'{g}年 年間',min(v),'〜',max(v))
print('時間割変更コマ',sum(1 for d in days for p,x in plan[d].items() if x!=(d.weekday(),p)))
json.dump({str(d):{str(p):[x[0],x[1]] for p,x in plan[d].items()} for d in days},open('plan.json','w'))
