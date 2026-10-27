import json,datetime,collections,itertools,random,sys
exec(open('solve.py').read().split('plan={}')[0])
plan={D.fromisoformat(k):{int(p):tuple(x) for p,x in v.items()} for k,v in json.load(open('plan.json')).items()}
def contrib(d,lab):
    out=collections.Counter()
    for p,x in lab.items():
        for g in G:
            if p in (POS[(d,g)] or []): out[(g,x)]+=1
    return out
def cost(cnt):
    t=0
    for g in G:
        vals=[cnt[(g,s)] for s in SLOTS]
        if sum(vals)==0: continue
        t+=sum(v*v for v in vals)+200*max(0,max(vals)-min(vals)-2)
    return t
def allplans(d):
    w0=d.weekday(); allp=sorted(set(p for g in G for p in (POS[(d,g)] or [])))
    res=[]
    for w2 in [None]+[w for w in range(5) if w!=w0]:
        opts=[cands(p,w0,w2) for p in allp]
        for combo in itertools.product(*opts):
            if len(set(combo))<len(combo): continue
            lab=dict(zip(allp,combo))
            if len({x[0] for x in combo})>2: continue
            res.append(lab)
    return res
segs=[int(a) for a in sys.argv[1:]] or range(6)
for si in segs:
    ds=[d for d in days if seg(d)==si]
    PL={d:allplans(d) for d in ds}
    for rnd in range(8):
        random.seed(rnd); order=ds[:]; random.shuffle(order); ch=0
        for d in order:
            base=collections.Counter()
            for d2 in ds:
                if d2!=d: base.update(contrib(d2,plan[d2]))
            prev=collections.Counter()
            for d2 in days:
                if seg(d2)<si: prev.update(contrib(d2,plan[d2]))
            best=None
            for lab in PL[d]:
                c=base.copy(); c.update(contrib(d,lab))
                cum=prev.copy(); cum.update(c)
                v=cost(c)+0.5*cost(cum)+0.3*sum(1 for p,x in lab.items() if x!=(d.weekday(),p))
                if best is None or v<best[0]: best=(v,lab)
            if best[1]!=plan[d]: plan[d]=best[1]; ch+=1
        if not ch: break
    cnt=collections.Counter()
    for d in ds: cnt.update(contrib(d,plan[d]))
    out=[]
    for g in G:
        vals=[cnt[(g,s)] for s in SLOTS]
        if sum(vals): out.append(f'{g}年 {min(vals)}〜{max(vals)}(差{max(vals)-min(vals)})')
    print(NAMES[si],' / '.join(out))
json.dump({str(d):{str(p):[x[0],x[1]] for p,x in plan[d].items()} for d in days},open('plan.json','w'))
