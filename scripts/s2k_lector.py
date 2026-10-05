import re,pandas as pd
def load(path="v10/2026-09-30_CASA_MR_REV10cambios.s2k"):
    lines=open(path,encoding="latin-1").read().splitlines()
    T={};cur=None
    for l in lines:
        m=re.match(r'\s*TABLE:\s+"(.*)"',l)
        if m: cur=m.group(1); T[cur]=[]; continue
        if not l.strip(): cur=None; continue
        if cur: T[cur].append(dict(re.findall(r'(\w+)=(\S+)',l)))
    return {k:pd.DataFrame(v) for k,v in T.items()}
