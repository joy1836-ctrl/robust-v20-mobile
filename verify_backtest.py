#!/usr/bin/env python3
"""엔진을 실제 시세로 다시 돌려 사용자 엑셀(Held_State)과 날짜별로 대조한다. 결과: verify/verify_report.json"""
import json, sys
from pathlib import Path
import numpy as np, pandas as pd
import daily_signal as D
E = D.E
HERE = Path(__file__).resolve().parent
ref = json.loads((HERE/"verify/excel_reference.json").read_text(encoding="utf-8"))
cfg, meta, integ = D.load_frozen(HERE/"v20_frozen_config.json")
data = E.load_data(None, "max", False)
f, ex = D.run(data, cfg, ref["cost"]*10000)
idx = [x.strftime("%Y-%m-%d") for x in data["QQQ"].index]
n = len(idx); pos = {d:i for i,d in enumerate(idx)}
eng = np.asarray(ex)[:n]
dates, st = ref["dates"], ref["state"]
res = {}
for off in (-1, 0, 1):          # 엑셀 상태가 '그 날 보유' 인지 '그 날 종가 신호' 인지 정렬을 자동 판별
    mism = []; cnt = 0
    for dte, s in zip(dates, st):
        i = pos.get(dte)
        if i is None or not (0 <= i+off < len(ex)): continue
        cnt += 1
        if int(ex[i+off]) != s: mism.append((dte, s, int(ex[i+off])))
    res[off] = (cnt, mism)
best = min(res, key=lambda o: len(res[o][1]))
cnt, mism = res[best]
missing = [d for d in dates if d not in pos]
# 전환일 비교
def trans(dts, sts): return [(dts[i], sts[i-1], sts[i]) for i in range(1, len(sts)) if sts[i] != sts[i-1]]
et = trans(dates, [int(ex[pos[d]+best]) if d in pos and pos[d]+best < len(ex) else -1 for d in dates])
xt = trans(dates, st)
# 성과 재계산
sr, ex2, stats = E.simulate_strategy(cfg, f, ref["cost"])
sr = np.asarray(sr)[:n]
w = [pos[d] for d in dates if d in pos]
r = sr[w[0]:w[-1]+1]
eq = np.cumprod(1+np.nan_to_num(r)); yrs = (pd.Timestamp(dates[-1])-pd.Timestamp(dates[0])).days/365.25
cagr = eq[-1]**(1/yrs)-1; mdd = float((eq/np.maximum.accumulate(eq)-1).min())
rep = {
 "integrity": integ, "engine_data_range": [idx[0], idx[-1]], "excel_range": [dates[0], dates[-1]],
 "excel_rows": len(dates), "rows_compared": cnt, "rows_missing_in_data": len(missing), "alignment_offset": best,
 "state_mismatch_days": len(mism), "state_match_rate": round(1-len(mism)/max(cnt,1), 6),
 "first_mismatches": mism[:15],
 "excel_transitions": len(xt), "engine_transitions": len(et),
 "transition_lists_identical": et == xt,
 "cagr_engine": round(float(cagr),4), "cagr_excel": round(ref["cagr"],4),
 "mdd_engine": round(mdd,4), "mdd_excel": round(ref["mdd"],4),
}
Path(HERE/"verify/verify_report.json").write_text(json.dumps(rep, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(rep, ensure_ascii=False, indent=2))
