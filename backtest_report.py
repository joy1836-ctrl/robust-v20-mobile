"""백테스트 보고서 데이터(D)를 최신 시세로 다시 만든다. docs/backtest.html 의 `const D = {...};` 줄을 갱신하는 데 쓴다.
기간: 마지막 거래일에서 거꾸로 15년(252거래일×15). V20 = 종가 기준 지수(daily_signal.v20_nav), 비교 = TQQQ 매수 후 보유."""
import json, re, sys
import numpy as np
import pandas as pd
import daily_signal as D

NAMES = ["CASH", "QQQ", "QLD", "TQQQ"]
YEARS = 15
H = 252


def _stats(eq, ret):
    n = len(eq) - 1
    cagr = eq[-1] ** (H / n) - 1.0
    mdd = float((eq / np.maximum.accumulate(eq) - 1.0).min())
    vol = float(np.std(ret, ddof=1) * np.sqrt(H))
    sharpe = float(np.mean(ret) / np.std(ret, ddof=1) * np.sqrt(H))
    return {"Total Return": float(eq[-1] - 1.0), "CAGR": float(cagr), "MDD": mdd, "Annualized Vol": vol, "Sharpe": sharpe,
            "Calmar": float(cagr / abs(mdd)) if mdd else 0.0}


TAX_RATE, TAX_DED = 0.22, 2.5e6
FX_CSV = "fx_usdkrw_daily.csv"


def load_fx(idx, extra=None, base_path=None):
    """원/달러 환율(거래일 정렬, 앞 값으로 채움). 저장소 CSV + 오늘 받은 최근 환율(extra={'dates','series'})."""
    import pandas as pd
    from pathlib import Path
    path = Path(base_path) if base_path else Path(__file__).with_name(FX_CSV)
    s = pd.read_csv(path, parse_dates=["date"]).set_index("date")["usd_krw"].astype(float)
    if extra and extra.get("dates"):
        e = pd.Series(extra["series"], index=pd.to_datetime(extra["dates"]), dtype=float)
        s = pd.concat([s[s.index < e.index[0]], e])
    s = s[~s.index.duplicated(keep="last")].sort_index()
    out = s.reindex(s.index.union(idx)).ffill().reindex(idx)
    return out.values.astype(float)


def krw_tax_years(idx, ex, O, C, fx, s0, cost, cap=1e8, ded=None):
    """V20을 원화로 환산하고 해외주식 양도세(22%, 연 250만 공제, 연말 납부)를 반영한 연도별 결과.
    매도 시 취득원가 대비 이익을 실현으로 보고(갈아타기 매수는 종가, 현금 매수는 시가). 납부는 다음 해 4월 30일 일부 매도, 아직 안 낸 세금은 마지막 날에 정산(예상)."""
    n = len(idx)
    ded = TAX_DED if ded is None else ded
    if not np.all(np.isfinite(fx[s0:n])):
        return None

    # 날짜: 매년 4월 30일(휴장이면 직전 거래일)에 전년 귀속 세금만큼 일부 매도해 납부(TAX_RESERVE_V2).
    apr = {}
    for i in range(s0, n):
        if idx[i] <= pd.Timestamp(year=idx[i].year, month=4, day=30):
            apr[idx[i].year] = i
    apr_days = set(apr.values())

    def sim(tax):
        eu = cap / fx[s0]; basis = cap; gains = 0.0; yr = idx[s0].year; ye = {}; taxes = {}; due = 0.0
        for i in range(s0 + 1, n):
            if idx[i].year != yr:
                t = TAX_RATE * max(0.0, gains - ded) if tax else 0.0
                due = t; ye[yr] = eu * fx[i - 1]; taxes[yr] = t; gains = 0.0; yr = idx[i].year
            p, q = int(ex[i - 1]), int(ex[i])
            if p == q:
                if p: eu *= C[p][i] / C[p][i - 1]
            else:
                if p:
                    eu *= O[p][i] / C[p][i - 1]; gains += eu * fx[i] - basis; eu *= (1 - cost)
                if q:
                    if p:
                        eu *= (1 - cost); basis = eu * fx[i]
                    else:
                        eu *= (1 - cost) * C[q][i] / O[q][i]; basis = eu / (C[q][i] / O[q][i]) * fx[i]
            if tax and i in apr_days and due > 0:      # 4/30: 전년 세금만큼 보유분 비례 매도(이익이 있으면 올해 실현이익에 포함)
                val = eu * fx[i]
                f = min(1.0, due / val) if val > 0 else 0.0
                if int(ex[i]) != 0: gains += f * (val - basis)
                eu *= (1 - f); basis *= (1 - f); due = 0.0
        t = TAX_RATE * max(0.0, gains - ded) if tax else 0.0
        if tax: eu -= (t + due) / fx[n - 1]              # 아직 안 낸 세금은 지금 정산한 것으로 가정(예상)
        ye[yr] = eu * fx[n - 1]; taxes[yr] = t
        return ye, taxes

    a, _ = sim(False); b, tb = sim(True)
    rows = {}; pa = pb = cap
    for y in sorted(a):
        rows[y] = {"kr": a[y] / pa - 1.0, "kt": b[y] / pb - 1.0, "tax": float(tb[y]), "kv": float(a[y]), "ktv": float(b[y])}
        pa, pb = a[y], b[y]
    yrs = (idx[n - 1] - idx[s0]).days / 365.25
    last = sorted(a)[-1]
    summ = {"cap": cap, "cagr_pre": float((a[last] / cap) ** (1 / yrs) - 1), "cagr_post": float((b[last] / cap) ** (1 / yrs) - 1),
            "final_pre": float(a[last]), "final_post": float(b[last]), "total_tax": float(sum(tb.values()))}
    return rows, summ


def usd_tax_nav(idx, ex, O, C, s0, cost, cap_usd, ded_usd, rate=None):
    """달러 기준 세후 가치(일별). 양도세 22%, 연 공제, 다음 해 4월 30일 보유분 일부 매도로 납부(TAX_RESERVE_V2).
    세후 가치 = 보유 평가액 - (이미 실현한 이익에 대한 미납 세금). 평가이익(미실현)에는 세금을 빼지 않는다.
    반환: (net, year_end_base) net[i] (i>=s0) = 그날 종가 기준 세후 가치(달러)."""
    rate = TAX_RATE if rate is None else rate
    n = len(idx)
    apr = {}
    for i in range(s0, n):
        if idx[i] <= pd.Timestamp(year=idx[i].year, month=4, day=30):
            apr[idx[i].year] = i
    apr_days = set(apr.values())
    eu = cap_usd; basis = cap_usd; gains = 0.0; yr = idx[s0].year; due = 0.0
    net = np.full(n, np.nan)
    net[s0] = eu
    for i in range(s0 + 1, n):
        if idx[i].year != yr:
            due += rate * max(0.0, gains - ded_usd); gains = 0.0; yr = idx[i].year
        p, q = int(ex[i - 1]), int(ex[i])
        if p == q:
            if p: eu *= C[p][i] / C[p][i - 1]
        else:
            if p:
                eu *= O[p][i] / C[p][i - 1]; gains += eu - basis; eu *= (1 - cost)
            if q:
                if p:
                    eu *= (1 - cost); basis = eu
                else:
                    eu *= (1 - cost) * C[q][i] / O[q][i]; basis = eu / (C[q][i] / O[q][i])
        if i in apr_days and due > 0:
            f = min(1.0, due / eu) if eu > 0 else 0.0
            if int(ex[i]) != 0: gains += f * (eu - basis)
            eu *= (1 - f); basis *= (1 - f); due = 0.0
        net[i] = eu - due - rate * max(0.0, gains - ded_usd)
    return net


def krw_tax_nav(idx, ex, O, C, fx, s0, cost, cap_krw, ded_krw=None, rate=None):
    """원화 세후 가치(일별). usd_tax_nav 와 같은 규칙인데 환율을 반영한다.
    취득원가·이익은 거래일 환율의 원화로 계산하고, 평가액은 그날 환율로 원화 환산한다."""
    rate = TAX_RATE if rate is None else rate
    ded = TAX_DED if ded_krw is None else ded_krw
    n = len(idx)
    apr = {}
    for i in range(s0, n):
        if idx[i] <= pd.Timestamp(year=idx[i].year, month=4, day=30):
            apr[idx[i].year] = i
    apr_days = set(apr.values())
    eu = cap_krw / fx[s0]; basis = cap_krw; gains = 0.0; yr = idx[s0].year; due = 0.0
    net = np.full(n, np.nan)
    net[s0] = cap_krw
    for i in range(s0 + 1, n):
        if idx[i].year != yr:
            due += rate * max(0.0, gains - ded); gains = 0.0; yr = idx[i].year
        p, q = int(ex[i - 1]), int(ex[i])
        if p == q:
            if p: eu *= C[p][i] / C[p][i - 1]
        else:
            if p:
                eu *= O[p][i] / C[p][i - 1]; gains += eu * fx[i] - basis; eu *= (1 - cost)
            if q:
                if p:
                    eu *= (1 - cost); basis = eu * fx[i]
                else:
                    eu *= (1 - cost) * C[q][i] / O[q][i]; basis = eu / (C[q][i] / O[q][i]) * fx[i]
        val = eu * fx[i]
        if i in apr_days and due > 0:
            f = min(1.0, due / val) if val > 0 else 0.0
            if int(ex[i]) != 0: gains += f * (val - basis)
            eu *= (1 - f); basis *= (1 - f); due = 0.0; val = eu * fx[i]
        net[i] = val - due - rate * max(0.0, gains - ded)
    return net


def build_report(idx, ex_full, v_nav, t_close, cost, label_extra="", krw=None):
    """idx: 날짜(DatetimeIndex), ex: 시가 상태, v_nav: V20 가치(시계열, 길이 n), t_close: TQQQ 가격."""
    n = len(idx)
    L = YEARS * H
    s = max(0, n - 1 - L)
    d = idx[s:n]
    ex = np.asarray(ex_full)[:n][s:]
    r_eq = np.asarray(v_nav, float)[s:] / v_nav[s]
    t_eq = np.asarray(t_close, float)[s:] / t_close[s]
    m = len(r_eq) - 1
    rr, tr = np.diff(r_eq) / r_eq[:-1], np.diff(t_eq) / t_eq[:-1]
    sr, st = _stats(r_eq, rr), _stats(t_eq, tr)
    full = {k: {"r": sr[k], "t": st[k]} for k in sr}
    # 전환
    trans = [{"date": d[i].strftime("%Y-%m-%d"), "from": NAMES[int(ex[i - 1])], "to": NAMES[int(ex[i])]} for i in range(1, len(ex)) if ex[i] != ex[i - 1]]
    sw = sum(1 for t in trans if t["from"] != "CASH" and t["to"] != "CASH")
    to_c = sum(1 for t in trans if t["to"] == "CASH"); fr_c = sum(1 for t in trans if t["from"] == "CASH")
    yrs = (d[-1] - d[0]).days / 365.25
    activity = {"buys": sw + fr_c, "sells": sw + to_c, "switches": sw, "changes": len(trans), "per_year": sw / yrs, "cash_moves": f"{to_c} / {fr_c}"}
    cnt = [float(np.mean(ex[:-1] == k)) for k in range(4)]
    exposure = {"cash": cnt[0], "qqq": cnt[1], "qld": cnt[2], "tqqq": cnt[3]}
    # 달력 연도
    cal = []
    years = sorted({x.year for x in d})
    for y in years:
        ii = [i for i, x in enumerate(d) if x.year == y]
        a = max(ii[0] - 1, 0)
        b = ii[-1]
        sr_, tr_ = r_eq[a:b + 1], t_eq[a:b + 1]
        cal.append({"y": y, "days": len(ii), "r": float(sr_[-1] / sr_[0] - 1), "t": float(tr_[-1] / tr_[0] - 1),
                    "rm": float((sr_ / np.maximum.accumulate(sr_) - 1).min()), "tm": float((tr_ / np.maximum.accumulate(tr_) - 1).min())})
    # 롤링
    roll, series = {}, {}
    for k in range(1, YEARS + 1):
        w = k * H
        nw = len(r_eq) - w
        if nw <= 0: continue
        rc, tc, rmd, tmd = [], [], [], []
        for i in range(nw):
            a, b = r_eq[i:i + w + 1], t_eq[i:i + w + 1]
            rc.append((a[-1] / a[0]) ** (H / w) - 1); tc.append((b[-1] / b[0]) ** (H / w) - 1)
            rmd.append((a / np.maximum.accumulate(a) - 1).min()); tmd.append((b / np.maximum.accumulate(b) - 1).min())
        rc, tc, rmd, tmd = map(np.array, (rc, tc, rmd, tmd))
        rw, mw = rc > tc, rmd > tmd
        roll[f"{k}Y"] = {"years": k, "windows": int(nw), "r_cagr": float(rc.mean()), "t_cagr": float(tc.mean()),
                         "r_cagr_worst": float(rc.min()), "t_cagr_worst": float(tc.min()), "r_cagr_best": float(rc.max()),
                         "r_mdd": float(rmd.mean()), "t_mdd": float(tmd.mean()), "r_mdd_worst": float(rmd.min()), "t_mdd_worst": float(tmd.min()),
                         "ret_win": float(rw.mean()), "mdd_win": float(mw.mean()), "both_win": float((rw & mw).mean())}
        if k in (1, 3, 5, 10):
            series[f"{k}Y"] = {"d": [x.strftime("%Y-%m-%d") for x in d[:nw]], "r": [round(float(x), 4) for x in rc], "t": [round(float(x), 4) for x in tc]}
    krw_sum = None
    if krw:
        res = krw_tax_years(idx, np.asarray(ex_full)[:n], krw["O"], krw["C"], krw["fx"], s, krw["cost"])
        if res:
            rows, krw_sum = res
            for r in cal:
                x = rows.get(r["y"])
                if x: r.update({k: float(v) for k, v in x.items()})
            krw_sum["fx_asof"] = idx[n - 1].strftime("%Y-%m-%d")
        # 달러 기준 세후: 환율을 1로 두고, 시작 환율로 원금(1억 원)과 공제(250만 원)를 달러로 바꿔 쓴다.
        f0 = float(krw["fx"][s])
        ru = krw_tax_years(idx, np.asarray(ex_full)[:n], krw["O"], krw["C"], np.ones(n), s, krw["cost"], cap=1e8 / f0, ded=TAX_DED / f0)
        if ru:
            rows_u, sum_u = ru
            for r in cal:
                x = rows_u.get(r["y"])
                if x: r["ut"] = float(x["kt"])
            sum_u = {k: sum_u[k] for k in ("cagr_pre", "cagr_post")}
            krw_sum = dict(krw_sum or {}, usd=sum_u)
    end = d[-1].strftime("%Y-%m-%d")
    return {"asof": f"ROBUST V20 vs TQQQ | Fixed 15Y Backtest as of {end}{label_extra}", "start": d[0].strftime("%Y-%m-%d"), "end": end, "cost": cost,
            "full": full, "activity": activity, "exposure": exposure, "cal": cal, "roll": roll, "series": series, "trans": trans,
            "krw": krw_sum, "eq": {"d": [x.strftime("%Y-%m-%d") for x in d], "s": [round(float(x), 3) for x in r_eq], "t": [round(float(x), 4) for x in t_eq], "st": [int(x) for x in ex]}}


def make(data, cfg, cost_bps=8.0, fx_hist=None):
    f, ex, sr = D.run_full(data, cfg, cost_bps)
    idx = data["QQQ"].index
    n = len(idx)
    v = D.v20_nav(sr, n, ex, f, cost_bps)
    O = {k: np.asarray(f[f"{a}_OPEN"], float)[:n] for k, a in ((1, "QQQ"), (2, "QLD"), (3, "TQQQ"))}
    C = {k: np.asarray(f[f"{a}_CLOSE"], float)[:n] for k, a in ((1, "QQQ"), (2, "QLD"), (3, "TQQQ"))}
    try:
        fx = load_fx(idx, fx_hist)
        krw = {"O": O, "C": C, "fx": fx, "cost": cost_bps / 10000.0}
    except Exception as e:
        print(f"[백테스트] 환율 없음, 원화·세후 생략: {e}", file=sys.stderr); krw = None
    return build_report(idx, ex, v, np.asarray(f["TQQQ_CLOSE"], float)[:n], cost_bps / 10000.0, krw=krw)


def write_html(path, rep):
    t = open(path, encoding="utf-8").read()
    new = "const D = " + json.dumps(rep, ensure_ascii=False, separators=(",", ":")) + ";\n"
    t2, c = re.subn(r"const D = \{.*?\};\n", lambda m: new, t, count=1, flags=re.S)
    if c != 1:
        raise RuntimeError("backtest.html 에서 const D 줄을 찾지 못했습니다")
    open(path, "w", encoding="utf-8").write(t2)
