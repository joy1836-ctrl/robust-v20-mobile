"""백테스트 보고서 데이터(D)를 최신 시세로 다시 만든다. docs/backtest.html 의 `const D = {...};` 줄을 갱신하는 데 쓴다.
기간: 마지막 거래일에서 거꾸로 15년(252거래일×15). V20 = 종가 기준 지수(daily_signal.v20_nav), 비교 = TQQQ 매수 후 보유."""
import json, re
import numpy as np
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


def build_report(idx, ex, v_nav, t_close, cost, label_extra=""):
    """idx: 날짜(DatetimeIndex), ex: 시가 상태, v_nav: V20 가치(시계열, 길이 n), t_close: TQQQ 가격."""
    n = len(idx)
    L = YEARS * H
    s = max(0, n - 1 - L)
    d = idx[s:n]
    ex = np.asarray(ex)[:n][s:]
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
    end = d[-1].strftime("%Y-%m-%d")
    return {"asof": f"ROBUST V20 vs TQQQ | Fixed 15Y Backtest as of {end}{label_extra}", "start": d[0].strftime("%Y-%m-%d"), "end": end, "cost": cost,
            "full": full, "activity": activity, "exposure": exposure, "cal": cal, "roll": roll, "series": series, "trans": trans,
            "eq": {"d": [x.strftime("%Y-%m-%d") for x in d], "s": [round(float(x), 3) for x in r_eq], "t": [round(float(x), 4) for x in t_eq], "st": [int(x) for x in ex]}}


def make(data, cfg, cost_bps=8.0):
    f, ex, sr = D.run_full(data, cfg, cost_bps)
    idx = data["QQQ"].index
    n = len(idx)
    v = D.v20_nav(sr, n, ex, f, cost_bps)
    return build_report(idx, ex, v, np.asarray(f["TQQQ_CLOSE"], float)[:n], cost_bps / 10000.0)


def write_html(path, rep):
    t = open(path, encoding="utf-8").read()
    new = "const D = " + json.dumps(rep, ensure_ascii=False, separators=(",", ":")) + ";\n"
    t2, c = re.subn(r"const D = \{.*?\};\n", lambda m: new, t, count=1, flags=re.S)
    if c != 1:
        raise RuntimeError("backtest.html 에서 const D 줄을 찾지 못했습니다")
    open(path, "w", encoding="utf-8").write(t2)
