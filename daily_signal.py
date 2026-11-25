#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ROBUST V20 daily signal generator.

Calls the frozen engine UNMODIFIED. Close[t] -> decision for Open[t+1].
A one-row placeholder session is appended so the engine's own loop produces the
decision for the most recent real close (the loop otherwise stops one row early).
The placeholder never influences indicators (all are causal).

Usage:
  python daily_signal.py                       # download via yfinance
  python daily_signal.py --data-dir ./csv      # QQQ.csv QLD.csv TQQQ.csv (Date,Open,High,Low,Close,Volume)
  python daily_signal.py --self-test           # synthetic data, pipeline check only
Output: signal.json (read by the mobile app and by send_telegram.py)
"""
from __future__ import annotations
import argparse, hashlib, json, sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import robust_engine_v20_frozen as E  # noqa: E402

STATE_NAME = E.STATE_NAME
NY = ZoneInfo("America/New_York")


# ---------------- NYSE calendar (regular full-day closures) ----------------
def _nth(y, m, wd, n):
    d = date(y, m, 1)
    d += timedelta(days=(wd - d.weekday()) % 7)
    return d + timedelta(weeks=n - 1)


def _last(y, m, wd):
    d = date(y + (m == 12), (m % 12) + 1, 1) - timedelta(days=1)
    return d - timedelta(days=(d.weekday() - wd) % 7)


def _obs(d):
    return d - timedelta(days=1) if d.weekday() == 5 else d + timedelta(days=1) if d.weekday() == 6 else d


def nyse_holidays(y):
    from dateutil.easter import easter
    h = {_nth(y, 1, 0, 3), _nth(y, 2, 0, 3), easter(y) - timedelta(days=2), _last(y, 5, 0),
         _obs(date(y, 7, 4)), _nth(y, 9, 0, 1), _nth(y, 11, 3, 4), _obs(date(y, 12, 25))}
    if y >= 2022:
        h.add(_obs(date(y, 6, 19)))
    ny = date(y, 1, 1)
    if ny.weekday() != 5:               # a Saturday New Year is not observed on Friday
        h.add(ny + timedelta(days=1) if ny.weekday() == 6 else ny)
    return h


def next_trading_day(d: date) -> date:
    d += timedelta(days=1)
    while d.weekday() >= 5 or d in nyse_holidays(d.year):
        d += timedelta(days=1)
    return d


def expected_last_session(now_utc=None, buffer_min=30):
    """가장 최근에 '마감이 끝난' NYSE 거래일 (16:00 ET + buffer 이후). 시세가 이보다 오래되면 지연 데이터다."""
    now = (now_utc or datetime.now(timezone.utc)).astimezone(NY)
    d = now.date()
    def is_td(x): return x.weekday() < 5 and x not in nyse_holidays(x.year)
    closed = datetime(d.year, d.month, d.day, 16, 0, tzinfo=NY) + timedelta(minutes=buffer_min)
    if not (is_td(d) and now >= closed):
        d -= timedelta(days=1)
        while not is_td(d):
            d -= timedelta(days=1)
    return d


# ---------------- integrity / engine ----------------
def sha256_file(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def load_frozen(cfg_path: Path):
    meta = json.loads(cfg_path.read_text(encoding="utf-8"))
    cfg_hash = hashlib.sha256(json.dumps(meta["config"], sort_keys=True, separators=(",", ":")).encode()).hexdigest()
    integrity = {
        "engine_ok": sha256_file(HERE / "robust_engine_v20_frozen.py") == meta["engine_sha256"],
        "config_ok": cfg_hash == meta["config_sha256"],
    }
    return E.Config(**meta["config"]), meta, integrity


def with_placeholder(data):
    out = {}
    for k, df in data.items():
        row = df.iloc[[-1]].copy()
        row.index = [df.index[-1] + pd.offsets.BDay(1)]
        out[k] = pd.concat([df, row])
    return out


def run_full(data, cfg, cost_bps=8.0):
    f = E.build_features(with_placeholder(data))
    sr, ex, _ = E.simulate_strategy(cfg, f, cost_bps / 10000.0)
    return f, ex, sr


def run(data, cfg, cost_bps=8.0):
    f, ex, _ = run_full(data, cfg, cost_bps)
    return f, ex


def v20_nav(sr, n, ex, f, cost_bps=8.0):
    """V20 전략 지수(종가 기준, 시작=1.0). 실제 주문 방식과 같은 체결 규칙을 쓴다.
    - 현금 → 매수(매도 없음): 시초가 매수 (그날 시초→종가 수익 반영)
    - 보유 → 현금: 시초가 매도
    - 자산 갈아타기(매도 있음): 시초가 매도, 종가 매수 (새 자산은 그날 종가부터 보유)
    한쪽당 비용 cost_bps. 엔진의 상태(ex)는 그대로 쓰고 가치 계산만 이 규칙으로 한다."""

    c = cost_bps / 10000.0
    ex = np.asarray(ex)[:n]
    O, C = {}, {}
    for k, a in ((1, "QQQ"), (2, "QLD"), (3, "TQQQ")):
        O[k] = np.asarray(f[f"{a}_OPEN"], dtype=float)[:n]
        C[k] = np.asarray(f[f"{a}_CLOSE"], dtype=float)[:n]
    nav = np.empty(n, dtype=float)
    v = 1.0
    for i in range(n):
        cur = int(ex[i])
        prv = int(ex[i - 1]) if i > 0 else 0
        if cur == prv:
            if cur:
                v *= C[cur][i] / (C[cur][i - 1] if i > 0 else O[cur][i])
        else:
            if prv:                                    # 시초가 매도
                v *= (O[prv][i] / C[prv][i - 1]) * (1.0 - c)
            if cur:
                v *= (1.0 - c)                         # 매수 비용
                if not prv:                            # 현금에서 매수: 시초가 매수
                    v *= C[cur][i] / O[cur][i]
        nav[i] = v
    return nav


CHART_BASE_DATE = "2021-01-04"   # 차트·V20 지수 기준일(2021년 첫 개장일). 고정이라 매일 바뀌지 않음
V20_BASE = 10.0            # V20 지수 눈금: 기준일 = 10


def chart_start(dates):
    """기준일(포함) 이후 첫 거래일의 위치. 데이터가 기준일보다 짧으면 0."""
    for k, d in enumerate(dates):
        if d.strftime("%Y-%m-%d") >= CHART_BASE_DATE:
            return k
    return 0


def ytd_mdd(c, ds):
    """올해 최대 낙폭: 전년도 마지막 종가를 시작점으로, 올해 안에서 (종가 ÷ 그때까지 최고 종가 − 1)의 최저값."""
    c = np.asarray(c, dtype=float)
    yr = ds[len(c) - 1].year
    prev = [k for k, d in enumerate(ds[:len(c)]) if d.year < yr]
    if not prev:
        return None
    k0 = prev[-1]
    seg = c[k0:]
    dd = seg / np.maximum.accumulate(seg) - 1.0
    j = int(np.argmin(dd))
    pk = int(np.argmax(seg[: j + 1]))
    return {"ytd_mdd": float(dd[j]), "ytd_mdd_date": ds[k0 + j].strftime("%Y-%m-%d"), "ytd_mdd_peak_date": ds[k0 + pk].strftime("%Y-%m-%d")}


def v20_scaled(nav, dates):
    """기준일(2021년 첫 개장일) 종가 지수를 V20_BASE(10)로 맞춘 지수."""
    return nav / nav[chart_start(dates)] * V20_BASE


def v20_pulse(sr, n, dates, ex, f, window=252, tail=60, cost_bps=8.0):
    """V20 카드용 지표(종가 기준, 5년 전=10 눈금)."""
    nav = v20_nav(sr, n, ex, f, cost_bps)
    ds = list(dates)[:n]
    sc = v20_scaled(nav, ds)
    idx = sc[-(window + 1):]                       # nav[k] 는 ds[k] 일자 종가 기준 지수
    yr = ds[-1].year
    prev = [k for k, d in enumerate(ds) if d.year < yr]
    ytd = float(nav[-1] / nav[prev[-1]] - 1.0) if prev else None
    k = int(np.argmax(nav))                    # 전체 기간 최고점(역대 고점)
    ath = {"dd_all": float(nav[-1] / nav[k] - 1.0), "peak_date": ds[k].strftime("%Y-%m-%d"),
           "days_since_peak": int(len(nav) - 1 - k), "since": ds[0].strftime("%Y-%m-%d")}
    bd = ds[chart_start(ds)]
    mdd = ytd_mdd(nav, ds) or {}
    return {**ath, **mdd, "base": V20_BASE, "base_date": bd.strftime("%Y-%m-%d"), "close": round(float(idx[-1]), 3), "ytd": ytd, "ytd_year": yr, "chg": float(idx[-1] / idx[-2] - 1),
            "dd252": float(idx[-1] / idx.max() - 1), "series": [round(float(x), 3) for x in idx[-tail:]],
            "dates": [d.strftime("%Y-%m-%d") for d in ds[-tail:]], "kind": "index"}


def classify(prev: int, tgt: int) -> str:
    if tgt == prev:
        return "HOLD"
    if tgt == 0:
        return "EXIT_TO_CASH"
    if prev == 0:
        return "REENTER"
    return "STEP_UP" if tgt > prev else "STEP_DOWN"


def guidance(signal: str, target: str, gate_pass: bool):
    """Per-asset notes + headline status. Shared by the app and the Telegram message."""
    fresh = signal != "HOLD"
    heads = {"EXIT_TO_CASH": "전량 매도", "STEP_DOWN": "하향 매도", "STEP_UP": "상향 매수", "REENTER": "신규 매수"}
    if fresh:
        status = heads[signal]
    elif target == "CASH":
        status = "현금 유지"
    else:
        status = "보유 유지" if gate_pass else "신규 매수 대기"
    g = {}
    for a in ("QQQ", "QLD", "TQQQ"):
        if target == "CASH":
            g[a] = "보유 시 매도 (현금 전환)"
        elif a == target:
            g[a] = f"{a} 매수" if fresh else ("보유 유지 / 신규 매수 가능" if gate_pass else "보유 유지 / 신규 매수 대기")
        else:
            g[a] = f"보유 시 매도 → {target} 전환"
    note = ""
    if not fresh and target != "CASH" and not gate_pass:
        note = "진입 게이트 미충족: 신규·추가 매수는 대기합니다. 비목표 자산 매도 안내는 유지됩니다. (GATE_CLOSED)"
    return {"status": status, "assets": g, "note": note, "fresh": fresh, "gate_closed": not gate_pass}


def fetch_fx():
    try:
        import yfinance as yf
        raw = yf.download("KRW=X", period="6y", interval="1d", auto_adjust=True, progress=False, threads=False)
        full = E._normalize_ohlcv(raw)["Close"].dropna()
        s = full.tail(60)
        if len(s) < 5:
            return None
        out = {"usdkrw": float(s.iloc[-1]), "chg": float(s.iloc[-1] / s.iloc[-2] - 1),
               "series": [round(float(x), 2) for x in s],
               "dates": [d.strftime("%Y-%m-%d") for d in s.index],
               "_hist": {"dates": [d.strftime("%Y-%m-%d") for d in full.index if d.strftime("%Y-%m-%d") >= CHART_BASE_DATE],
                         "series": [round(float(x), 2) for d, x in full.items() if d.strftime("%Y-%m-%d") >= CHART_BASE_DATE]}}
        prev = full[full.index.year < full.index[-1].year]   # 전년도 마지막 환율 = 올해 수익률 기준점
        if len(prev):
            out["ytd_base"] = round(float(prev.iloc[-1]), 2)
            out["ytd_base_date"] = prev.index[-1].strftime("%Y-%m-%d")
        return out
    except Exception as e:  # FX is optional
        print("[fx] skipped:", e, file=sys.stderr)
        return None


def build_signal(data, cfg, meta, integrity, cost_bps=8.0, synthetic=False, fx=None):
    f, ex, sr = run_full(data, cfg, cost_bps)
    n = len(data["QQQ"])
    i = n - 1
    prev, tgt = int(ex[i]), int(ex[i + 1])
    g = lambda k: float(f[k][i])
    q = g("QQQ_CLOSE")
    sma, roc, vol, dd = g("QQQ_SMA180"), g("QQQ_ROC120"), g("QQQ_VOL40"), g("QQQ_DD252")
    trend, mom, volok = q > sma, roc > cfg.mom_threshold, vol < cfg.vol_threshold
    crash = (not trend) and dd <= cfg.crash_dd and vol >= cfg.crash_vol_threshold
    rsi, mfi, atr = g("QQQ_RSI14"), g("QQQ_MFI14"), g("QQQ_ATRPCT14")
    gate = rsi >= cfg.rsi_min and mfi >= cfg.mfi_min and atr <= cfg.atr_max_pct
    fast_base = bool(q > g("QQQ_EMA15") and g("QQQ_ROC15") >= cfg.fast_threshold and roc > cfg.fast_long_mom_floor and not crash)
    idx = data["QQQ"].index
    asof = idx[-1].date()
    nxt = next_trading_day(asof)
    valid_until = datetime(nxt.year, nxt.month, nxt.day, 16, 0, tzinfo=NY).astimezone(timezone.utc)
    sig = classify(prev, tgt)
    guide = guidance(sig, STATE_NAME[tgt], bool(gate))
    pulse = {}
    for a in ("QQQ", "QLD", "TQQQ"):
        c = f[f"{a}_CLOSE"][: n]
        pulse[a] = {"close": float(c[-1]), "chg": float(c[-1] / c[-2] - 1), "dd252": g(f"{a}_DD252"),
                    "series": [round(float(x), 2) for x in c[-60:]],
                    "dates": [d.strftime("%Y-%m-%d") for d in idx[-60:]]}
        k = int(np.nanargmax(c))                       # 전체 기간 최고 종가
        pulse[a].update({"dd_all": float(c[-1] / c[k] - 1.0), "peak_date": idx[k].strftime("%Y-%m-%d"),
                         "days_since_peak": int(n - 1 - k), "since": idx[0].strftime("%Y-%m-%d")})
        prevk = [j for j in range(n) if idx[j].year < idx[n - 1].year]   # 전년도 마지막 종가 = 올해 수익률 기준점
        if prevk:
            y = float(c[n - 1] / c[prevk[-1]] - 1.0)
            pulse[a].update({"ytd": y, "ytd_year": int(idx[n - 1].year), **(ytd_mdd(c, idx) or {})})
            if fx and fx.get("ytd_base"):
                pulse[a]["ytd_krw"] = float((1.0 + y) * (fx["usdkrw"] / fx["ytd_base"]) - 1.0)
    closes = {a: pulse[a]["close"] for a in pulse}
    pulse["V20"] = v20_pulse(sr, n, data["QQQ"].index, ex, f, cost_bps=cost_bps)
    v = pulse["V20"]
    if fx and fx.get("ytd_base") and v.get("ytd") is not None:   # 원화 기준 = (1+달러수익률) × 환율변화 − 1
        v["ytd_krw"] = float((1.0 + v["ytd"]) * (fx["usdkrw"] / fx["ytd_base"]) - 1.0)
    c0 = chart_start(idx[:n])                  # 기준일 이후 이력. 차트용 장기 이력은 signal.json이 아니라 history.json에 따로 저장
    vnav = v20_nav(sr, n, ex, f, cost_bps)
    chart = {"dates": [d.strftime("%Y-%m-%d") for d in idx[c0:n]]}
    for a in ("QQQ", "QLD", "TQQQ"):
        chart[a] = [round(float(x), 2) for x in np.asarray(f[f"{a}_CLOSE"], dtype=float)[:n][c0:]]
    chart["V20"] = [round(float(x), 3) for x in v20_scaled(vnav, idx[:n])[c0:]]
    return {
        "strategy": meta["strategy_name"], "freeze_date": meta["freeze_date"],
        "integrity": {**integrity, "synthetic_data": synthetic},
        "generated_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "asof_close": asof.strftime("%Y-%m-%d"),
        "next_exec_date": nxt.strftime("%Y-%m-%d"),
        "valid_until_utc": valid_until.strftime("%Y-%m-%dT%H:%M:%S+00:00"),
        "prev_state": prev, "prev_asset": STATE_NAME[prev],
        "target_state": tgt, "target_asset": STATE_NAME[tgt],
        "signal": sig, "guide": guide,
        "closes": closes,
        "pulse": pulse, "fx": fx, "_chart": chart,
        "indicators": {
            "core_score": int(trend) + int(mom) + int(volok),
            "trend_pass": bool(trend), "mom_pass": bool(mom), "vol_pass": bool(volok),
            "sma180": sma, "roc120": roc, "roc5": g("QQQ_ROC5"), "vol40": vol, "dd252": dd,
            "rsi14": rsi, "mfi14": mfi, "atr_pct14": atr,
            "entry_gate_pass": bool(gate), "crash": bool(crash), "fast_reentry_ready": fast_base,
        },
        "history": [{"date": idx[j].strftime("%Y-%m-%d"), "state": int(ex[j])} for j in range(max(0, n - 40), n)],
    }


def write_intraday(path, live, cfg, meta, integ, cost_bps):
    """참고용 장중 신호: 미완성 봉(현재가)을 오늘 종가라고 가정하고 같은 엔진으로 계산한다.
    정식 신호(signal.json)는 건드리지 않고 별도 파일에 쓴다. 장중이 아니면 파일을 지운다."""
    if live is None:
        if path.exists():
            path.unlink()
        return
    try:
        if any(live[k]["Close"].iloc[-1] != live[k]["Close"].iloc[-1] for k in ("QQQ", "QLD", "TQQQ")):
            raise ValueError("현재가 없음")
        sg = build_signal(live, cfg, meta, integ, cost_bps, synthetic=False, fx=None)
        keep = {k: sg[k] for k in ("asof_close", "next_exec_date", "prev_asset", "target_asset", "signal", "guide", "closes", "indicators")}
        keep.update({"intraday": True, "generated_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
                     "note": "장중 현재가를 오늘 종가라고 가정한 참고용 신호입니다. 마감 때 달라질 수 있습니다."})
        path.write_text(json.dumps(keep, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[장중 참고] {keep['asof_close']} 현재가 기준: {keep['prev_asset']} -> {keep['target_asset']} ({keep['signal']})", file=sys.stderr)
    except Exception as e:   # 참고용이므로 실패해도 정식 신호에는 영향 없음
        if path.exists():
            path.unlink()
        print(f"[장중 참고] 생성 실패(무시): {e}", file=sys.stderr)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir"); ap.add_argument("--period", default="max")
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--config", default=str(HERE / "v20_frozen_config.json"))
    ap.add_argument("--out", default="signal.json")
    ap.add_argument("--cost-bps", type=float, default=8.0)
    ap.add_argument("--no-stale-check", action="store_true", help="시세 신선도 검사 생략 (테스트용)")
    a = ap.parse_args()
    cfg, meta, integ = load_frozen(Path(a.config))
    data = E.load_data(a.data_dir, a.period, a.self_test)
    live = None   # 장중 미완성 봉이 붙어 있으면 참고용 장중 신호에 쓴다
    if not (a.self_test or a.data_dir or a.no_stale_check):
        last, need = data["QQQ"].index[-1].date(), expected_last_session()
        if last > need:   # 장중에 실행하면 오늘의 미완성 봉이 끝에 붙는다 -> 버리고 마지막 마감일 기준으로 계산
            cut = pd.Timestamp(need)
            live = data
            print(f"[장중] 미완성 봉({last}) 제외, 마지막 마감일 {need} 기준으로 계산합니다.", file=sys.stderr)
            data = {k: df[df.index <= cut] for k, df in data.items()}
            last = data["QQQ"].index[-1].date()
        if last < need:
            print(f"[STALE] 시세가 오래됨: 필요한 마감일 {need}, 받은 마지막 날짜 {last}. 오래된 데이터로 신호를 만들지 않습니다.", file=sys.stderr)
            sys.exit(3)
    sig = build_signal(data, cfg, meta, integ, a.cost_bps, synthetic=a.self_test,
                       fx=None if (a.self_test or a.data_dir) else fetch_fx())
    chart = sig.pop("_chart", None)
    fxh = (sig.get("fx") or {}).pop("_hist", None)
    if chart is not None:
        if fxh:
            chart["fx"] = fxh
        Path(a.out).with_name("history.json").write_text(json.dumps(chart, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    Path(a.out).write_text(json.dumps(sig, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: sig[k] for k in ("asof_close", "next_exec_date", "prev_asset", "target_asset", "signal")}, ensure_ascii=False),
          sig["guide"]["status"])
    write_intraday(Path(a.out).with_name("signal_intraday.json"), live, cfg, meta, integ, a.cost_bps)
    bt = Path(a.out).with_name("backtest.html")
    if bt.exists() and not a.self_test:   # 백테스트 화면 데이터도 최신으로 갱신(실패해도 정식 신호에는 영향 없음)
        try:
            import backtest_report as BR
            BR.write_html(bt, BR.make(data, cfg, a.cost_bps))
            print("[백테스트] 갱신:", data["QQQ"].index[-1].date(), file=sys.stderr)
        except Exception as e:
            print(f"[백테스트] 갱신 실패(무시): {e}", file=sys.stderr)
    if not (integ["engine_ok"] and integ["config_ok"]):
        print("[WARN] engine/config SHA256 mismatch vs frozen config", file=sys.stderr)


if __name__ == "__main__":
    main()
