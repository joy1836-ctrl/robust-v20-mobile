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


def run(data, cfg, cost_bps=5.0):
    f = E.build_features(with_placeholder(data))
    _, ex, _ = E.simulate_strategy(cfg, f, cost_bps / 10000.0)
    return f, ex


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
        raw = yf.download("KRW=X", period="3mo", interval="1d", auto_adjust=True, progress=False, threads=False)
        s = E._normalize_ohlcv(raw)["Close"].dropna().tail(60)
        if len(s) < 5:
            return None
        return {"usdkrw": float(s.iloc[-1]), "chg": float(s.iloc[-1] / s.iloc[-2] - 1),
                "series": [round(float(x), 2) for x in s]}
    except Exception as e:  # FX is optional
        print("[fx] skipped:", e, file=sys.stderr)
        return None


def build_signal(data, cfg, meta, integrity, cost_bps=5.0, synthetic=False, fx=None):
    f, ex = run(data, cfg, cost_bps)
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
                    "series": [round(float(x), 2) for x in c[-60:]]}
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
        "closes": {a: pulse[a]["close"] for a in pulse},
        "pulse": pulse, "fx": fx,
        "indicators": {
            "core_score": int(trend) + int(mom) + int(volok),
            "trend_pass": bool(trend), "mom_pass": bool(mom), "vol_pass": bool(volok),
            "sma180": sma, "roc120": roc, "roc5": g("QQQ_ROC5"), "vol40": vol, "dd252": dd,
            "rsi14": rsi, "mfi14": mfi, "atr_pct14": atr,
            "entry_gate_pass": bool(gate), "crash": bool(crash), "fast_reentry_ready": fast_base,
        },
        "history": [{"date": idx[j].strftime("%Y-%m-%d"), "state": int(ex[j])} for j in range(max(0, n - 40), n)],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data-dir"); ap.add_argument("--period", default="max")
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--config", default=str(HERE / "v20_frozen_config.json"))
    ap.add_argument("--out", default="signal.json")
    ap.add_argument("--cost-bps", type=float, default=5.0)
    ap.add_argument("--no-stale-check", action="store_true", help="시세 신선도 검사 생략 (테스트용)")
    a = ap.parse_args()
    cfg, meta, integ = load_frozen(Path(a.config))
    data = E.load_data(a.data_dir, a.period, a.self_test)
    if not (a.self_test or a.data_dir or a.no_stale_check):
        last, need = data["QQQ"].index[-1].date(), expected_last_session()
        if last < need:
            print(f"[STALE] 시세가 오래됨: 필요한 마감일 {need}, 받은 마지막 날짜 {last}. 오래된 데이터로 신호를 만들지 않습니다.", file=sys.stderr)
            sys.exit(3)
    sig = build_signal(data, cfg, meta, integ, a.cost_bps, synthetic=a.self_test,
                       fx=None if (a.self_test or a.data_dir) else fetch_fx())
    Path(a.out).write_text(json.dumps(sig, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: sig[k] for k in ("asof_close", "next_exec_date", "prev_asset", "target_asset", "signal")}, ensure_ascii=False),
          sig["guide"]["status"])
    if not (integ["engine_ok"] and integ["config_ok"]):
        print("[WARN] engine/config SHA256 mismatch vs frozen config", file=sys.stderr)


if __name__ == "__main__":
    main()
