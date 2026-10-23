#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ROBUST Tight Rolling Lab V21
========================

Purpose
-------
Search tighter variants of the original ROBUST QQQ/QLD/TQQQ/CASH state strategy,
using the proven V12 stable-split core and a focused V15 impulse-veto neighborhood with 12-process parallel search.
The final selection is NOT based on single-period CAGR alone. It prioritizes
rolling-window consistency vs TQQQ Buy & Hold across:

    6 months, 1 year, 2 years, ... 10 years

For finalists, every possible daily rolling start is checked for BOTH:
    1) strategy return >= TQQQ Buy & Hold return
    2) strategy MDD is better (less negative) than TQQQ Buy & Hold MDD

The script produces one Excel workbook with a dashboard and detailed sheets.

Default Windows CMD usage
-------------------------
    python robust_tight_rolling_lab_v21.py

Deep search example
-------------------
    python robust_tight_rolling_lab_v21.py --trials 480000 --workers 12 --finalists 480 --detail-top 40 --audit-top 16

CSV mode (QQQ.csv, QLD.csv, TQQQ.csv in the folder)
--------------------------------------------------
    python robust_tight_rolling_lab_v21.py --data-dir C:\\market_data --workers 12

Fast smoke test without internet
--------------------------------
    python robust_tight_rolling_lab_v21.py --self-test --trials 120 --workers 4 --finalists 5

Dependencies
------------
    pip install numpy pandas openpyxl yfinance

Notes
-----
* Signals are calculated on CLOSE and executed at the NEXT session OPEN.
* Prices downloaded from yfinance are auto-adjusted.
* Transaction cost is charged on each buy/sell side at the configured bps.
* Original ROBUST is always included as a baseline candidate.
* A strategy that wins only a few lucky periods is penalized heavily.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import platform
import random
import sys
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime
from multiprocessing import freeze_support
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import numpy as np
    import pandas as pd
except ImportError as exc:
    print("[ERROR] numpy/pandas are required.")
    print("Run: pip install numpy pandas openpyxl yfinance")
    raise

try:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("[ERROR] openpyxl is required.")
    print("Run: pip install openpyxl")
    raise

TRADING_DAYS = 252
HORIZONS: List[Tuple[str, int]] = [
    ("6M", 126),
    ("1Y", 252),
    ("2Y", 504),
    ("3Y", 756),
    ("4Y", 1008),
    ("5Y", 1260),
    ("6Y", 1512),
    ("7Y", 1764),
    ("8Y", 2016),
    ("9Y", 2268),
    ("10Y", 2520),
]
STATE_NAME = {0: "CASH", 1: "QQQ", 2: "QLD", 3: "TQQQ"}
SCRIPT_VERSION = "2026-09-12-rolling-lab-v21-bidirectional-preview-overlay"

# -----------------------------
# Strategy configuration
# -----------------------------

@dataclass(frozen=True, slots=True)
class Config:
    name: str
    family: str

    # ROBUST core
    sma_days: int = 165
    mom_days: int = 120
    mom_threshold: float = 0.00
    vol_days: int = 40
    vol_threshold: float = 0.25
    fast_ema_days: int = 10
    fast_mom_days: int = 10
    fast_threshold: float = 0.04
    fast_long_mom_floor: float = -0.08
    crash_dd: float = -0.23
    crash_vol_threshold: float = 0.25
    down_confirm: int = 2

    # V6 score-aware hysteresis: hold through a mild 3->2 wobble, but react faster to severe score breaks.
    adaptive_down_mode: str = "off"  # off, tqqq_score
    soft_down_confirm: int = 3         # required days when TQQQ and base score == 2
    hard_down_confirm: int = 1         # required days when TQQQ and base score <= 1

    # Fast re-entry behavior
    fast_mode: str = "direct"  # direct, qld, conditional, normal
    gate_scope: str = "all_up"  # all_up, tqqq_only, fast_only
    gate_logic: str = "all"      # all, rma_2of3 (RSI/MFI/ATR: any 2 of 3)

    # Optional entry filters
    require_qqq_ma50: bool = False
    require_qqq_ma200: bool = False
    require_tqqq_ma20: bool = False
    require_tqqq_ma50: bool = False
    macd_gate: str = "none"  # none, qqq, tqqq, both
    rsi_asset: str = "none"  # none, qqq, tqqq
    rsi_min: int = 50
    stoch_asset: str = "none"
    cci_asset: str = "none"
    boll_asset: str = "none"
    boll_rule: str = "mid"  # mid, upper_half, not_overbought
    adx_asset: str = "none"
    adx_min: int = 18
    mfi_asset: str = "none"
    mfi_min: int = 50
    willr_asset: str = "none"
    willr_min: int = -50
    obv_asset: str = "none"
    donchian_asset: str = "none"
    donchian_min_pos: float = 0.50
    atr_asset: str = "none"
    atr_max_pct: float = 0.08
    roc_asset: str = "none"
    roc_days: int = 20
    roc_min: float = 0.00

    # v15-inspired composite defense gate
    defense_mode: str = "off"  # off, cap2, cap1, cash
    defense_score_threshold: int = 8

    # Optional immediate TQQQ-only de-risking (v15-inspired, but milder than full cash)
    tqqq_exit_mode: str = "off"  # off, ma20_macd, qqq50_macd, risk
    tqqq_exit_score: int = 7

    # V3 recovery / participation modules. These are deliberately narrow so defense is preserved.
    recovery_mode: str = "off"  # off, relaxed_fast, strong_roc, hybrid
    recovery_days: int = 63
    recovery_mom_floor: float = -0.25
    recovery_roc20_min: float = 0.05
    recovery_vol_max: float = 0.60

    # V6 generic snapback after ANY real TQQQ deleveraging, not just an official crash.
    # Designed for generic choppy correction/recovery regimes.
    snapback_mode: str = "off"  # off, qld, gated3, direct3
    snapback_days: int = 20
    snapback_roc10_min: float = 0.01
    snapback_rsi_min: int = 50
    snapback_vol_max: float = 0.50
    snapback_ema_days: int = 20
    snapback_dd_floor: float = -0.15

    # V7 early crash-release: a strong, confirmed reversal can restore some exposure
    # even while the long-horizon crash flag is still active. No calendar/date rules.
    crash_release_mode: str = "off"  # off, qld, gated3, direct3
    crash_release_wait: int = 5
    crash_release_roc5_min: float = 0.06
    crash_release_roc10_min: float = 0.10
    crash_release_rsi_min: int = 45
    crash_release_mfi_min: int = 40
    crash_release_vol_max: float = 1.00
    crash_release_dd_floor: float = -0.45

    # V8 staged re-entry after any real TQQQ deleveraging. Ordinary corrections recover
    # to QLD first; TQQQ is restored only after a second, stronger confirmation.
    # Recovery ATR is intentionally separate from the normal ATR gate so the V6 ATR8
    # defense can remain intact outside the recovery window.
    staged_reentry_mode: str = "off"  # off, qld_gate3, qld_confirm3
    staged_window_days: int = 15
    staged_ema_days: int = 10
    staged_roc10_min: float = 0.00
    staged_rsi_min: int = 45
    staged_mfi_min: int = 40
    staged_atr_max: float = 0.12
    staged_vol_max: float = 0.60
    staged_dd_floor: float = -0.30
    staged_promote_wait: int = 2
    staged_promote_roc5_min: float = 0.02
    staged_promote_roc10_min: float = 0.04
    staged_promote_rsi_min: int = 50
    staged_promote_mfi_min: int = 45
    staged_promote_atr_max: float = 0.10

    # V8 explosive rebound. Only a recently observed genuine crash can arm this path.
    # A very strong short-horizon reversal can jump directly to TQQQ, then a short latch
    # prevents the slow long-horizon crash flag from immediately forcing CASH again.
    explosive_rebound_mode: str = "off"  # off, direct3, gated3
    explosive_window_days: int = 15
    explosive_min_crash_days: int = 2
    explosive_roc5_min: float = 0.08
    explosive_roc10_min: float = 0.15
    explosive_rsi_min: int = 50
    explosive_mfi_min: int = 40
    explosive_atr_max: float = 0.18
    explosive_vol_max: float = 1.50
    explosive_hold_days: int = 5
    explosive_hold_roc5_min: float = 0.00

    # V9 generic correction re-entry. This path is armed ONLY after a real TQQQ
    # deleveraging that was not caused by the deep-crash flag. It targets ordinary
    # whipsaw/correction recoveries without weakening crash defense.
    correction_reentry_mode: str = "off"  # off, qld, gated3, direct3
    correction_window_days: int = 15
    correction_ema_days: int = 10
    correction_roc5_min: float = 0.00
    correction_roc10_min: float = -0.01
    correction_rsi_min: int = 45
    correction_mfi_min: int = 35
    correction_atr_max: float = 0.12
    correction_vol_max: float = 0.45
    correction_dd_floor: float = -0.25

    # V9 crash drawdown-rebound release. During a genuine crash episode the engine
    # remembers the worst 252d drawdown. A recovery of X percentage points from that
    # trough can re-arm exposure, subject to generic momentum/liquidity filters.
    # No calendar/date conditions are used.
    dd_rebound_mode: str = "off"  # off, qld, gated3, direct3
    dd_rebound_window_days: int = 42
    dd_rebound_min: float = 0.08
    dd_rebound_roc5_min: float = 0.02
    dd_rebound_rsi_min: int = 45
    dd_rebound_mfi_min: int = 35
    dd_rebound_atr_max: float = 0.25
    dd_rebound_vol_max: float = 1.50
    dd_rebound_hold_days: int = 5
    dd_rebound_hold_roc5_min: float = -0.03

    # V11 trough-confirmed crash recovery (DDR2). This is armed only by a genuine
    # crash episode, tracks a separate rolling crash trough, resets on every new low,
    # and requires the trough to age plus ROC10 confirmation before restoring risk.
    # It is deliberately separate from legacy DDR so all V8/V9/V10 anchors remain exact.
    trough_recovery_mode: str = "off"  # off, direct3, gated3, qld_confirm3, qld_split3
    trough_recovery_window_days: int = 42
    trough_age_min: int = 4
    trough_min_crash_days: int = 3
    trough_rebound_min: float = 0.09
    trough_roc5_min: float = 0.06
    trough_roc10_min: float = 0.00
    trough_rsi_min: int = 48
    trough_mfi_min: int = 40
    trough_atr_max: float = 0.25
    trough_vol_max: float = 1.50
    trough_confirm_days: int = 1
    trough_hold_days: int = 5
    trough_hold_roc5_min: float = -0.03
    trough_hold_roc10_min: float = -0.05

    # V12 split QLD bridge -> TQQQ promotion. Legacy qld_confirm3 keeps V11 semantics
    # exactly; qld_split3 uses the V11 trough trigger only for QLD entry, then a
    # separate stronger promotion gate. Fresh crash lows still reset the whole path.
    trough_bridge_max_days: int = 10
    trough_promote_rebound_min: float = 0.11
    trough_promote_roc5_min: float = 0.02
    trough_promote_roc10_min: float = 0.08
    trough_promote_rsi_min: int = 50
    trough_promote_mfi_min: int = 45
    trough_promote_atr_max: float = 0.10
    trough_promote_vol_max: float = 0.85

    # V12 ordinary re-promotion guard. It NEVER delays TQQQ exits and is disabled
    # inside the bounded post-crash recovery context. It only filters ordinary
    # lower-leverage -> TQQQ re-promotions after repeated recent TQQQ delevers.
    repromote_guard_mode: str = "off"  # off, score3_only, score3_or_strong, cooldown_strong
    repromote_lookback_days: int = 60
    repromote_trigger_count: int = 2
    repromote_cooldown_days: int = 5
    repromote_roc10_min: float = 0.02
    repromote_rsi_min: int = 50
    repromote_vol_util_max: float = 0.95

    # V14 selective soft-volatility re-promotion. Unlike raising the global volatility
    # threshold, this can ONLY promote QLD -> TQQQ when long trend and long momentum
    # are both bullish and volatility is the sole failed core vote. It never delays a
    # TQQQ exit, and it is disabled during/soon after crash episodes.
    soft_vol_mode: str = "off"  # off, recent3, score2_strong3
    soft_vol_max: float = 0.33
    soft_vol_recent_delever_days: int = 20
    soft_vol_confirm_days: int = 1
    soft_vol_roc10_min: float = 0.02
    soft_vol_roc20_min: float = 0.04
    soft_vol_roc120_min: float = 0.00
    soft_vol_rsi_min: int = 52
    soft_vol_mfi_min: int = 45
    soft_vol_dd_floor: float = -0.15

    # V15 selective impulse exit veto. This is intentionally much narrower than the
    # rejected V10 adaptive/hysteresis idea. It can only veto an ORDINARY TQQQ->QLD
    # step-down for a very short time when short-term rebound evidence is strong,
    # volatility itself is still acceptable, and the market is outside crash recovery.
    # Later crash/defense/TQQQ-exit rules can still override the veto on the same day.
    impulse_veto_mode: str = "off"  # off, score1_strong, score2_strong, score12_strong
    impulse_veto_max_days: int = 1
    impulse_veto_ema_days: int = 5
    impulse_veto_roc5_min: float = 0.04
    impulse_veto_roc10_min: float = 0.04
    impulse_veto_roc20_min: float = -0.01
    impulse_veto_rsi_min: int = 52
    impulse_veto_mfi_min: int = 40
    impulse_veto_atr_max: float = 0.04
    impulse_veto_vol_max: float = 0.35
    impulse_veto_dd_floor: float = -0.15

    # V16 precision transition veto. This is layered AFTER the proven V15 veto and
    # only searches for additional one-session TQQQ->QLD exits that are likely to be
    # transient. Modes explicitly distinguish which slow vote failed; no calendar rules.
    precision_veto_mode: str = "off"  # off, trend_near, mom_near, score1_near, fresh_reentry
    precision_veto_max_days: int = 1
    precision_veto_ema_days: int = 10
    precision_veto_roc5_min: float = 0.02
    precision_veto_roc10_min: float = 0.02
    precision_veto_roc20_min: float = -0.02
    precision_veto_rsi_min: int = 50
    precision_veto_mfi_min: int = 40
    precision_veto_atr_max: float = 0.04
    precision_veto_vol_max: float = 0.35
    precision_veto_dd_floor: float = -0.18
    precision_veto_trend_gap_floor: float = -0.03  # QQQ/SMA - 1; only relevant when trend vote fails
    precision_veto_mom_gap_floor: float = -0.02    # Mom - threshold; only relevant when momentum vote fails
    precision_veto_recent_reentry_days: int = 3   # used by fresh_reentry mode
    precision_veto_score_min: int = 1              # fresh_reentry can require score >=1 or >=2

    # V17 recent re-entry grace. V16 showed that strong score-1 rebounds were a trap and
    # that profitable missed exits were concentrated in score-2 states shortly after a
    # TQQQ re-entry. This layer therefore excludes score-1 entirely and only considers
    # one-session ordinary TQQQ->QLD exits. It can distinguish which single slow vote
    # failed and can relax the old blanket 63-session post-crash exclusion without ever
    # blocking an active crash exit. No dates/years are encoded.
    reentry_grace_mode: str = "off"  # off, mom_reentry, trend_reentry, score2_reentry, split_reentry
    reentry_grace_max_days: int = 1
    reentry_grace_min_age: int = 1
    reentry_grace_max_age: int = 20
    reentry_grace_min_post_crash_days: int = 30
    reentry_grace_entry_gate: str = "pass"  # pass, any
    reentry_grace_ema_days: int = 0  # 0 disables EMA requirement; otherwise 10 or 20
    reentry_grace_trend_gap_floor: float = -0.03
    reentry_grace_mom_gap_floor: float = -0.015
    reentry_grace_roc5_min: float = -0.04
    reentry_grace_roc10_min: float = -0.06
    reentry_grace_roc20_min: float = -0.08
    reentry_grace_rsi_min: int = 40
    reentry_grace_mfi_min: int = 35
    reentry_grace_atr_max: float = 0.04
    reentry_grace_vol_max: float = 0.35
    reentry_grace_dd_floor: float = -0.25

    # V18 split Score-2 branch controls. Momentum-fail and Trend-fail exits showed
    # different safe ROC5 / Entry-Gate behavior in V17. Split mode evaluates them
    # separately while keeping all other score-2 / crash safety conditions common.
    reentry_grace_mom_entry_gate: str = "pass"   # pass, any
    reentry_grace_trend_entry_gate: str = "any"  # pass, any
    reentry_grace_mom_roc5_min: float = -0.03
    reentry_grace_trend_roc5_min: float = 0.00

    # V19 micro rescue layer. V18 already captures the profitable recent Score-2 exits,
    # so V19 does NOT broaden that rule. Instead it independently tests three sparse
    # residual state families found in the generic exit audit: (1) a recent Score-1
    # pullback with negative ROC5 but positive ROC10, (2) a Score-1 near-trend boundary
    # where the long trend is only barely below threshold, and (3) a mature TQQQ run
    # where Momentum is only microscopically below its long threshold. No dates/years.
    micro_rescue_mode: str = "off"  # off, score1_pullback, score1_neartrend, mature_mom, combo
    micro_rescue_max_days: int = 1
    micro_rescue_min_post_crash_days: int = 30
    micro_rescue_recent_max_age: int = 10
    micro_rescue_neartrend_max_age: int = 60
    micro_rescue_mature_min_age: int = 60
    micro_rescue_mature_max_age: int = 250
    micro_rescue_score1_trend_gap_floor: float = -0.05
    micro_rescue_score1_mom_gap_floor: float = -0.012
    micro_rescue_score1_roc5_min: float = -0.03
    micro_rescue_score1_roc5_max: float = 0.00
    micro_rescue_score1_roc10_min: float = 0.00
    micro_rescue_score1_roc20_min: float = -0.05
    micro_rescue_score1_rsi_min: int = 40
    micro_rescue_score1_mfi_min: int = 35
    micro_rescue_neartrend_gap_floor: float = -0.01
    micro_rescue_neartrend_roc5_min: float = -0.03
    micro_rescue_neartrend_roc5_max: float = 0.03
    micro_rescue_neartrend_roc20_min: float = 0.00
    micro_rescue_neartrend_rsi_min: int = 48
    micro_rescue_neartrend_mfi_min: int = 35
    micro_rescue_mature_mom_gap_floor: float = -0.0025
    micro_rescue_mature_trend_gap_floor: float = 0.00
    micro_rescue_mature_roc5_min: float = -0.03
    micro_rescue_mature_rsi_min: int = 25
    micro_rescue_mature_mfi_min: int = 20
    micro_rescue_atr_max: float = 0.04
    micro_rescue_vol_max: float = 0.35
    micro_rescue_dd_floor: float = -0.25

    # V20 stateless shadow overlay. Unlike a veto, this changes ONLY the actual exposure
    # for the next execution session while the underlying state machine proceeds to QLD.
    # This prevents a one-day tactical hold from cascading into a different multi-day path.
    # It is deliberately limited to residual states not already handled by V19 near-trend.
    shadow_rescue_mode: str = "off"  # off, score1_recent, mature_mom, combo
    shadow_min_post_crash_days: int = 30
    shadow_recent_max_age: int = 10
    shadow_mature_min_age: int = 60
    shadow_mature_max_age: int = 250
    shadow_score1_trend_gap_floor: float = -0.05
    shadow_score1_mom_gap_floor: float = -0.015
    shadow_score1_roc5_min: float = -0.03
    shadow_score1_roc5_max: float = 0.00
    shadow_score1_roc10_min: float = 0.00
    shadow_score1_roc20_min: float = -0.05
    shadow_score1_rsi_min: int = 38
    shadow_score1_mfi_min: int = 35
    shadow_mature_mom_gap_floor: float = -0.0025
    shadow_mature_trend_gap_floor: float = 0.00
    shadow_mature_roc5_min: float = -0.03
    shadow_mature_roc5_max: float = 0.00
    shadow_mature_rsi_min: int = 25
    shadow_mature_mfi_min: int = 20
    shadow_atr_max: float = 0.04
    shadow_vol_max: float = 0.35
    shadow_dd_floor: float = -0.25

    # V21 stateless forward preview overlay. V20 proved that a one-session execution-only
    # overlay can improve rolling consistency without changing the control state path.
    # V21 adds the opposite direction: while the control remains/steps only to QLD,
    # a narrowly confirmed session may execute TQQQ for one session, then immediately
    # hand control back to the unchanged QLD/CASH/QQQ state machine.
    preview_overlay_mode: str = "off"  # off, qld_score3, qld_score2, stepup, combo
    preview_min_post_crash_days: int = 30
    preview_qld_max_delever_age: int = 40
    preview_qld3_roc5_min: float = -0.01
    preview_qld3_roc10_min: float = 0.00
    preview_qld3_roc20_min: float = 0.00
    preview_qld3_rsi_min: int = 45
    preview_qld3_mfi_min: int = 35
    preview_qld2_trend_gap_floor: float = -0.01
    preview_qld2_mom_gap_floor: float = -0.005
    preview_qld2_roc5_min: float = -0.01
    preview_qld2_roc10_min: float = 0.00
    preview_qld2_roc20_min: float = 0.00
    preview_qld2_rsi_min: int = 48
    preview_qld2_mfi_min: int = 40
    preview_stepup_score_min: int = 2
    preview_stepup_roc5_min: float = 0.02
    preview_stepup_roc10_min: float = 0.02
    preview_stepup_roc20_min: float = 0.00
    preview_stepup_rsi_min: int = 50
    preview_stepup_mfi_min: int = 40
    preview_atr_max: float = 0.035
    preview_vol_max: float = 0.325
    preview_dd_floor: float = -0.25

    # Allow a strong bull impulse to restore 3x even while the base vol score is temporarily weak.
    bull_override_mode: str = "off"  # off, trend_roc, ema_roc
    bull_roc20_min: float = 0.05
    bull_rsi_min: int = 50

    # Very narrow TQQQ trailing-drawdown guard for rare windows where strategy MDD exceeds B&H.
    dd_guard_mode: str = "off"  # off, dd_only, dd_macd, dd_ema
    dd_guard_days: int = 20
    dd_guard_threshold: float = -0.10

    # V4: post-crash explosive rebound bypass. Unlike ordinary Fast, this can bypass entry gates
    # for a short window after an actual crash so V-shaped rebounds are not permanently under-owned.
    rebound_bypass_mode: str = "off"  # off, roc10, roc20, hybrid
    rebound_window_days: int = 42
    rebound_roc10_min: float = 0.10
    rebound_roc20_min: float = 0.18
    rebound_rsi_min: int = 50
    rebound_vol_max: float = 1.00

    # V4: medium-trend bridge for choppy bull / post-correction regimes.
    # It can hold 3x through a one-score wobble, or re-enter using a medium trend confirmation.
    bridge_mode: str = "off"  # off, hold, qld, gated3, direct3, hold_gated3, hold_direct3
    bridge_roc60_min: float = 0.00
    bridge_rsi_min: int = 50
    bridge_dd_floor: float = -0.15
    bridge_vol_max: float = 0.40

    # Search meta (complexity penalty)
    filter_count: int = 0


ORIGINAL_CONFIG = Config(name="ROBUST_ORIGINAL", family="baseline")


# -----------------------------
# CLI
# -----------------------------

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="ROBUST tighter-strategy V21 V20-anchor + bidirectional stateless preview-overlay rolling laboratory",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--workers", type=int, default=12, help="Parallel worker processes")
    p.add_argument("--trials", type=int, default=240000, help="Total candidate count (V20 80.86%% anchor + V21 stateless preview-overlay refinement search), excluding seeded baselines")
    p.add_argument("--finalists", type=int, default=180, help="Candidates receiving exact every-start rolling MDD analysis")
    p.add_argument("--detail-top", type=int, default=20, help="Top finalists written to detailed rolling-window Excel sheet")
    p.add_argument("--audit-top", type=int, default=8, help="Top finalists included in crash-recovery and TQQQ re-promotion audit sheets")
    p.add_argument("--seed", type=int, default=260910, help="Random seed")
    p.add_argument("--data-dir", default=None, help="Folder containing QQQ.csv, QLD.csv, TQQQ.csv")
    p.add_argument("--period", default="max", help="yfinance period if --data-dir is omitted")
    p.add_argument("--output-dir", default="robust_tight_lab_results", help="Output directory")
    p.add_argument("--cost-bps", type=float, default=5.0, help="Estimated transaction cost per trade side in basis points")
    p.add_argument("--self-test", action="store_true", help="Use synthetic data; useful for verifying the script")
    p.add_argument("--top-search", type=int, default=2200, help="Keep this many search-stage candidates before de-dup/finalist selection")
    p.add_argument("--chunk-size", type=int, default=80, help="Configs processed per worker task")
    p.add_argument("--focus-share", type=float, default=0.999, help="Share of trials concentrated on the V21 bidirectional preview-overlay neighborhood")
    p.add_argument("--save-csv", action="store_true", help="Also save top result tables as CSV")
    return p.parse_args()


# -----------------------------
# Data loading
# -----------------------------

def _normalize_ohlcv(raw: pd.DataFrame) -> pd.DataFrame:
    if isinstance(raw.columns, pd.MultiIndex):
        # yfinance may return (Price, Ticker) or (Ticker, Price)
        lvl0 = [str(x).lower() for x in raw.columns.get_level_values(0)]
        if any(x in {"open", "high", "low", "close", "volume"} for x in lvl0):
            raw = raw.copy()
            raw.columns = raw.columns.get_level_values(0)
        else:
            raw = raw.copy()
            raw.columns = raw.columns.get_level_values(-1)
    cols = {str(c).strip().lower(): c for c in raw.columns}
    if "open" not in cols or "close" not in cols:
        raise RuntimeError(f"Open/Close columns missing. Columns={list(raw.columns)}")
    idx = pd.to_datetime(raw.index).tz_localize(None)
    out = pd.DataFrame(index=idx)
    for c in ["Open", "High", "Low", "Close", "Volume"]:
        src = cols.get(c.lower())
        if src is not None:
            out[c] = pd.to_numeric(raw[src], errors="coerce")
    if "High" not in out:
        out["High"] = out["Close"]
    if "Low" not in out:
        out["Low"] = out["Close"]
    if "Volume" not in out:
        out["Volume"] = 0.0
    return out.dropna(subset=["Open", "Close"]).sort_index()


def _read_csv(path: Path) -> pd.DataFrame:
    raw = pd.read_csv(path)
    dcols = [c for c in raw.columns if str(c).strip().lower() in {"date", "datetime", "timestamp"}]
    if not dcols:
        raise RuntimeError(f"No Date/Datetime column in {path}")
    dc = dcols[0]
    raw[dc] = pd.to_datetime(raw[dc])
    raw = raw.set_index(dc)
    return _normalize_ohlcv(raw)


def make_synthetic_data(n: int = 4300, seed: int = 777) -> Dict[str, pd.DataFrame]:
    rng = np.random.default_rng(seed)
    dates = pd.bdate_range("2010-03-01", periods=n)
    # Regime-switching QQQ-like process with several synthetic crashes/rebounds.
    mu = np.full(n, 0.00045)
    sig = np.full(n, 0.012)
    for a, b, drift, vol in [
        (450, 570, -0.0028, 0.027),
        (1150, 1240, -0.0035, 0.030),
        (2450, 2510, -0.0060, 0.045),
        (3000, 3220, -0.0012, 0.023),
        (3800, 3890, -0.0030, 0.030),
    ]:
        if a < n:
            mu[a:min(b, n)] = drift
            sig[a:min(b, n)] = vol
    r = mu + sig * rng.standard_normal(n)
    # Add rebound bursts after crash blocks.
    for a, b in [(570, 630), (1240, 1310), (2510, 2580), (3220, 3300), (3890, 3960)]:
        if a < n:
            r[a:min(b, n)] += 0.0018
    qclose = 100 * np.exp(np.cumsum(r))
    # Overnight/open noise.
    overnight = 0.0025 * rng.standard_normal(n)
    qopen = qclose / np.exp(0.35 * r + overnight)
    qhigh = np.maximum(qopen, qclose) * (1 + np.abs(0.006 * rng.standard_normal(n)))
    qlow = np.minimum(qopen, qclose) * (1 - np.abs(0.006 * rng.standard_normal(n)))
    vol = rng.lognormal(mean=16.0, sigma=0.4, size=n)

    def leveraged(mult: float, decay: float) -> Tuple[np.ndarray, np.ndarray, np.ndarray, np.ndarray]:
        lr = mult * r - decay - 0.5 * (mult * sig) ** 2 / TRADING_DAYS
        c = 50 * np.exp(np.cumsum(lr))
        o = c / np.exp(0.35 * lr + 0.004 * rng.standard_normal(n))
        h = np.maximum(o, c) * (1 + np.abs(0.009 * mult * rng.standard_normal(n)))
        l = np.minimum(o, c) * (1 - np.abs(0.009 * mult * rng.standard_normal(n)))
        return o, h, l, c

    qld_o, qld_h, qld_l, qld_c = leveraged(2.0, 0.00010)
    t_o, t_h, t_l, t_c = leveraged(3.0, 0.00018)

    def frame(o, h, l, c, vmult):
        return pd.DataFrame(
            {"Open": o, "High": h, "Low": l, "Close": c, "Volume": vol * vmult},
            index=dates,
        )

    return {"QQQ": frame(qopen, qhigh, qlow, qclose, 1.0), "QLD": frame(qld_o, qld_h, qld_l, qld_c, 0.25), "TQQQ": frame(t_o, t_h, t_l, t_c, 0.45)}


def load_data(data_dir: Optional[str], period: str, self_test: bool) -> Dict[str, pd.DataFrame]:
    if self_test:
        print("[data] SELF-TEST synthetic data")
        return make_synthetic_data()

    data: Dict[str, pd.DataFrame] = {}
    if data_dir:
        folder = Path(data_dir)
        for ticker in ("QQQ", "QLD", "TQQQ"):
            p = folder / f"{ticker}.csv"
            if not p.exists():
                raise FileNotFoundError(f"Missing CSV: {p}")
            data[ticker] = _read_csv(p)
    else:
        try:
            import yfinance as yf
        except ImportError as exc:
            raise RuntimeError("yfinance is required. Run: pip install yfinance") from exc
        print(f"[data] Downloading QQQ / QLD / TQQQ, period={period} ...")
        for ticker in ("QQQ", "QLD", "TQQQ"):
            raw = yf.download(
                ticker,
                period=period,
                interval="1d",
                auto_adjust=True,
                actions=False,
                progress=False,
                threads=False,
            )
            if raw is None or raw.empty:
                raise RuntimeError(f"Download failed for {ticker}")
            data[ticker] = _normalize_ohlcv(raw)

    idx = data["QQQ"].index.intersection(data["QLD"].index).intersection(data["TQQQ"].index)
    idx = idx.sort_values()
    if len(idx) < 2800:
        print(f"[warning] Only {len(idx)} common rows. 10Y rolling windows may be unavailable.")
    if len(idx) < 600:
        raise RuntimeError(f"Not enough common trading days: {len(idx)}")
    return {k: v.reindex(idx).dropna(subset=["Open", "Close"]) for k, v in data.items()}


# -----------------------------
# Indicators
# -----------------------------

def _ema(s: pd.Series, span: int) -> pd.Series:
    return s.ewm(span=span, adjust=False, min_periods=span).mean()


def _rsi(s: pd.Series, n: int = 14) -> pd.Series:
    d = s.diff()
    up = d.clip(lower=0.0)
    dn = -d.clip(upper=0.0)
    au = up.ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    ad = dn.ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    rs = au / ad.replace(0, np.nan)
    out = 100 - 100 / (1 + rs)
    return out.fillna(50.0)


def _macd_hist(s: pd.Series, fast: int = 12, slow: int = 26, signal: int = 9) -> pd.Series:
    m = _ema(s, fast) - _ema(s, slow)
    sig = m.ewm(span=signal, adjust=False, min_periods=signal).mean()
    return m - sig


def _true_range(df: pd.DataFrame) -> pd.Series:
    pc = df["Close"].shift(1)
    a = df["High"] - df["Low"]
    b = (df["High"] - pc).abs()
    c = (df["Low"] - pc).abs()
    return pd.concat([a, b, c], axis=1).max(axis=1)


def _atr_pct(df: pd.DataFrame, n: int = 14) -> pd.Series:
    atr = _true_range(df).ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    return atr / df["Close"]


def _adx_di(df: pd.DataFrame, n: int = 14) -> Tuple[pd.Series, pd.Series, pd.Series]:
    h, l = df["High"], df["Low"]
    up = h.diff()
    dn = -l.diff()
    plus_dm = pd.Series(np.where((up > dn) & (up > 0), up, 0.0), index=df.index)
    minus_dm = pd.Series(np.where((dn > up) & (dn > 0), dn, 0.0), index=df.index)
    atr = _true_range(df).ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    plus_di = 100 * plus_dm.ewm(alpha=1 / n, adjust=False, min_periods=n).mean() / atr.replace(0, np.nan)
    minus_di = 100 * minus_dm.ewm(alpha=1 / n, adjust=False, min_periods=n).mean() / atr.replace(0, np.nan)
    dx = 100 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    adx = dx.ewm(alpha=1 / n, adjust=False, min_periods=n).mean()
    return adx.fillna(0), plus_di.fillna(0), minus_di.fillna(0)


def _mfi(df: pd.DataFrame, n: int = 14) -> pd.Series:
    tp = (df["High"] + df["Low"] + df["Close"]) / 3.0
    mf = tp * df["Volume"].fillna(0.0)
    pos = pd.Series(np.where(tp.diff() > 0, mf, 0.0), index=df.index)
    neg = pd.Series(np.where(tp.diff() < 0, mf, 0.0), index=df.index)
    ps = pos.rolling(n, min_periods=n).sum()
    ns = neg.rolling(n, min_periods=n).sum()
    ratio = ps / ns.replace(0, np.nan)
    return (100 - 100 / (1 + ratio)).fillna(50.0)


def _obv(df: pd.DataFrame) -> pd.Series:
    sign = np.sign(df["Close"].diff()).fillna(0.0)
    return (sign * df["Volume"].fillna(0.0)).cumsum()


def add_asset_indicators(df: pd.DataFrame, prefix: str, feat: Dict[str, np.ndarray]) -> None:
    c = df["Close"]
    h = df["High"]
    l = df["Low"]

    for n in [10, 20, 50, 80, 100, 120, 140, 150, 165, 180, 200, 220]:
        feat[f"{prefix}_SMA{n}"] = c.rolling(n, min_periods=n).mean().to_numpy(dtype=float)
    for n in [5, 10, 15, 20, 30, 50]:
        feat[f"{prefix}_EMA{n}"] = _ema(c, n).to_numpy(dtype=float)
    for n in [10, 20, 40, 60]:
        feat[f"{prefix}_VOL{n}"] = (
            c.pct_change().rolling(n, min_periods=n).std(ddof=0) * math.sqrt(TRADING_DAYS)
        ).to_numpy(dtype=float)
    for n in [5, 10, 15, 20, 40, 60, 90, 120, 150]:
        feat[f"{prefix}_ROC{n}"] = c.pct_change(n).to_numpy(dtype=float)

    feat[f"{prefix}_RSI14"] = _rsi(c, 14).to_numpy(dtype=float)
    feat[f"{prefix}_MACD_HIST"] = _macd_hist(c).to_numpy(dtype=float)

    low14 = l.rolling(14, min_periods=14).min()
    high14 = h.rolling(14, min_periods=14).max()
    stoch_k = 100 * (c - low14) / (high14 - low14).replace(0, np.nan)
    stoch_d = stoch_k.rolling(3, min_periods=3).mean()
    feat[f"{prefix}_STOCH_K"] = stoch_k.to_numpy(dtype=float)
    feat[f"{prefix}_STOCH_D"] = stoch_d.to_numpy(dtype=float)

    mid = c.rolling(20, min_periods=20).mean()
    sd = c.rolling(20, min_periods=20).std(ddof=0)
    upper = mid + 2 * sd
    lower = mid - 2 * sd
    feat[f"{prefix}_BOLL_MID"] = mid.to_numpy(dtype=float)
    feat[f"{prefix}_BOLL_UPPER"] = upper.to_numpy(dtype=float)
    feat[f"{prefix}_BOLL_LOWER"] = lower.to_numpy(dtype=float)
    feat[f"{prefix}_BOLL_POS"] = ((c - lower) / (upper - lower).replace(0, np.nan)).to_numpy(dtype=float)
    feat[f"{prefix}_BOLL_WIDTH"] = ((upper - lower) / mid.replace(0, np.nan)).to_numpy(dtype=float)

    tp = (h + l + c) / 3.0
    tpma = tp.rolling(20, min_periods=20).mean()
    md = (tp - tpma).abs().rolling(20, min_periods=20).mean()
    feat[f"{prefix}_CCI20"] = ((tp - tpma) / (0.015 * md.replace(0, np.nan))).to_numpy(dtype=float)

    adx, pdi, mdi = _adx_di(df, 14)
    feat[f"{prefix}_ADX14"] = adx.to_numpy(dtype=float)
    feat[f"{prefix}_PDI14"] = pdi.to_numpy(dtype=float)
    feat[f"{prefix}_MDI14"] = mdi.to_numpy(dtype=float)

    feat[f"{prefix}_ATRPCT14"] = _atr_pct(df, 14).to_numpy(dtype=float)
    feat[f"{prefix}_MFI14"] = _mfi(df, 14).to_numpy(dtype=float)

    wr = -100 * (high14 - c) / (high14 - low14).replace(0, np.nan)
    feat[f"{prefix}_WILLR14"] = wr.to_numpy(dtype=float)

    obv = _obv(df)
    feat[f"{prefix}_OBV"] = obv.to_numpy(dtype=float)
    feat[f"{prefix}_OBV_MA20"] = obv.rolling(20, min_periods=20).mean().to_numpy(dtype=float)

    d_hi = h.rolling(20, min_periods=20).max()
    d_lo = l.rolling(20, min_periods=20).min()
    feat[f"{prefix}_DONCHIAN_POS20"] = ((c - d_lo) / (d_hi - d_lo).replace(0, np.nan)).to_numpy(dtype=float)

    for n in [20, 40, 60, 90]:
        hin = c.rolling(n, min_periods=max(10, n // 2)).max()
        feat[f"{prefix}_DD{n}"] = (c / hin - 1.0).to_numpy(dtype=float)

    hi252 = c.rolling(252, min_periods=126).max()
    feat[f"{prefix}_DD252"] = (c / hi252 - 1.0).to_numpy(dtype=float)


def build_features(data: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    idx = data["QQQ"].index
    feat: Dict[str, Any] = {"dates": idx.to_numpy(dtype="datetime64[ns]")}
    for ticker in ("QQQ", "QLD", "TQQQ"):
        df = data[ticker]
        feat[f"{ticker}_OPEN"] = df["Open"].to_numpy(dtype=float)
        feat[f"{ticker}_CLOSE"] = df["Close"].to_numpy(dtype=float)
        feat[f"{ticker}_HIGH"] = df["High"].to_numpy(dtype=float)
        feat[f"{ticker}_LOW"] = df["Low"].to_numpy(dtype=float)
        feat[f"{ticker}_VOLUME"] = df["Volume"].to_numpy(dtype=float)
        add_asset_indicators(df, ticker, feat)

    # Fixed v15-inspired TQQQ risk score, calculated on close.
    c_t = feat["TQQQ_CLOSE"]
    c_q = feat["QQQ_CLOSE"]
    risk = np.zeros(len(idx), dtype=np.int16)
    risk += (c_t < feat["TQQQ_SMA20"]).astype(np.int16) * 1
    risk += (c_t < feat["TQQQ_SMA50"]).astype(np.int16) * 2
    risk += (c_q < feat["QQQ_SMA50"]).astype(np.int16) * 1
    risk += (c_q < feat["QQQ_SMA200"]).astype(np.int16) * 3
    risk += (feat["TQQQ_MACD_HIST"] < 0).astype(np.int16) * 1
    risk += (feat["TQQQ_RSI14"] < 50).astype(np.int16) * 1
    risk += (feat["TQQQ_STOCH_K"] < feat["TQQQ_STOCH_D"]).astype(np.int16) * 1
    risk += (feat["TQQQ_CCI20"] < 0).astype(np.int16) * 1
    risk += (c_t < feat["TQQQ_BOLL_MID"]).astype(np.int16) * 1
    risk += (feat["TQQQ_OBV"] < feat["TQQQ_OBV_MA20"]).astype(np.int16) * 1
    t_ret1 = np.r_[np.nan, c_t[1:] / c_t[:-1] - 1.0]
    vol20 = pd.Series(data["TQQQ"]["Volume"], index=idx).rolling(20, min_periods=20).mean().to_numpy()
    risk += ((t_ret1 < 0) & (feat["TQQQ_VOLUME"] > vol20)).astype(np.int16) * 1
    feat["V15_RISK_SCORE"] = risk

    # Open-to-open returns. Position chosen at today's open earns today's open -> next open return.
    for ticker in ("QQQ", "QLD", "TQQQ"):
        o = feat[f"{ticker}_OPEN"]
        r = np.zeros(len(o), dtype=float)
        r[:-1] = o[1:] / o[:-1] - 1.0
        r[-1] = 0.0
        feat[f"{ticker}_O2O_RET"] = r

    return feat


# -----------------------------
# Candidate generation
# -----------------------------

def _filter_count(cfg: Dict[str, Any]) -> int:
    n = 0
    for k in ["require_qqq_ma50", "require_qqq_ma200", "require_tqqq_ma20", "require_tqqq_ma50"]:
        n += int(bool(cfg.get(k, False)))
    for k in ["macd_gate", "rsi_asset", "stoch_asset", "cci_asset", "boll_asset", "adx_asset", "mfi_asset", "willr_asset", "obv_asset", "donchian_asset", "atr_asset", "roc_asset"]:
        n += int(cfg.get(k, "none") != "none")
    n += int(cfg.get("defense_mode", "off") != "off")
    n += int(cfg.get("tqqq_exit_mode", "off") != "off")
    n += int(cfg.get("recovery_mode", "off") != "off")
    n += int(cfg.get("bull_override_mode", "off") != "off")
    n += int(cfg.get("dd_guard_mode", "off") != "off")
    n += int(cfg.get("rebound_bypass_mode", "off") != "off")
    n += int(cfg.get("bridge_mode", "off") != "off")
    n += int(cfg.get("gate_logic", "all") != "all")
    n += int(cfg.get("adaptive_down_mode", "off") != "off")
    n += int(cfg.get("snapback_mode", "off") != "off")
    n += int(cfg.get("crash_release_mode", "off") != "off")
    n += int(cfg.get("staged_reentry_mode", "off") != "off")
    n += int(cfg.get("explosive_rebound_mode", "off") != "off")
    n += int(cfg.get("correction_reentry_mode", "off") != "off")
    n += int(cfg.get("dd_rebound_mode", "off") != "off")
    n += int(cfg.get("trough_recovery_mode", "off") != "off")
    n += int(cfg.get("repromote_guard_mode", "off") != "off")
    n += int(cfg.get("soft_vol_mode", "off") != "off")
    n += int(cfg.get("impulse_veto_mode", "off") != "off")
    n += int(cfg.get("precision_veto_mode", "off") != "off")
    n += int(cfg.get("reentry_grace_mode", "off") != "off")
    n += int(cfg.get("micro_rescue_mode", "off") != "off")
    n += int(cfg.get("shadow_rescue_mode", "off") != "off")
    n += int(cfg.get("preview_overlay_mode", "off") != "off")
    return n


def seeded_configs() -> List[Config]:
    seeds = [ORIGINAL_CONFIG]
    seeds += [
        Config(name="ROBUST_NO_FAST", family="seed", fast_mode="normal"),
        Config(name="ROBUST_FAST_TO_QLD", family="seed", fast_mode="qld"),
        Config(
            name="ROBUST_V15_ENTRY",
            family="seed",
            require_qqq_ma50=True,
            require_tqqq_ma20=True,
            macd_gate="tqqq",
            rsi_asset="tqqq",
            rsi_min=50,
            filter_count=4,
        ),
        Config(
            name="ROBUST_V15_ENTRY_FAST_QLD",
            family="seed",
            fast_mode="qld",
            require_qqq_ma50=True,
            require_tqqq_ma20=True,
            macd_gate="tqqq",
            filter_count=3,
        ),
        Config(
            name="ROBUST_RISK_CAP2_8",
            family="seed",
            defense_mode="cap2",
            defense_score_threshold=8,
            filter_count=1,
        ),
        Config(
            name="ROBUST_RISK_CAP1_8",
            family="seed",
            defense_mode="cap1",
            defense_score_threshold=8,
            filter_count=1,
        ),
        Config(
            name="ROBUST_TIGHT_COMBO",
            family="seed",
            fast_mode="qld",
            down_confirm=1,
            require_qqq_ma50=True,
            require_tqqq_ma20=True,
            macd_gate="tqqq",
            rsi_asset="tqqq",
            rsi_min=50,
            defense_mode="cap2",
            defense_score_threshold=8,
            filter_count=5,
        ),
    ]
    # V1 real-run winners are preserved as anchors for focused V2 exploration.
    seeds += [
        Config(name="V1_RND_000171", family="v1_winner", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=20, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=10, fast_threshold=0.02, fast_long_mom_floor=-0.12, crash_dd=-0.15, crash_vol_threshold=0.275, down_confirm=2, fast_mode="direct", macd_gate="tqqq", rsi_asset="qqq", rsi_min=50, boll_asset="qqq", boll_rule="mid", willr_asset="tqqq", willr_min=-70, filter_count=4),
        Config(name="V1_RND_000171_TQQQ_ONLY", family="v2_seed", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=20, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=10, fast_threshold=0.02, fast_long_mom_floor=-0.12, crash_dd=-0.15, crash_vol_threshold=0.275, down_confirm=2, fast_mode="conditional", gate_scope="tqqq_only", macd_gate="tqqq", rsi_asset="qqq", rsi_min=50, boll_asset="qqq", boll_rule="mid", willr_asset="tqqq", willr_min=-70, filter_count=4),
        Config(name="V1_RND_004288", family="v1_winner", sma_days=165, mom_days=150, mom_threshold=-0.01, vol_days=20, vol_threshold=0.30, fast_ema_days=20, fast_mom_days=10, fast_threshold=0.03, fast_long_mom_floor=-0.08, crash_dd=-0.15, crash_vol_threshold=0.25, down_confirm=3, fast_mode="direct", cci_asset="tqqq", obv_asset="qqq", filter_count=2),
        Config(name="V1_RND_005848", family="v1_winner", sma_days=180, mom_days=120, mom_threshold=-0.03, vol_days=20, vol_threshold=0.275, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.02, fast_long_mom_floor=-0.12, crash_dd=-0.15, crash_vol_threshold=0.225, down_confirm=3, fast_mode="direct", macd_gate="both", cci_asset="tqqq", obv_asset="qqq", filter_count=3),
    ]
    # V2 elites from the user's 60k run. These are exact anchors for V3 mutation.
    seeds += [
        Config(name="V2_FOCUS_031316", family="v2_elite", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.30, fast_ema_days=20, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.12, crash_dd=-0.23, crash_vol_threshold=0.275, down_confirm=2, fast_mode="conditional", gate_scope="tqqq_only", boll_asset="tqqq", boll_rule="upper_half", mfi_asset="tqqq", mfi_min=45, filter_count=2),
        Config(name="V2_FOCUS_018911", family="v2_elite", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=20, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.15, crash_dd=-0.23, crash_vol_threshold=0.225, down_confirm=2, fast_mode="direct", gate_scope="tqqq_only", boll_asset="qqq", boll_rule="not_overbought", filter_count=1),
        Config(name="V2_FOCUS_025256", family="v2_elite", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=20, vol_threshold=0.35, fast_ema_days=15, fast_mom_days=10, fast_threshold=0.02, fast_long_mom_floor=-0.12, crash_dd=-0.23, crash_vol_threshold=0.225, down_confirm=2, fast_mode="conditional", gate_scope="all_up", require_qqq_ma50=True, cci_asset="tqqq", willr_asset="qqq", willr_min=-80, filter_count=3),
        Config(name="V2_FOCUS_027205", family="v2_elite", sma_days=165, mom_days=120, mom_threshold=0.0, vol_days=20, vol_threshold=0.30, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.12, crash_dd=-0.18, crash_vol_threshold=0.225, down_confirm=3, fast_mode="conditional", gate_scope="tqqq_only", require_qqq_ma50=True, boll_asset="qqq", boll_rule="mid", willr_asset="tqqq", willr_min=-40, donchian_asset="tqqq", donchian_min_pos=0.50, filter_count=4),
        Config(name="V2_FOCUS_031870", family="v2_elite", sma_days=180, mom_days=150, mom_threshold=-0.01, vol_days=20, vol_threshold=0.30, fast_ema_days=15, fast_mom_days=20, fast_threshold=0.02, fast_long_mom_floor=-0.08, crash_dd=-0.15, crash_vol_threshold=0.25, down_confirm=2, fast_mode="direct", gate_scope="all_up", boll_asset="qqq", boll_rule="upper_half", donchian_asset="tqqq", donchian_min_pos=0.50, filter_count=2),
        Config(name="V2_031316_RECOVERY", family="v3_seed", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.30, fast_ema_days=20, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.12, crash_dd=-0.23, crash_vol_threshold=0.275, down_confirm=2, fast_mode="conditional", gate_scope="tqqq_only", boll_asset="tqqq", boll_rule="upper_half", mfi_asset="tqqq", mfi_min=45, recovery_mode="hybrid", recovery_days=63, recovery_mom_floor=-0.30, recovery_roc20_min=0.05, recovery_vol_max=0.65, filter_count=3),
    ]
    # V3 elites from the user's 120k run. Exact anchors for V5 local search.
    seeds += [
        Config(name="V3_FOCUS_071187", family="v3_elite", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.40, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.15, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", mfi_asset="qqq", mfi_min=45, atr_asset="qqq", atr_max_pct=0.05, filter_count=2),
        Config(name="V3_FOCUS_079471", family="v3_elite", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=20, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.20, crash_dd=-0.23, crash_vol_threshold=0.275, down_confirm=2, fast_mode="conditional", gate_scope="tqqq_only", obv_asset="qqq", donchian_asset="qqq", donchian_min_pos=0.50, filter_count=2),
        Config(name="V3_FOCUS_055894", family="v3_elite", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=20, vol_threshold=0.40, fast_ema_days=10, fast_mom_days=10, fast_threshold=0.0, fast_long_mom_floor=-0.30, crash_dd=-0.23, crash_vol_threshold=0.275, down_confirm=2, fast_mode="conditional", gate_scope="tqqq_only", adx_asset="tqqq", adx_min=15, bull_override_mode="trend_roc", bull_roc20_min=0.07, bull_rsi_min=60, filter_count=2),
        Config(name="V3_FOCUS_060529", family="v3_elite", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.30, fast_ema_days=20, fast_mom_days=10, fast_threshold=0.0, fast_long_mom_floor=-0.15, crash_dd=-0.23, crash_vol_threshold=0.275, down_confirm=2, fast_mode="conditional", gate_scope="tqqq_only", rsi_asset="qqq", rsi_min=50, boll_asset="qqq", boll_rule="mid", filter_count=2),
        Config(name="V3_FOCUS_099955", family="v3_elite", sma_days=165, mom_days=120, mom_threshold=-0.01, vol_days=20, vol_threshold=0.375, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.02, fast_long_mom_floor=-0.15, crash_dd=-0.23, crash_vol_threshold=0.275, down_confirm=3, fast_mode="direct", gate_scope="all_up", filter_count=0),
        # Deliberate V4 probes around the V3 winner: bypass explosive rebound regimes and bridge choppy bull regimes.
        Config(name="V4_071187_REBOUND", family="v4_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.40, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.15, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", mfi_asset="qqq", mfi_min=45, atr_asset="qqq", atr_max_pct=0.05, rebound_bypass_mode="hybrid", rebound_window_days=63, rebound_roc10_min=0.08, rebound_roc20_min=0.15, rebound_rsi_min=50, rebound_vol_max=1.20, filter_count=3),
        Config(name="V4_071187_BRIDGE", family="v4_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.40, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.15, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", mfi_asset="qqq", mfi_min=45, atr_asset="qqq", atr_max_pct=0.05, bridge_mode="hold_gated3", bridge_roc60_min=0.0, bridge_rsi_min=50, bridge_dd_floor=-0.18, bridge_vol_max=0.45, filter_count=3),
        Config(name="V4_071187_BOTH", family="v4_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.40, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.015, fast_long_mom_floor=-0.15, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", mfi_asset="qqq", mfi_min=45, atr_asset="qqq", atr_max_pct=0.05, rebound_bypass_mode="hybrid", rebound_window_days=63, rebound_roc10_min=0.08, rebound_roc20_min=0.15, rebound_rsi_min=50, rebound_vol_max=1.20, bridge_mode="hold_gated3", bridge_roc60_min=0.0, bridge_rsi_min=50, bridge_dd_floor=-0.18, bridge_vol_max=0.45, filter_count=4),
    ]
    # V4 winner from the user's 240k run, plus targeted probes for the two remaining weak regimes.
    seeds += [
        Config(name="V4_FOCUS_204252", family="v4_elite", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.35, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, filter_count=3),
        Config(name="V5_FOCUS_366726", family="v5_elite", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.25, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, bull_override_mode="ema_roc", bull_roc20_min=0.03, bull_rsi_min=55, filter_count=4),
        Config(name="V5_FOCUS_267771", family="v5_elite", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.375, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.35, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="direct", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=35, atr_asset="qqq", atr_max_pct=0.08, recovery_mode="relaxed_fast", recovery_days=84, recovery_mom_floor=-0.45, recovery_roc20_min=0.03, recovery_vol_max=0.65, filter_count=4),
        Config(name="V6_366726_2OF3", family="v6_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.25, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", gate_logic="rma_2of3", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, bull_override_mode="ema_roc", bull_roc20_min=0.03, bull_rsi_min=55, filter_count=5),
        Config(name="V6_366726_ADAPT", family="v6_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.25, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, adaptive_down_mode="tqqq_score", soft_down_confirm=4, hard_down_confirm=1, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, bull_override_mode="ema_roc", bull_roc20_min=0.03, bull_rsi_min=55, filter_count=5),
        Config(name="V6_366726_SNAP", family="v6_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.25, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=20, snapback_roc10_min=0.01, snapback_rsi_min=50, snapback_vol_max=0.50, bull_override_mode="ema_roc", bull_roc20_min=0.03, bull_rsi_min=55, filter_count=5),
        Config(name="V5_204252_FAST5", family="v5_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=5, fast_mom_days=5, fast_threshold=0.0, fast_long_mom_floor=-0.45, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, filter_count=3),
        Config(name="V5_204252_REC_FAST", family="v5_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=5, fast_mom_days=5, fast_threshold=0.0, fast_long_mom_floor=-0.45, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, recovery_mode="relaxed_fast", recovery_days=42, recovery_mom_floor=-0.50, recovery_roc20_min=0.03, recovery_vol_max=1.20, filter_count=4),
        Config(name="V5_204252_REC_HYBRID", family="v5_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=5, fast_mom_days=10, fast_threshold=-0.01, fast_long_mom_floor=-0.45, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.10, recovery_mode="hybrid", recovery_days=63, recovery_mom_floor=-0.50, recovery_roc20_min=0.03, recovery_vol_max=1.20, filter_count=4),
        Config(name="V5_204252_TQQQ_ONLY", family="v5_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.35, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="tqqq_only", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, filter_count=3),
        Config(name="V5_204252_FAST_ONLY", family="v5_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=10, fast_mom_days=10, fast_threshold=0.0, fast_long_mom_floor=-0.40, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="fast_only", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, filter_count=3),
        Config(name="V5_204252_CRASH12", family="v5_probe", sma_days=180, mom_days=120, mom_threshold=0.0, vol_days=40, vol_threshold=0.35, fast_ema_days=10, fast_mom_days=10, fast_threshold=0.0, fast_long_mom_floor=-0.45, crash_dd=-0.12, crash_vol_threshold=0.275, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.10, recovery_mode="relaxed_fast", recovery_days=42, recovery_mom_floor=-0.50, recovery_vol_max=1.20, filter_count=4),
    ]
    # V6 winner from the user's 480k run: exact anchor for V7.
    seeds += [
        Config(name="V6_FOCUS_107803", family="v6_elite", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", gate_logic="all", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.15, filter_count=4),
        Config(name="V7_107803_SNAP_DD25", family="v7_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=20, snapback_roc10_min=0.00, snapback_rsi_min=45, snapback_vol_max=0.50, snapback_ema_days=10, snapback_dd_floor=-0.25, filter_count=4),
        Config(name="V7_107803_CR_Q", family="v7_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=15, snapback_roc10_min=0.00, snapback_rsi_min=45, snapback_vol_max=0.50, snapback_ema_days=10, snapback_dd_floor=-0.30, crash_release_mode="qld", crash_release_wait=5, crash_release_roc5_min=0.06, crash_release_roc10_min=0.10, crash_release_rsi_min=45, crash_release_mfi_min=40, crash_release_vol_max=1.20, crash_release_dd_floor=-0.45, filter_count=5),
        # Exact V7 480k winner from the uploaded dashboard. Preserved as the second anchor.
        Config(name="V7_FOCUS_326611", family="v7_elite", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.3375, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", gate_logic="all", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.10, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.25, crash_release_mode="gated3", crash_release_wait=10, crash_release_roc5_min=0.04, crash_release_roc10_min=0.05, crash_release_rsi_min=45, crash_release_mfi_min=50, crash_release_vol_max=0.65, crash_release_dd_floor=-0.60, filter_count=5),
        # V8 hand-built probes: base ATR stays 8%; only recovery modules get wider ATR allowances.
        Config(name="V8_STAGE_BALANCED", family="v8_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="off", staged_reentry_mode="qld_confirm3", staged_window_days=15, staged_ema_days=10, staged_roc10_min=0.00, staged_rsi_min=45, staged_mfi_min=40, staged_atr_max=0.12, staged_vol_max=0.60, staged_dd_floor=-0.25, staged_promote_wait=2, staged_promote_roc5_min=0.02, staged_promote_roc10_min=0.04, staged_promote_rsi_min=50, staged_promote_mfi_min=45, staged_promote_atr_max=0.10, filter_count=4),
        Config(name="V8_EXPLOSIVE", family="v8_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.15, explosive_rebound_mode="direct3", explosive_window_days=20, explosive_min_crash_days=2, explosive_roc5_min=0.08, explosive_roc10_min=0.15, explosive_rsi_min=50, explosive_mfi_min=40, explosive_atr_max=0.18, explosive_vol_max=1.50, explosive_hold_days=5, explosive_hold_roc5_min=0.00, filter_count=5),
        Config(name="V8_COMBO_BALANCED", family="v8_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="off", staged_reentry_mode="qld_confirm3", staged_window_days=15, staged_ema_days=10, staged_roc10_min=0.00, staged_rsi_min=45, staged_mfi_min=40, staged_atr_max=0.12, staged_vol_max=0.60, staged_dd_floor=-0.25, staged_promote_wait=2, staged_promote_roc5_min=0.02, staged_promote_roc10_min=0.04, staged_promote_rsi_min=50, staged_promote_mfi_min=45, staged_promote_atr_max=0.10, explosive_rebound_mode="direct3", explosive_window_days=20, explosive_min_crash_days=2, explosive_roc5_min=0.08, explosive_roc10_min=0.15, explosive_rsi_min=50, explosive_mfi_min=40, explosive_atr_max=0.18, explosive_vol_max=1.50, explosive_hold_days=5, explosive_hold_roc5_min=0.00, filter_count=5),
        Config(name="V8_COMBO_PLUS_SNAP", family="v8_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.325, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.15, staged_reentry_mode="qld_confirm3", staged_window_days=15, staged_ema_days=10, staged_roc10_min=0.00, staged_rsi_min=45, staged_mfi_min=40, staged_atr_max=0.12, staged_vol_max=0.60, staged_dd_floor=-0.25, staged_promote_wait=2, staged_promote_roc5_min=0.02, staged_promote_roc10_min=0.04, staged_promote_rsi_min=50, staged_promote_mfi_min=45, staged_promote_atr_max=0.10, explosive_rebound_mode="direct3", explosive_window_days=20, explosive_min_crash_days=2, explosive_roc5_min=0.08, explosive_roc10_min=0.15, explosive_rsi_min=50, explosive_mfi_min=40, explosive_atr_max=0.18, explosive_vol_max=1.50, explosive_hold_days=5, explosive_hold_roc5_min=0.00, filter_count=6),
    ]
    # Exact V8 480k anchors from the uploaded V8 dashboard.
    seeds += [
        Config(name="V8_FOCUS_211857", family="v8_elite_stable", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.3125, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", gate_logic="all", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.15, staged_reentry_mode="qld_confirm3", staged_window_days=7, staged_ema_days=20, staged_roc10_min=0.01, staged_rsi_min=50, staged_mfi_min=45, staged_atr_max=0.08, staged_vol_max=0.60, staged_dd_floor=-0.25, staged_promote_wait=5, staged_promote_roc5_min=0.02, staged_promote_roc10_min=0.05, staged_promote_rsi_min=45, staged_promote_mfi_min=35, staged_promote_atr_max=0.10, explosive_rebound_mode="direct3", explosive_window_days=20, explosive_min_crash_days=2, explosive_roc5_min=0.10, explosive_roc10_min=0.18, explosive_rsi_min=50, explosive_mfi_min=40, explosive_atr_max=0.12, explosive_vol_max=1.20, explosive_hold_days=5, explosive_hold_roc5_min=-0.02, filter_count=6),
        Config(name="V8_FOCUS_291951", family="v8_elite_aggressive", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.3375, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", gate_logic="all", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.30, staged_reentry_mode="qld_gate3", staged_window_days=7, staged_ema_days=15, staged_roc10_min=0.01, staged_rsi_min=50, staged_mfi_min=40, staged_atr_max=0.12, staged_vol_max=0.80, staged_dd_floor=-0.30, staged_promote_wait=3, staged_promote_roc5_min=0.01, staged_promote_roc10_min=0.04, staged_promote_rsi_min=50, staged_promote_mfi_min=50, staged_promote_atr_max=0.12, explosive_rebound_mode="direct3", explosive_window_days=20, explosive_min_crash_days=2, explosive_roc5_min=0.12, explosive_roc10_min=0.20, explosive_rsi_min=50, explosive_mfi_min=40, explosive_atr_max=0.18, explosive_vol_max=0.80, explosive_hold_days=5, explosive_hold_roc5_min=-0.02, filter_count=6),
    ]

    # V9 deterministic probes ensure the precision-vol and new recovery paths are
    # represented even if random finalist selection is unlucky.
    for j, vt in enumerate([0.3000,0.3025,0.3050,0.3075,0.3100,0.3125,0.3150,0.3175,0.3200,0.3225,0.3250,0.3275,0.3300,0.3325,0.3350,0.3375,0.3400,0.3425,0.3450]):
        seeds.append(Config(name=f"V9_VOL_{int(round(vt*10000)):04d}", family="v9_vol_grid", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=vt, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", gate_logic="all", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.15, filter_count=4))
    seeds += [
        Config(name="V9_CORR_BALANCED", family="v9_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.3125, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="off", correction_reentry_mode="gated3", correction_window_days=15, correction_ema_days=10, correction_roc5_min=0.00, correction_roc10_min=-0.01, correction_rsi_min=45, correction_mfi_min=35, correction_atr_max=0.12, correction_vol_max=0.45, correction_dd_floor=-0.20, filter_count=4),
        Config(name="V9_DD_REBOUND", family="v9_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.3125, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.15, dd_rebound_mode="direct3", dd_rebound_window_days=42, dd_rebound_min=0.06, dd_rebound_roc5_min=0.02, dd_rebound_rsi_min=45, dd_rebound_mfi_min=35, dd_rebound_atr_max=0.25, dd_rebound_vol_max=1.50, dd_rebound_hold_days=5, dd_rebound_hold_roc5_min=-0.03, filter_count=5),
        Config(name="V9_COMBO_BALANCED", family="v9_probe", sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.3150, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30, crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2, fast_mode="conditional", gate_scope="all_up", rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08, snapback_mode="off", correction_reentry_mode="gated3", correction_window_days=15, correction_ema_days=10, correction_roc5_min=0.00, correction_roc10_min=-0.01, correction_rsi_min=45, correction_mfi_min=35, correction_atr_max=0.12, correction_vol_max=0.45, correction_dd_floor=-0.20, dd_rebound_mode="direct3", dd_rebound_window_days=42, dd_rebound_min=0.06, dd_rebound_roc5_min=0.02, dd_rebound_rsi_min=45, dd_rebound_mfi_min=35, dd_rebound_atr_max=0.25, dd_rebound_vol_max=1.50, dd_rebound_hold_days=5, dd_rebound_hold_roc5_min=-0.03, filter_count=5),
    ]


    # Exact V9 frontier anchors from the uploaded 480k V9 dashboard.
    # These are preserved unchanged so V11 can prove real Pareto improvements against prior frontiers.
    seeds += [
        Config(name='V9_FOCUS_128967', family='v9_frontier_win79', sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.34, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.3, crash_dd=-0.15, crash_vol_threshold=0.3, down_confirm=2, adaptive_down_mode='off', soft_down_confirm=3, hard_down_confirm=1, fast_mode='conditional', gate_scope='all_up', gate_logic='all', rsi_asset='qqq', rsi_min=45, mfi_asset='qqq', mfi_min=40, atr_asset='qqq', atr_max_pct=0.08, snapback_mode='gated3', snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.35, snapback_ema_days=20, snapback_dd_floor=-0.175, staged_dd_floor=-0.25, explosive_window_days=20, correction_reentry_mode='qld', correction_window_days=10, correction_ema_days=10, correction_roc5_min=0.01, correction_roc10_min=-0.02, correction_rsi_min=42, correction_mfi_min=35, correction_atr_max=0.12, correction_vol_max=0.5, correction_dd_floor=-0.15, dd_rebound_mode='direct3', dd_rebound_window_days=42, dd_rebound_min=0.04, dd_rebound_roc5_min=0.04, dd_rebound_rsi_min=50, dd_rebound_mfi_min=40, dd_rebound_atr_max=0.25, dd_rebound_vol_max=1.5, dd_rebound_hold_days=5, dd_rebound_hold_roc5_min=-0.03, filter_count=6),
        Config(name='V9_FOCUS_147025', family='v9_frontier_balanced', sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.335, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.3, crash_dd=-0.15, crash_vol_threshold=0.3, down_confirm=2, adaptive_down_mode='off', soft_down_confirm=3, hard_down_confirm=1, fast_mode='conditional', gate_scope='all_up', gate_logic='all', rsi_asset='qqq', rsi_min=45, mfi_asset='qqq', mfi_min=40, atr_asset='qqq', atr_max_pct=0.08, snapback_mode='off', snapback_days=10, snapback_roc10_min=0.005, snapback_rsi_min=48, snapback_vol_max=0.4, snapback_ema_days=15, snapback_dd_floor=-0.175, staged_dd_floor=-0.25, explosive_window_days=20, correction_reentry_mode='gated3', correction_window_days=10, correction_ema_days=10, correction_roc5_min=0.01, correction_roc10_min=-0.02, correction_rsi_min=42, correction_mfi_min=35, correction_atr_max=0.12, correction_vol_max=0.4, correction_dd_floor=-0.15, dd_rebound_mode='direct3', dd_rebound_window_days=30, dd_rebound_min=0.1, dd_rebound_roc5_min=0.02, dd_rebound_rsi_min=45, dd_rebound_mfi_min=35, dd_rebound_atr_max=0.18, dd_rebound_vol_max=1.2, dd_rebound_hold_days=8, dd_rebound_hold_roc5_min=-0.05, filter_count=5),
        Config(name='V9_FOCUS_207661', family='v9_frontier_cagr', sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.345, fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.3, crash_dd=-0.15, crash_vol_threshold=0.3, down_confirm=2, adaptive_down_mode='off', soft_down_confirm=3, hard_down_confirm=1, fast_mode='conditional', gate_scope='all_up', gate_logic='all', rsi_asset='qqq', rsi_min=45, mfi_asset='qqq', mfi_min=40, atr_asset='qqq', atr_max_pct=0.08, snapback_mode='off', snapback_days=20, snapback_roc10_min=0.01, snapback_rsi_min=45, snapback_vol_max=0.45, snapback_ema_days=20, snapback_dd_floor=-0.3, staged_dd_floor=-0.25, explosive_window_days=20, correction_reentry_mode='gated3', correction_window_days=10, correction_ema_days=15, correction_roc5_min=0.01, correction_roc10_min=-0.02, correction_rsi_min=45, correction_mfi_min=35, correction_atr_max=0.1, correction_vol_max=0.45, correction_dd_floor=-0.15, dd_rebound_mode='gated3', dd_rebound_window_days=30, dd_rebound_min=0.04, dd_rebound_roc5_min=0.04, dd_rebound_rsi_min=50, dd_rebound_mfi_min=35, dd_rebound_atr_max=0.3, dd_rebound_vol_max=1, dd_rebound_hold_days=8, dd_rebound_hold_roc5_min=-0.05, filter_count=5),
    ]

    # Deterministic V10 probes: DDR sweet-spot grid and score-hysteresis controls.
    # Nothing here references a calendar year/date; the same rules apply to every episode.
    for reb in [0.085, 0.090, 0.095, 0.100]:
        for rsi in [47, 48, 49, 50]:
            for r5 in [0.04, 0.06, 0.08]:
                seeds.append(Config(
                    name=f"V10_DDR_R{int(round(reb*1000)):03d}_RSI{rsi}_R5{int(round(r5*100)):02d}", family="v10_ddr_probe",
                    sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.34,
                    fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30,
                    crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2,
                    fast_mode="conditional", gate_scope="all_up", gate_logic="all",
                    rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08,
                    snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45,
                    snapback_vol_max=0.35, snapback_ema_days=20, snapback_dd_floor=-0.175,
                    correction_reentry_mode="qld", correction_window_days=10, correction_ema_days=10,
                    correction_roc5_min=0.01, correction_roc10_min=-0.02, correction_rsi_min=42,
                    correction_mfi_min=35, correction_atr_max=0.12, correction_vol_max=0.50, correction_dd_floor=-0.15,
                    dd_rebound_mode="direct3", dd_rebound_window_days=42, dd_rebound_min=reb,
                    dd_rebound_roc5_min=r5, dd_rebound_rsi_min=rsi, dd_rebound_mfi_min=40,
                    dd_rebound_atr_max=0.25, dd_rebound_vol_max=1.50, dd_rebound_hold_days=5,
                    dd_rebound_hold_roc5_min=-0.03, filter_count=6,
                ))
    for soft in [3, 4, 5, 6]:
        seeds.append(Config(
            name=f"V10_HYST_STABLE_S{soft}", family="v10_hyst_probe",
            sma_days=180, mom_days=120, mom_threshold=-0.01, vol_days=40, vol_threshold=0.3125,
            fast_ema_days=15, fast_mom_days=15, fast_threshold=0.005, fast_long_mom_floor=-0.30,
            crash_dd=-0.15, crash_vol_threshold=0.30, down_confirm=2,
            adaptive_down_mode="tqqq_score", soft_down_confirm=soft, hard_down_confirm=1,
            fast_mode="conditional", gate_scope="all_up", gate_logic="all",
            rsi_asset="qqq", rsi_min=45, mfi_asset="qqq", mfi_min=40, atr_asset="qqq", atr_max_pct=0.08,
            snapback_mode="gated3", snapback_days=10, snapback_roc10_min=0.01, snapback_rsi_min=45,
            snapback_vol_max=0.40, snapback_ema_days=20, snapback_dd_floor=-0.15, filter_count=5,
        ))


    # V11 control: exact clone of V8_FOCUS_211857. Keeping every field identical
    # makes the pre-DDR2 state machine an exact anchor rather than an approximation.
    d_core = asdict(next(x for x in seeds if x.name == "V8_FOCUS_211857"))
    d_core.update(name="V11_CORE_211857", family="v11_core_stable")
    seeds.append(Config(**d_core))

    # Deterministic trough-confirmed probes around the observed generic crash-recovery
    # sweet spot. No year or date is used in strategy logic.
    probe_specs = [
        ("D3_A3_R090", "direct3", 3, 0.090, 0.06, 0.00, 48, 40, 1),
        ("D3_A4_R090", "direct3", 4, 0.090, 0.06, 0.00, 48, 40, 1),
        ("D3_A5_R095", "direct3", 5, 0.095, 0.06, 0.00, 48, 40, 1),
        ("D3_A6_R095", "direct3", 6, 0.095, 0.08, 0.00, 48, 40, 1),
        ("D3_A4_R095_R10P2", "direct3", 4, 0.095, 0.06, 0.02, 48, 40, 1),
        ("G3_A4_R095", "gated3", 4, 0.095, 0.06, 0.00, 48, 40, 1),
        ("Q3_A3_R090_C1", "qld_confirm3", 3, 0.090, 0.06, 0.00, 47, 40, 1),
        ("Q3_A4_R095_C1", "qld_confirm3", 4, 0.095, 0.06, 0.00, 48, 40, 1),
        ("Q3_A4_R095_C2", "qld_confirm3", 4, 0.095, 0.06, 0.00, 48, 40, 2),
        ("Q3_A5_R095_C2", "qld_confirm3", 5, 0.095, 0.08, 0.00, 48, 45, 2),
    ]
    for tag, mode, age, reb, r5, r10, rsi, mfi, confirm in probe_specs:
        d = asdict(next(x for x in seeds if x.name == "V11_CORE_211857"))
        d.update(
            name=f"V11_TR2_{tag}", family="v11_trough_probe",
            trough_recovery_mode=mode, trough_recovery_window_days=42,
            trough_age_min=age, trough_min_crash_days=3,
            trough_rebound_min=reb, trough_roc5_min=r5, trough_roc10_min=r10,
            trough_rsi_min=rsi, trough_mfi_min=mfi, trough_atr_max=0.25,
            trough_vol_max=1.20, trough_confirm_days=confirm,
            trough_hold_days=5, trough_hold_roc5_min=-0.03, trough_hold_roc10_min=-0.05,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Exact V11 user-run frontiers from ROBUST_TIGHT_V11_DASHBOARD_20260911_133612.xlsx.
    # All pre-V12 fields are cloned from V8_FOCUS_211857 and only the observed V11
    # DDR2 fields below are changed. These remain immutable controls in V12.
    v11_frontiers = [
        ("V11_FOCUS_000365", "v11_frontier_cagr_crash", dict(
            trough_recovery_mode="qld_confirm3", trough_age_min=5, trough_rebound_min=0.0975,
            trough_roc5_min=0.04, trough_roc10_min=0.05, trough_rsi_min=46,
            trough_atr_max=0.30, trough_vol_max=1.20, trough_confirm_days=2,
            trough_hold_roc5_min=-0.05, trough_hold_roc10_min=0.00)),
        ("V11_FOCUS_003898", "v11_frontier_stable_mdd", dict(
            trough_recovery_mode="gated3", trough_min_crash_days=4, trough_rebound_min=0.11,
            trough_roc5_min=0.04, trough_roc10_min=0.02, trough_rsi_min=49,
            trough_atr_max=0.15, trough_confirm_days=2, trough_hold_days=3,
            trough_hold_roc5_min=-0.01)),
        ("V11_FOCUS_012077", "v11_frontier_stable_direct", dict(
            trough_recovery_mode="direct3", trough_age_min=5, trough_rebound_min=0.10,
            trough_roc5_min=0.04, trough_roc10_min=0.01, trough_rsi_min=50,
            trough_atr_max=0.22, trough_hold_days=3, trough_hold_roc5_min=-0.01)),
    ]
    for nm, fam, upd in v11_frontiers:
        d = asdict(next(x for x in seeds if x.name == "V8_FOCUS_211857"))
        d.update(name=nm, family=fam, **upd)
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Deterministic V12 split-bridge probes. QLD entry is intentionally close to
    # V11_FOCUS_000365; TQQQ promotion is a separate gate and never reuses entry criteria.
    split_specs = [
        ("V12_SPLIT_BASE", 5, 0.0975, 0.04, 0.05, 46, 2, 8, 0.11, 0.02, 0.08, 50, 45, 0.10, 0.85),
        ("V12_SPLIT_EARLY_Q", 4, 0.0950, 0.04, 0.04, 46, 3, 10, 0.11, 0.01, 0.08, 50, 45, 0.10, 0.85),
        ("V12_SPLIT_STRONG3", 5, 0.0975, 0.04, 0.05, 46, 4, 10, 0.12, 0.02, 0.10, 52, 45, 0.10, 0.80),
        ("V12_SPLIT_PATIENT", 5, 0.1000, 0.03, 0.04, 46, 5, 12, 0.12, 0.00, 0.08, 50, 40, 0.12, 0.90),
    ]
    for spec in split_specs:
        (nm, age, reb, r5, r10, rsi, wait, bridge_max, preb, pr5, pr10, prsi, pmfi, patr, pvol) = spec
        d = asdict(next(x for x in seeds if x.name == "V8_FOCUS_211857"))
        d.update(
            name=nm, family="v12_split_probe", trough_recovery_mode="qld_split3",
            trough_recovery_window_days=42, trough_age_min=age, trough_min_crash_days=3,
            trough_rebound_min=reb, trough_roc5_min=r5, trough_roc10_min=r10,
            trough_rsi_min=rsi, trough_mfi_min=40, trough_atr_max=0.30, trough_vol_max=1.20,
            trough_confirm_days=wait, trough_hold_days=5, trough_hold_roc5_min=-0.05, trough_hold_roc10_min=-0.05,
            trough_bridge_max_days=bridge_max, trough_promote_rebound_min=preb,
            trough_promote_roc5_min=pr5, trough_promote_roc10_min=pr10, trough_promote_rsi_min=prsi,
            trough_promote_mfi_min=pmfi, trough_promote_atr_max=patr, trough_promote_vol_max=pvol,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Deterministic V12 whipsaw-regime probes on the V11 000365 frontier.
    for mode, lb, cnt, cool, r10, rsi, vu in [
        ("score3_or_strong", 60, 2, 0, 0.02, 50, 0.95),
        ("score3_or_strong", 40, 2, 0, 0.03, 52, 0.90),
        ("cooldown_strong", 60, 2, 5, 0.02, 50, 0.95),
        ("score3_only", 60, 2, 0, 0.02, 50, 0.95),
    ]:
        d = asdict(next(x for x in seeds if x.name == "V11_FOCUS_000365"))
        tag = f"{mode}_L{lb}_N{cnt}_C{cool}_R{int(r10*100):02d}_RSI{rsi}_V{int(vu*100):02d}"
        d.update(name=f"V12_GUARD_{tag}", family="v12_guard_probe",
                 repromote_guard_mode=mode, repromote_lookback_days=lb,
                 repromote_trigger_count=cnt, repromote_cooldown_days=cool,
                 repromote_roc10_min=r10, repromote_rsi_min=rsi, repromote_vol_util_max=vu)
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Exact V12 user-run winner/frontier from ROBUST_TIGHT_V12_DASHBOARD_20260911_143815.xlsx.
    # This is the new V13 stable anchor: it Pareto-improves the prior stable frontiers while
    # preserving the 98% rolling-MDD win rate. No calendar/date rule is introduced.
    d = asdict(next(x for x in seeds if x.name == "V8_FOCUS_211857"))
    d.update(
        name="V12_FOCUS_011962", family="v12_frontier_stable_split",
        trough_recovery_mode="qld_split3", trough_recovery_window_days=42,
        trough_age_min=5, trough_min_crash_days=4, trough_rebound_min=0.095,
        trough_roc5_min=0.03, trough_roc10_min=0.05, trough_rsi_min=45,
        trough_mfi_min=45, trough_atr_max=0.30, trough_vol_max=1.20,
        trough_confirm_days=4, trough_hold_days=3,
        trough_hold_roc5_min=-0.05, trough_hold_roc10_min=-0.05,
        trough_bridge_max_days=15, trough_promote_rebound_min=0.10,
        trough_promote_roc5_min=0.03, trough_promote_roc10_min=0.06,
        trough_promote_rsi_min=50, trough_promote_mfi_min=45,
        trough_promote_atr_max=0.10, trough_promote_vol_max=1.00,
        repromote_guard_mode="off",
    )
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    # Deterministic V13 micro-grid. Every point keeps the successful V12 split bridge
    # fixed and changes only the core volatility boundary + snapback DD floor. This
    # guarantees exact-stage coverage of the most interpretable path from the 78.76%
    # stable frontier toward the historical 79.00% return frontier.
    v13_vol_grid = [0.3050,0.3075,0.3100,0.3125,0.3150,0.3175,0.3200,0.3225,
                    0.3250,0.3275,0.3300,0.3325,0.3350,0.3375,0.3400]
    v13_snap_grid = [-0.1400,-0.1500,-0.1625,-0.1750,-0.2000]
    for vol in v13_vol_grid:
        for sdd in v13_snap_grid:
            g = asdict(next(x for x in seeds if x.name == "V12_FOCUS_011962"))
            g.update(
                name=f"V13_GRID_V{int(round(vol*10000)):04d}_D{int(round(abs(sdd)*10000)):04d}",
                family="v13_core_grid", vol_threshold=vol, snapback_dd_floor=sdd,
                repromote_guard_mode="off",
            )
            g["filter_count"] = _filter_count(g)
            seeds.append(Config(**g))

    # V14 deterministic selective soft-volatility probes. All start from the exact
    # V12/V13 stable split anchor and keep the global 31.25% volatility threshold.
    # The new module can only promote QLD -> TQQQ; it never delays an exit.
    v14_soft_specs = [
        ("RECENT_V325_R10_2_R20_4_RSI52", "recent3", 0.3250, 20, 1, 0.02, 0.04, 0.00, 52, 45, -0.15),
        ("RECENT_V330_R10_3_R20_5_RSI54", "recent3", 0.3300, 30, 1, 0.03, 0.05, 0.00, 54, 45, -0.15),
        ("RECENT_V335_R10_4_R20_6_RSI55", "recent3", 0.3350, 30, 2, 0.04, 0.06, 0.00, 55, 50, -0.12),
        ("STRONG_V325_R10_2_R20_4_RSI52", "score2_strong3", 0.3250, 20, 1, 0.02, 0.04, 0.00, 52, 45, -0.15),
        ("STRONG_V330_R10_3_R20_5_RSI54", "score2_strong3", 0.3300, 20, 1, 0.03, 0.05, 0.02, 54, 45, -0.15),
        ("STRONG_V335_R10_4_R20_6_RSI56", "score2_strong3", 0.3350, 20, 2, 0.04, 0.06, 0.02, 56, 50, -0.12),
    ]
    for tag, mode, vmax, recent, confirm, r10, r20, r120, rsi, mfi, ddf in v14_soft_specs:
        d = asdict(next(x for x in seeds if x.name == "V12_FOCUS_011962"))
        d.update(
            name=f"V14_SV_{tag}", family="v14_softvol_probe",
            vol_threshold=0.3125, snapback_dd_floor=-0.15, repromote_guard_mode="off",
            soft_vol_mode=mode, soft_vol_max=vmax,
            soft_vol_recent_delever_days=recent, soft_vol_confirm_days=confirm,
            soft_vol_roc10_min=r10, soft_vol_roc20_min=r20, soft_vol_roc120_min=r120,
            soft_vol_rsi_min=rsi, soft_vol_mfi_min=mfi, soft_vol_dd_floor=ddf,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # V15 deterministic impulse-veto probes. These are based on the exact V12 stable
    # anchor. Soft-vol is disabled because the V14 real-data run showed that branch
    # was structurally dead. The first probe is intentionally permissive and serves
    # as an activation canary; the others are centered on the observed 1-day churn
    # pattern where short ROC/RSI/MFI are strong but the slow core votes still lag.
    v15_impulse_specs = [
        # Activation canary: deliberately permissive, but still requires the ordinary entry gate
        # and never applies in crash/post-crash context. It exists to prove the new branch is alive.
        ("CANARY", "score12_strong", 1, 5, -0.01, -0.03, -0.05, 40, 30, 0.10, 0.50, -0.30),
        # Strong short impulse while the slow trend/momentum votes lag.
        ("S1_STRONG", "score1_strong", 1, 5, 0.04, 0.04, -0.01, 52, 40, 0.04, 0.35, -0.15),
        # Score-2 cases are safer by construction because only one slow core vote is missing.
        ("S2_BALANCED", "score2_strong", 1, 5, 0.00, 0.00, -0.02, 50, 40, 0.04, 0.35, -0.15),
        ("S2_RECOVERY", "score2_strong", 1, 5, -0.015, -0.005, 0.00, 49, 45, 0.03, 0.30, -0.12),
        ("S2_MOMENTUM", "score2_strong", 1, 5, 0.00, 0.03, 0.02, 49, 50, 0.03, 0.30, -0.12),
        # Broad score-1/2 version for interaction testing.
        ("S12_BALANCED", "score12_strong", 1, 5, 0.02, 0.02, -0.02, 50, 40, 0.04, 0.35, -0.18),
        ("S12_STRONG_H2", "score12_strong", 2, 5, 0.04, 0.04, -0.01, 52, 40, 0.04, 0.35, -0.15),
    ]
    for tag, mode, maxd, ema, r5, r10, r20, rsi, mfi, atr, vmax, ddf in v15_impulse_specs:
        d = asdict(next(x for x in seeds if x.name == "V12_FOCUS_011962"))
        d.update(
            name=f"V15_IV_{tag}", family="v15_impulse_probe",
            vol_threshold=0.3125, snapback_dd_floor=-0.15, repromote_guard_mode="off",
            soft_vol_mode="off",
            impulse_veto_mode=mode, impulse_veto_max_days=maxd, impulse_veto_ema_days=ema,
            impulse_veto_roc5_min=r5, impulse_veto_roc10_min=r10, impulse_veto_roc20_min=r20,
            impulse_veto_rsi_min=rsi, impulse_veto_mfi_min=mfi, impulse_veto_atr_max=atr,
            impulse_veto_vol_max=vmax, impulse_veto_dd_floor=ddf,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Exact V15 real-data winner from ROBUST_TIGHT_V15_DASHBOARD_20260911_223549.xlsx.
    # It is the first model to clear BOTH>=79% and MDD-Win>=97% simultaneously.
    d = asdict(next(x for x in seeds if x.name == "V12_FOCUS_011962"))
    d.update(
        name="V15_FOCUS_000979", family="v15_frontier_joint",
        vol_threshold=0.3125, snapback_dd_floor=-0.15, repromote_guard_mode="off", soft_vol_mode="off",
        impulse_veto_mode="score2_strong", impulse_veto_max_days=1, impulse_veto_ema_days=10,
        impulse_veto_roc5_min=0.00, impulse_veto_roc10_min=0.03, impulse_veto_roc20_min=-0.01,
        impulse_veto_rsi_min=48, impulse_veto_mfi_min=45, impulse_veto_atr_max=0.04,
        impulse_veto_vol_max=0.40, impulse_veto_dd_floor=-0.18, precision_veto_mode="off",
    )
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    # Conservative V15 plateau anchor: same realized V15 path, tighter unused margins.
    d = asdict(next(x for x in seeds if x.name == "V15_FOCUS_000979"))
    d.update(
        name="V15_FOCUS_006433", family="v15_frontier_conservative",
        impulse_veto_roc20_min=0.02, impulse_veto_mfi_min=50, impulse_veto_atr_max=0.03,
        impulse_veto_vol_max=0.30, impulse_veto_dd_floor=-0.15,
    )
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    # Deterministic V16 probes provide interpretable anchors in real-data runs.
    # Branch coverage is checked separately with permissive synthetic validation configs.
    v16_specs = [
        ("TREND_NEAR", "trend_near", 1, 10, -0.01, 0.00, -0.02, 48, 40, 0.05, 0.40, -0.20, -0.05, -0.05, 3, 1),
        ("MOM_NEAR", "mom_near", 1, 10, -0.01, 0.00, -0.02, 48, 40, 0.05, 0.40, -0.20, -0.10, -0.03, 3, 1),
        ("SCORE1_NEAR", "score1_near", 1, 10, 0.02, 0.03, -0.02, 50, 40, 0.04, 0.35, -0.18, -0.10, -0.03, 3, 1),
        ("FRESH_REENTRY", "fresh_reentry", 1, 10, 0.02, 0.03, -0.02, 50, 40, 0.04, 0.35, -0.18, -0.20, -0.08, 3, 1),
        ("CANARY", "fresh_reentry", 1, 20, -0.05, -0.05, -0.10, 40, 30, 0.10, 0.50, -0.30, -0.30, -0.20, 8, 1),
    ]
    for tag, mode, maxd, ema, r5, r10, r20, rsi, mfi, atr, vmax, ddf, tgap, mgap, recent, smin in v16_specs:
        d = asdict(next(x for x in seeds if x.name == "V15_FOCUS_000979"))
        d.update(
            name=f"V16_PV_{tag}", family="v16_precision_probe", precision_veto_mode=mode,
            precision_veto_max_days=maxd, precision_veto_ema_days=ema,
            precision_veto_roc5_min=r5, precision_veto_roc10_min=r10, precision_veto_roc20_min=r20,
            precision_veto_rsi_min=rsi, precision_veto_mfi_min=mfi, precision_veto_atr_max=atr,
            precision_veto_vol_max=vmax, precision_veto_dd_floor=ddf,
            precision_veto_trend_gap_floor=tgap, precision_veto_mom_gap_floor=mgap,
            precision_veto_recent_reentry_days=recent, precision_veto_score_min=smin,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Deterministic V17 probes. The primary probes are based on the generic pattern
    # observed in V16 diagnostics: score=2, Trend PASS, Momentum FAIL, Entry Gate PASS,
    # and a recent TQQQ re-entry. No event date is encoded; these are state predicates.
    v17_specs = [
        # tag, mode, min_age, max_age, post_crash_min, entry_gate, ema, tgap, mgap, r5, r10, r20, rsi, mfi, atr, vmax, ddf
        ("MOM_A17_CORE", "mom_reentry", 2, 17, 30, "pass", 0, -0.05, -0.012, -0.03, -0.06, -0.08, 40, 40, 0.03, 0.30, -0.22),
        ("MOM_A20_LOOSE", "mom_reentry", 1, 20, 20, "pass", 0, -0.08, -0.020, -0.05, -0.07, -0.10, 35, 35, 0.04, 0.35, -0.25),
        ("MOM_A13_TIGHT", "mom_reentry", 2, 13, 30, "pass", 0, -0.03, -0.010, -0.03, -0.06, -0.06, 45, 40, 0.03, 0.30, -0.22),
        ("S2_A20_PASS", "score2_reentry", 2, 20, 30, "pass", 0, -0.03, -0.015, -0.04, -0.06, -0.08, 40, 35, 0.04, 0.35, -0.25),
        ("TREND_A5_ANY", "trend_reentry", 1, 5, 30, "any", 0, -0.030, -0.10, -0.04, -0.06, -0.08, 35, 30, 0.04, 0.35, -0.20),
        # Activation canary: intentionally broad but still score-2 only and one session.
        ("CANARY", "score2_reentry", 1, 30, 10, "any", 0, -0.10, -0.10, -0.10, -0.12, -0.15, 25, 20, 0.10, 0.50, -0.40),
    ]
    for tag, mode, amin, amax, post, eg, ema, tgap, mgap, r5, r10, r20, rsi, mfi, atr, vmax, ddf in v17_specs:
        d = asdict(next(x for x in seeds if x.name == "V15_FOCUS_000979"))
        d.update(
            name=f"V17_RG_{tag}", family="v17_reentry_probe", precision_veto_mode="off",
            reentry_grace_mode=mode, reentry_grace_max_days=1,
            reentry_grace_min_age=amin, reentry_grace_max_age=amax,
            reentry_grace_min_post_crash_days=post, reentry_grace_entry_gate=eg,
            reentry_grace_ema_days=ema, reentry_grace_trend_gap_floor=tgap,
            reentry_grace_mom_gap_floor=mgap, reentry_grace_roc5_min=r5,
            reentry_grace_roc10_min=r10, reentry_grace_roc20_min=r20,
            reentry_grace_rsi_min=rsi, reentry_grace_mfi_min=mfi,
            reentry_grace_atr_max=atr, reentry_grace_vol_max=vmax, reentry_grace_dd_floor=ddf,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Exact V17 winner from ROBUST_TIGHT_V17_DASHBOARD_20260912_145240.xlsx.
    # It improved V15 from 2,892 to 2,904 1Y BOTH wins while keeping full MDD unchanged.
    d = asdict(next(x for x in seeds if x.name == "V15_FOCUS_000979"))
    d.update(
        name="V17_FOCUS_471542", family="v17_frontier_reentry", precision_veto_mode="off",
        reentry_grace_mode="score2_reentry", reentry_grace_max_days=1,
        reentry_grace_min_age=2, reentry_grace_max_age=20, reentry_grace_min_post_crash_days=30,
        reentry_grace_entry_gate="any", reentry_grace_ema_days=0,
        reentry_grace_trend_gap_floor=-0.08, reentry_grace_mom_gap_floor=-0.012,
        reentry_grace_roc5_min=-0.02, reentry_grace_roc10_min=-0.07, reentry_grace_roc20_min=-0.06,
        reentry_grace_rsi_min=40, reentry_grace_mfi_min=35, reentry_grace_atr_max=0.05,
        reentry_grace_vol_max=0.325, reentry_grace_dd_floor=-0.22,
        reentry_grace_mom_entry_gate="pass", reentry_grace_trend_entry_gate="any",
        reentry_grace_mom_roc5_min=-0.03, reentry_grace_trend_roc5_min=0.00,
    )
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    # V18 deterministic split probes. DAY1 tests the four-event union logic; TREND2
    # allows at most one adjacent second day and only when the second day is Trend-fail.
    for tag, max_days in [("SPLIT_DAY1", 1), ("SPLIT_TREND2", 2)]:
        d = asdict(next(x for x in seeds if x.name == "V17_FOCUS_471542"))
        d.update(
            name=f"V18_{tag}", family="v18_split_probe", reentry_grace_mode="split_reentry",
            reentry_grace_max_days=max_days, reentry_grace_min_age=2, reentry_grace_max_age=20,
            reentry_grace_min_post_crash_days=30, reentry_grace_entry_gate="any", reentry_grace_ema_days=0,
            reentry_grace_trend_gap_floor=-0.03, reentry_grace_mom_gap_floor=-0.012,
            reentry_grace_roc5_min=-0.03, reentry_grace_roc10_min=-0.06, reentry_grace_roc20_min=-0.06,
            reentry_grace_rsi_min=40, reentry_grace_mfi_min=35, reentry_grace_atr_max=0.03,
            reentry_grace_vol_max=0.325, reentry_grace_dd_floor=-0.22,
            reentry_grace_mom_entry_gate="pass", reentry_grace_trend_entry_gate="any",
            reentry_grace_mom_roc5_min=-0.03, reentry_grace_trend_roc5_min=0.00,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Exact V18 79.5%+97% anchor from ROBUST_TIGHT_V18_DASHBOARD_20260912_154739.xlsx.
    # Same path as deterministic V18_SPLIT_TREND2: 2,914 / 3,659 1Y wins,
    # Worst BOTH 79.6392%, MDD Win 98.1054%, CAGR ~51.64%, Full MDD unchanged.
    d = asdict(next(x for x in seeds if x.name == "V18_SPLIT_TREND2"))
    d.update(
        name="V18_FOCUS_011262", family="v18_frontier_795", micro_rescue_mode="off",
        reentry_grace_trend_gap_floor=-0.05, reentry_grace_roc5_min=-0.02,
        reentry_grace_roc10_min=-0.07, reentry_grace_roc20_min=-0.10,
        reentry_grace_mfi_min=30, reentry_grace_atr_max=0.05,
        reentry_grace_mom_roc5_min=-0.025, reentry_grace_trend_roc5_min=0.01,
    )
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    # V19 deterministic state probes. These are state predicates only; the dates that
    # motivated them are NOT encoded. Each layer is additive on top of the V18 anchor.
    v19_specs = [
        # tag, mode
        ("S1_PULLBACK", "score1_pullback"),
        ("S1_NEARTREND", "score1_neartrend"),
        ("MATURE_MOM", "mature_mom"),
        ("COMBO", "combo"),
    ]
    for tag, mode in v19_specs:
        d = asdict(next(x for x in seeds if x.name == "V18_FOCUS_011262"))
        d.update(
            name=f"V19_MICRO_{tag}", family="v19_micro_probe", micro_rescue_mode=mode,
            micro_rescue_max_days=1, micro_rescue_min_post_crash_days=30,
            micro_rescue_recent_max_age=10, micro_rescue_neartrend_max_age=60,
            micro_rescue_mature_min_age=60, micro_rescue_mature_max_age=250,
            micro_rescue_score1_trend_gap_floor=-0.05, micro_rescue_score1_mom_gap_floor=-0.012,
            micro_rescue_score1_roc5_min=-0.03, micro_rescue_score1_roc5_max=0.00,
            micro_rescue_score1_roc10_min=0.00, micro_rescue_score1_roc20_min=-0.05,
            micro_rescue_score1_rsi_min=40, micro_rescue_score1_mfi_min=35,
            micro_rescue_neartrend_gap_floor=-0.01, micro_rescue_neartrend_roc5_min=-0.03,
            micro_rescue_neartrend_roc5_max=0.03, micro_rescue_neartrend_roc20_min=0.00,
            micro_rescue_neartrend_rsi_min=48, micro_rescue_neartrend_mfi_min=35,
            micro_rescue_mature_mom_gap_floor=-0.0025, micro_rescue_mature_trend_gap_floor=0.0,
            micro_rescue_mature_roc5_min=-0.03, micro_rescue_mature_rsi_min=25, micro_rescue_mature_mfi_min=20,
            micro_rescue_atr_max=0.04, micro_rescue_vol_max=0.35, micro_rescue_dd_floor=-0.25,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Exact V19 80%+97% anchor. This is the two-block near-trend plateau path:
    # 1Y 2,957/3,659, Worst BOTH 80.7749% (2Y bottleneck), MDD Win 98.1054%.
    d = asdict(next(x for x in seeds if x.name == "V18_FOCUS_011262"))
    d.update(
        name="V19_FOCUS_001037", family="v19_frontier_80",
        micro_rescue_mode="combo", micro_rescue_max_days=1, micro_rescue_min_post_crash_days=30,
        micro_rescue_recent_max_age=10, micro_rescue_neartrend_max_age=90,
        micro_rescue_mature_min_age=60, micro_rescue_mature_max_age=250,
        micro_rescue_score1_trend_gap_floor=-0.03, micro_rescue_score1_mom_gap_floor=-0.01,
        micro_rescue_score1_roc5_min=-0.03, micro_rescue_score1_roc5_max=0.00,
        micro_rescue_score1_roc10_min=0.01, micro_rescue_score1_roc20_min=0.00,
        micro_rescue_score1_rsi_min=45, micro_rescue_score1_mfi_min=35,
        micro_rescue_neartrend_gap_floor=-0.0075, micro_rescue_neartrend_roc5_min=-0.02,
        micro_rescue_neartrend_roc5_max=0.03, micro_rescue_neartrend_roc20_min=0.00,
        micro_rescue_neartrend_rsi_min=48, micro_rescue_neartrend_mfi_min=35,
        micro_rescue_mature_mom_gap_floor=-0.001, micro_rescue_mature_trend_gap_floor=0.003,
        micro_rescue_mature_roc5_min=0.00, micro_rescue_mature_rsi_min=30, micro_rescue_mature_mfi_min=35,
        micro_rescue_atr_max=0.05, micro_rescue_vol_max=0.40, micro_rescue_dd_floor=-0.18,
        shadow_rescue_mode="off",
    )
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    # V20 deterministic shadow probes. Underlying control state still steps to QLD;
    # only next-session execution is temporarily held at TQQQ.
    for tag, mode in [("S1_RECENT", "score1_recent"), ("MATURE_MOM", "mature_mom"), ("COMBO", "combo")]:
        d = asdict(next(x for x in seeds if x.name == "V19_FOCUS_001037"))
        d.update(
            name=f"V20_SHADOW_{tag}", family="v20_shadow_probe", shadow_rescue_mode=mode,
            shadow_min_post_crash_days=30, shadow_recent_max_age=10,
            shadow_mature_min_age=60, shadow_mature_max_age=250,
            shadow_score1_trend_gap_floor=-0.05, shadow_score1_mom_gap_floor=-0.012,
            shadow_score1_roc5_min=-0.03, shadow_score1_roc5_max=0.00,
            shadow_score1_roc10_min=0.00, shadow_score1_roc20_min=-0.05,
            shadow_score1_rsi_min=40, shadow_score1_mfi_min=35,
            shadow_mature_mom_gap_floor=-0.0025, shadow_mature_trend_gap_floor=0.0,
            shadow_mature_roc5_min=-0.03, shadow_mature_roc5_max=0.00,
            shadow_mature_rsi_min=25, shadow_mature_mfi_min=20,
            shadow_atr_max=0.04, shadow_vol_max=0.35, shadow_dd_floor=-0.25,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Exact V20 frontier anchor observed in DEEP: V20_SHADOW_COMBO.
    # Worst BOTH 80.8629% (2Y bottleneck), MDD Win 98.1054%, CAGR ~51.98%.
    d = asdict(next(x for x in seeds if x.name == "V20_SHADOW_COMBO"))
    d.update(name="V20_FOCUS_SHADOW_COMBO", family="v20_frontier_8086", preview_overlay_mode="off")
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    # V21 deterministic preview probes. These never alter the control-state path.
    for tag, mode in [("QLD3", "qld_score3"), ("QLD2", "qld_score2"), ("STEPUP", "stepup"), ("COMBO", "combo")]:
        d = asdict(next(x for x in seeds if x.name == "V20_FOCUS_SHADOW_COMBO"))
        d.update(
            name=f"V21_PREVIEW_{tag}", family="v21_preview_probe", preview_overlay_mode=mode,
            preview_min_post_crash_days=30, preview_qld_max_delever_age=40,
            preview_qld3_roc5_min=-0.01, preview_qld3_roc10_min=0.00, preview_qld3_roc20_min=0.00,
            preview_qld3_rsi_min=45, preview_qld3_mfi_min=35,
            preview_qld2_trend_gap_floor=-0.01, preview_qld2_mom_gap_floor=-0.005,
            preview_qld2_roc5_min=-0.01, preview_qld2_roc10_min=0.00, preview_qld2_roc20_min=0.00,
            preview_qld2_rsi_min=48, preview_qld2_mfi_min=40,
            preview_stepup_score_min=2, preview_stepup_roc5_min=0.02, preview_stepup_roc10_min=0.02,
            preview_stepup_roc20_min=0.00, preview_stepup_rsi_min=50, preview_stepup_mfi_min=40,
            preview_atr_max=0.035, preview_vol_max=0.325, preview_dd_floor=-0.25,
        )
        d["filter_count"] = _filter_count(d)
        seeds.append(Config(**d))

    # Synthetic canary: intentionally loose, only to prove V21 branch coverage.
    d = asdict(next(x for x in seeds if x.name == "V20_FOCUS_SHADOW_COMBO"))
    d.update(
        name="V21_PREVIEW_CANARY", family="v21_preview_probe", preview_overlay_mode="combo",
        preview_min_post_crash_days=0, preview_qld_max_delever_age=10000,
        preview_qld3_roc5_min=-1.0, preview_qld3_roc10_min=-1.0, preview_qld3_roc20_min=-1.0,
        preview_qld3_rsi_min=0, preview_qld3_mfi_min=0,
        preview_qld2_trend_gap_floor=-1.0, preview_qld2_mom_gap_floor=-1.0,
        preview_qld2_roc5_min=-1.0, preview_qld2_roc10_min=-1.0, preview_qld2_roc20_min=-1.0,
        preview_qld2_rsi_min=0, preview_qld2_mfi_min=0,
        preview_stepup_score_min=0, preview_stepup_roc5_min=-1.0, preview_stepup_roc10_min=-1.0,
        preview_stepup_roc20_min=-1.0, preview_stepup_rsi_min=0, preview_stepup_mfi_min=0,
        preview_atr_max=1.0, preview_vol_max=10.0, preview_dd_floor=-1.0,
    )
    d["filter_count"] = _filter_count(d)
    seeds.append(Config(**d))

    return seeds


def generate_random_configs(trials: int, seed: int) -> List[Config]:
    rng = random.Random(seed)
    seen = set()
    out: List[Config] = []

    # Parameter neighborhoods intentionally include the original values.
    sma_choices = [120, 140, 150, 165, 180, 200, 220]
    mom_days_choices = [60, 90, 120, 150]
    mom_thresh_choices = [-0.03, -0.01, 0.00, 0.02, 0.04]
    vol_days_choices = [20, 40, 60]
    vol_thresh_choices = [0.20, 0.225, 0.25, 0.275, 0.30, 0.325]
    fast_ema_choices = [5, 10, 15, 20]
    fast_mom_days_choices = [5, 10, 15, 20]
    fast_thresh_choices = [0.02, 0.03, 0.04, 0.05, 0.06, 0.08]
    long_floor_choices = [-0.12, -0.08, -0.05, 0.00]
    crash_dd_choices = [-0.15, -0.18, -0.20, -0.23, -0.25, -0.28, -0.30]
    crash_vol_choices = [0.225, 0.25, 0.275, 0.30, 0.35]

    for i in range(trials * 3):
        if len(out) >= trials:
            break

        # Favor moderate complexity: most candidates use 1-5 extra gates, not 12 at once.
        intensity = rng.choices([0, 1, 2, 3, 4, 5], weights=[5, 12, 20, 22, 16, 8], k=1)[0]
        d: Dict[str, Any] = dict(
            name=f"RND_{len(out)+1:06d}",
            family="random",
            sma_days=rng.choice(sma_choices),
            mom_days=rng.choice(mom_days_choices),
            mom_threshold=rng.choice(mom_thresh_choices),
            vol_days=rng.choice(vol_days_choices),
            vol_threshold=rng.choice(vol_thresh_choices),
            fast_ema_days=rng.choice(fast_ema_choices),
            fast_mom_days=rng.choice(fast_mom_days_choices),
            fast_threshold=rng.choice(fast_thresh_choices),
            fast_long_mom_floor=rng.choice(long_floor_choices),
            crash_dd=rng.choice(crash_dd_choices),
            crash_vol_threshold=rng.choice(crash_vol_choices),
            down_confirm=rng.choice([2, 2, 2, 3, 3, 4]),
            fast_mode=rng.choices(["direct", "qld", "conditional", "normal"], weights=[28, 25, 32, 15], k=1)[0],
            gate_scope=rng.choices(["all_up", "tqqq_only", "fast_only"], weights=[25, 55, 20], k=1)[0],
            require_qqq_ma50=False,
            require_qqq_ma200=False,
            require_tqqq_ma20=False,
            require_tqqq_ma50=False,
            macd_gate="none",
            rsi_asset="none",
            rsi_min=rng.choice([45, 50, 55, 60]),
            stoch_asset="none",
            cci_asset="none",
            boll_asset="none",
            boll_rule=rng.choice(["mid", "upper_half", "not_overbought"]),
            adx_asset="none",
            adx_min=rng.choice([15, 18, 20, 22, 25]),
            mfi_asset="none",
            mfi_min=rng.choice([45, 50, 55, 60]),
            willr_asset="none",
            willr_min=rng.choice([-70, -60, -50, -40, -30]),
            obv_asset="none",
            donchian_asset="none",
            donchian_min_pos=rng.choice([0.35, 0.50, 0.65, 0.80]),
            atr_asset="none",
            atr_max_pct=rng.choice([0.04, 0.05, 0.06, 0.08, 0.10]),
            roc_asset="none",
            roc_days=rng.choice([10, 20, 40, 60]),
            roc_min=rng.choice([-0.03, 0.00, 0.02, 0.04]),
            defense_mode="off",
            defense_score_threshold=rng.choice([6, 7, 8, 9, 10]),
            tqqq_exit_mode="off",
            tqqq_exit_score=rng.choice([5, 6, 7, 8, 9]),
            recovery_mode="off", recovery_days=rng.choice([42, 63, 84, 126]), recovery_mom_floor=rng.choice([-0.35, -0.30, -0.25, -0.20]), recovery_roc20_min=rng.choice([0.03, 0.05, 0.07, 0.10]), recovery_vol_max=rng.choice([0.45, 0.55, 0.65, 0.80]),
            bull_override_mode="off", bull_roc20_min=rng.choice([0.03, 0.05, 0.07, 0.10]), bull_rsi_min=rng.choice([50, 55, 60]),
            dd_guard_mode="off", dd_guard_days=rng.choice([20, 40, 60, 90]), dd_guard_threshold=rng.choice([-0.08, -0.10, -0.12, -0.15]),
            rebound_bypass_mode="off", rebound_window_days=rng.choice([20, 30, 42, 63, 84]), rebound_roc10_min=rng.choice([0.06, 0.08, 0.10, 0.12, 0.15]), rebound_roc20_min=rng.choice([0.10, 0.12, 0.15, 0.18, 0.22]), rebound_rsi_min=rng.choice([45, 50, 55, 60]), rebound_vol_max=rng.choice([0.60, 0.80, 1.00, 1.20]),
            bridge_mode="off", bridge_roc60_min=rng.choice([-0.03, -0.01, 0.00, 0.02, 0.04]), bridge_rsi_min=rng.choice([45, 50, 55]), bridge_dd_floor=rng.choice([-0.10, -0.15, -0.18, -0.20, -0.23]), bridge_vol_max=rng.choice([0.30, 0.35, 0.40, 0.45, 0.50]),
            filter_count=0,
        )

        gate_pool = [
            "qqq_ma50", "qqq_ma200", "tqqq_ma20", "tqqq_ma50",
            "macd", "rsi", "stoch", "cci", "boll", "adx", "mfi", "willr",
            "obv", "donchian", "atr", "roc", "recovery", "bull", "dd_guard", "rebound", "bridge", "defense", "tqqq_exit",
        ]
        chosen = rng.sample(gate_pool, k=min(intensity, len(gate_pool)))
        for g in chosen:
            if g == "qqq_ma50": d["require_qqq_ma50"] = True
            elif g == "qqq_ma200": d["require_qqq_ma200"] = True
            elif g == "tqqq_ma20": d["require_tqqq_ma20"] = True
            elif g == "tqqq_ma50": d["require_tqqq_ma50"] = True
            elif g == "macd": d["macd_gate"] = rng.choice(["qqq", "tqqq", "both"])
            elif g == "rsi": d["rsi_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "stoch": d["stoch_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "cci": d["cci_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "boll": d["boll_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "adx": d["adx_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "mfi": d["mfi_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "willr": d["willr_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "obv": d["obv_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "donchian": d["donchian_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "atr": d["atr_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "roc": d["roc_asset"] = rng.choice(["qqq", "tqqq"])
            elif g == "recovery": d["recovery_mode"] = rng.choice(["relaxed_fast", "strong_roc", "hybrid"])
            elif g == "bull": d["bull_override_mode"] = rng.choice(["trend_roc", "ema_roc"])
            elif g == "dd_guard": d["dd_guard_mode"] = rng.choice(["dd_macd", "dd_ema"])
            elif g == "rebound": d["rebound_bypass_mode"] = rng.choice(["roc10", "roc20", "hybrid"])
            elif g == "bridge": d["bridge_mode"] = rng.choice(["hold", "qld", "gated3", "direct3", "hold_gated3", "hold_direct3"])
            elif g == "defense": d["defense_mode"] = rng.choice(["cap2", "cap1", "cash"])
            elif g == "tqqq_exit": d["tqqq_exit_mode"] = rng.choice(["ma20_macd", "qqq50_macd", "risk"])

        d["filter_count"] = _filter_count(d)
        # Mild anti-overfit guard: cap active filters at 6.
        if d["filter_count"] > 6:
            continue

        key_items = tuple((k, d[k]) for k in sorted(d) if k not in {"name", "family", "filter_count"})
        if key_items in seen:
            continue
        seen.add(key_items)
        out.append(Config(**d))

    return out


def generate_focus_configs(trials: int, seed: int) -> List[Config]:
    """V21 focused search around the proven V20 80.86%+98.1% frontier.

    V20 is frozen. V21 adds a second stateless execution-only direction: a one-session
    TQQQ preview while the underlying control state remains/steps only to QLD. This
    attacks residual 2Y near misses without changing the multi-day state path.
    """
    if trials <= 0:
        return []
    rng = random.Random(seed)
    seeds = seeded_configs()
    anchor = next(x for x in seeds if x.name == "V20_FOCUS_SHADOW_COMBO")
    out: List[Config] = []
    seen = set()

    mode_choices = ["combo", "qld_score2", "qld_score3", "stepup", "off"]
    post_choices = [20, 30, 40, 50, 63, 90]
    qld_age_choices = [5, 10, 20, 30, 40, 60, 90, 120]
    q3_r5_choices = [-0.03, -0.02, -0.01, 0.0, 0.01, 0.02]
    q3_r10_choices = [-0.02, -0.01, 0.0, 0.01, 0.02, 0.03]
    q3_r20_choices = [-0.03, -0.01, 0.0, 0.02, 0.04, 0.06]
    q3_rsi_choices = [38, 42, 45, 48, 50, 52, 55]
    q3_mfi_choices = [30, 35, 40, 45, 50, 55]
    q2_tgap_choices = [-0.005, -0.0075, -0.01, -0.015, -0.02, -0.03]
    q2_mgap_choices = [-0.001, -0.0025, -0.005, -0.0075, -0.01, -0.015, -0.02]
    q2_r5_choices = [-0.03, -0.02, -0.01, 0.0, 0.01, 0.02]
    q2_r10_choices = [-0.03, -0.02, -0.01, 0.0, 0.01, 0.02, 0.03]
    q2_r20_choices = [-0.04, -0.02, 0.0, 0.02, 0.04, 0.06]
    q2_rsi_choices = [40, 45, 48, 50, 52, 55]
    q2_mfi_choices = [30, 35, 40, 45, 50, 55]
    step_score_choices = [2, 3]
    step_r5_choices = [0.0, 0.01, 0.02, 0.03, 0.04, 0.05]
    step_r10_choices = [0.0, 0.01, 0.02, 0.04, 0.06, 0.08]
    step_r20_choices = [-0.02, 0.0, 0.02, 0.04, 0.06, 0.10]
    step_rsi_choices = [45, 48, 50, 52, 55, 58]
    step_mfi_choices = [35, 40, 45, 50, 55, 60]
    atr_choices = [0.02, 0.025, 0.03, 0.035, 0.04, 0.05]
    vmax_choices = [0.25, 0.28, 0.30, 0.3125, 0.325, 0.35, 0.40]
    dd_choices = [-0.12, -0.15, -0.18, -0.22, -0.25, -0.30, -0.35]

    for _ in range(max(trials * 32, 1)):
        if len(out) >= trials:
            break
        d = asdict(anchor)
        d.update(name=f"FOCUS_{len(out)+1:06d}", family="v21_preview_overlay")
        d["preview_overlay_mode"] = rng.choices(mode_choices, weights=[56, 22, 10, 10, 2], k=1)[0]
        d["preview_min_post_crash_days"] = rng.choices(post_choices, weights=[8, 32, 24, 16, 14, 6], k=1)[0]
        d["preview_qld_max_delever_age"] = rng.choices(qld_age_choices, weights=[4, 8, 16, 20, 20, 14, 10, 8], k=1)[0]
        d["preview_qld3_roc5_min"] = rng.choice(q3_r5_choices)
        d["preview_qld3_roc10_min"] = rng.choice(q3_r10_choices)
        d["preview_qld3_roc20_min"] = rng.choice(q3_r20_choices)
        d["preview_qld3_rsi_min"] = rng.choice(q3_rsi_choices)
        d["preview_qld3_mfi_min"] = rng.choice(q3_mfi_choices)
        d["preview_qld2_trend_gap_floor"] = rng.choice(q2_tgap_choices)
        d["preview_qld2_mom_gap_floor"] = rng.choice(q2_mgap_choices)
        d["preview_qld2_roc5_min"] = rng.choice(q2_r5_choices)
        d["preview_qld2_roc10_min"] = rng.choice(q2_r10_choices)
        d["preview_qld2_roc20_min"] = rng.choice(q2_r20_choices)
        d["preview_qld2_rsi_min"] = rng.choice(q2_rsi_choices)
        d["preview_qld2_mfi_min"] = rng.choice(q2_mfi_choices)
        d["preview_stepup_score_min"] = rng.choice(step_score_choices)
        d["preview_stepup_roc5_min"] = rng.choice(step_r5_choices)
        d["preview_stepup_roc10_min"] = rng.choice(step_r10_choices)
        d["preview_stepup_roc20_min"] = rng.choice(step_r20_choices)
        d["preview_stepup_rsi_min"] = rng.choice(step_rsi_choices)
        d["preview_stepup_mfi_min"] = rng.choice(step_mfi_choices)
        d["preview_atr_max"] = rng.choice(atr_choices)
        d["preview_vol_max"] = rng.choice(vmax_choices)
        d["preview_dd_floor"] = rng.choice(dd_choices)
        d["filter_count"] = _filter_count(d)
        key = tuple(d[k] for k in (
            "preview_overlay_mode", "preview_min_post_crash_days", "preview_qld_max_delever_age",
            "preview_qld3_roc5_min", "preview_qld3_roc10_min", "preview_qld3_roc20_min",
            "preview_qld3_rsi_min", "preview_qld3_mfi_min", "preview_qld2_trend_gap_floor",
            "preview_qld2_mom_gap_floor", "preview_qld2_roc5_min", "preview_qld2_roc10_min",
            "preview_qld2_roc20_min", "preview_qld2_rsi_min", "preview_qld2_mfi_min",
            "preview_stepup_score_min", "preview_stepup_roc5_min", "preview_stepup_roc10_min",
            "preview_stepup_roc20_min", "preview_stepup_rsi_min", "preview_stepup_mfi_min",
            "preview_atr_max", "preview_vol_max", "preview_dd_floor",
        ))
        if key in seen:
            continue
        seen.add(key)
        out.append(Config(**d))
    if len(out) < trials:
        raise RuntimeError(f"V21 focus generator exhausted unique combinations: requested={trials}, generated={len(out)}")
    return out

def init_worker(features: Dict[str, Any], cost_bps: float, eval_start: int) -> None:
    global _WF, _W_COST, _W_EVAL_START, _W_BENCH_LOG_CUM, _W_BENCH_FULL
    _WF = features
    _W_COST = float(cost_bps) / 10000.0
    _W_EVAL_START = int(eval_start)
    bret = features["TQQQ_O2O_RET"][_W_EVAL_START:-1]
    _W_BENCH_LOG_CUM = np.r_[0.0, np.cumsum(np.log1p(np.clip(bret, -0.999999, None)))]
    beq = np.exp(_W_BENCH_LOG_CUM)
    years = max(len(bret) / TRADING_DAYS, 1 / TRADING_DAYS)
    _W_BENCH_FULL = {
        "cagr": float(beq[-1] ** (1 / years) - 1.0),
        "mdd": float(calc_mdd(beq)),
        "total_return": float(beq[-1] - 1.0),
    }


def entry_mask_for_config(c: Config, f: Dict[str, Any]) -> np.ndarray:
    n = len(f["QQQ_CLOSE"])
    m = np.ones(n, dtype=bool)
    q = f["QQQ_CLOSE"]
    t = f["TQQQ_CLOSE"]

    if c.require_qqq_ma50: m &= q > f["QQQ_SMA50"]
    if c.require_qqq_ma200: m &= q > f["QQQ_SMA200"]
    if c.require_tqqq_ma20: m &= t > f["TQQQ_SMA20"]
    if c.require_tqqq_ma50: m &= t > f["TQQQ_SMA50"]

    if c.macd_gate in {"qqq", "both"}: m &= f["QQQ_MACD_HIST"] > 0
    if c.macd_gate in {"tqqq", "both"}: m &= f["TQQQ_MACD_HIST"] > 0

    rma_2of3 = (c.gate_logic == "rma_2of3" and c.rsi_asset != "none" and c.mfi_asset != "none" and c.atr_asset != "none")
    if c.rsi_asset != "none" and not rma_2of3: m &= f[f"{c.rsi_asset.upper()}_RSI14"] >= c.rsi_min
    if c.stoch_asset != "none":
        p = c.stoch_asset.upper(); m &= f[f"{p}_STOCH_K"] > f[f"{p}_STOCH_D"]
    if c.cci_asset != "none": m &= f[f"{c.cci_asset.upper()}_CCI20"] > 0
    if c.boll_asset != "none":
        p = c.boll_asset.upper(); close = f[f"{p}_CLOSE"]
        if c.boll_rule == "mid":
            m &= close > f[f"{p}_BOLL_MID"]
        elif c.boll_rule == "upper_half":
            m &= f[f"{p}_BOLL_POS"] >= 0.50
        else:  # not_overbought: bullish enough but not extreme upper-band chase
            m &= (f[f"{p}_BOLL_POS"] >= 0.35) & (f[f"{p}_BOLL_POS"] <= 1.10)
    if c.adx_asset != "none":
        p = c.adx_asset.upper()
        m &= (f[f"{p}_ADX14"] >= c.adx_min) & (f[f"{p}_PDI14"] > f[f"{p}_MDI14"])
    if c.mfi_asset != "none" and not rma_2of3: m &= f[f"{c.mfi_asset.upper()}_MFI14"] >= c.mfi_min
    if c.willr_asset != "none": m &= f[f"{c.willr_asset.upper()}_WILLR14"] >= c.willr_min
    if c.obv_asset != "none":
        p = c.obv_asset.upper(); m &= f[f"{p}_OBV"] > f[f"{p}_OBV_MA20"]
    if c.donchian_asset != "none": m &= f[f"{c.donchian_asset.upper()}_DONCHIAN_POS20"] >= c.donchian_min_pos
    if c.atr_asset != "none" and not rma_2of3: m &= f[f"{c.atr_asset.upper()}_ATRPCT14"] <= c.atr_max_pct
    if rma_2of3:
        r_ok = f[f"{c.rsi_asset.upper()}_RSI14"] >= c.rsi_min
        m_ok = f[f"{c.mfi_asset.upper()}_MFI14"] >= c.mfi_min
        a_ok = f[f"{c.atr_asset.upper()}_ATRPCT14"] <= c.atr_max_pct
        m &= (r_ok.astype(np.int8) + m_ok.astype(np.int8) + a_ok.astype(np.int8)) >= 2
    if c.roc_asset != "none": m &= f[f"{c.roc_asset.upper()}_ROC{c.roc_days}"] >= c.roc_min

    # NaN comparisons above naturally become False in most cases; force finite behavior.
    return np.nan_to_num(m, nan=False).astype(bool)


def simulate_strategy(c: Config, f: Dict[str, Any], cost_per_side: float) -> Tuple[np.ndarray, np.ndarray, Dict[str, Any]]:
    n = len(f["QQQ_CLOSE"])
    q = f["QQQ_CLOSE"]
    sma = f[f"QQQ_SMA{c.sma_days}"]
    mom = f[f"QQQ_ROC{c.mom_days}"]
    vol = f[f"QQQ_VOL{c.vol_days}"]
    ema_fast = f[f"QQQ_EMA{c.fast_ema_days}"]
    fmom = f[f"QQQ_ROC{c.fast_mom_days}"]
    dd252 = f["QQQ_DD252"]

    valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(ema_fast) & np.isfinite(fmom) & np.isfinite(dd252)
    trend_ok = q > sma
    mom_ok = mom > c.mom_threshold
    vol_ok = vol < c.vol_threshold
    trend_gap = np.divide(q, sma, out=np.full(n, np.nan, dtype=float), where=np.isfinite(sma) & (sma != 0)) - 1.0
    mom_gap = mom - c.mom_threshold
    score = trend_ok.astype(np.int8) + mom_ok.astype(np.int8) + vol_ok.astype(np.int8)
    crash = valid & (~trend_ok) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
    fast_base = valid & (q > ema_fast) & (fmom >= c.fast_threshold) & (mom > c.fast_long_mom_floor) & (~crash)
    entry_mask = entry_mask_for_config(c, f)
    fast_ok = fast_base & entry_mask

    # V3 bull-impulse override. It only changes participation; crash remains absolute priority.
    bull_override = np.zeros(n, dtype=bool)
    if c.bull_override_mode != "off":
        bull_common = valid & (~crash) & (f["QQQ_ROC20"] >= c.bull_roc20_min) & (f["QQQ_RSI14"] >= c.bull_rsi_min) & (q > f["QQQ_EMA20"])
        if c.bull_override_mode == "trend_roc":
            bull_override = bull_common & trend_ok
        else:  # ema_roc can participate before the long SMA fully recovers
            bull_override = bull_common & (dd252 > c.crash_dd)

    # V4 medium-trend bridge: deliberately excludes deep drawdown / high-vol regimes.
    bridge_ok = np.zeros(n, dtype=bool)
    if c.bridge_mode != "off":
        bridge_ok = valid & (~crash) & (q > f["QQQ_SMA50"]) & (q > f["QQQ_EMA20"]) & (f["QQQ_ROC60"] >= c.bridge_roc60_min) & (f["QQQ_RSI14"] >= c.bridge_rsi_min) & (dd252 >= c.bridge_dd_floor) & (vol <= c.bridge_vol_max)

    states = np.zeros(n, dtype=np.int8)  # underlying control state at each session OPEN
    shadow_exec = np.full(n, -1, dtype=np.int8)  # V20 execution-only one-session overlay
    down_count = 0
    state = 0
    days_since_crash = 10**9
    days_since_delever = 10**9
    days_since_correction = 10**9
    crash_streak = 0
    last_crash_streak = 0
    crash_trough_dd = np.nan
    staged_qld_age = -1
    explosive_latch = 0
    dd_rebound_latch = 0
    # V11 separate trough state; legacy crash_trough_dd above is intentionally untouched.
    tr2_trough_dd = np.nan
    tr2_trough_idx = -1
    tr2_episode_crash_days = 0
    tr2_latch = 0
    tr2_qld_armed = False
    tr2_qld_age = -1
    tr2_new_lows = 0
    transitions = 0
    crash_exits = 0
    fast_entries = 0
    defense_events = 0
    staged_entries = 0
    staged_promotions = 0
    explosive_entries = 0
    correction_entries = 0
    correction_tqqq_entries = 0
    dd_rebound_entries = 0
    trough_recovery_entries = 0
    trough_recovery_qld_entries = 0
    trough_recovery_promotions = 0
    trough_recovery_hold_days = 0
    trough_split_qld_days = 0
    trough_split_promotions = 0
    trough_split_aborts = 0
    repromote_guard_blocks = 0
    repromote_guard_passes = 0
    repromote_guard_regime_days = 0
    soft_vol_signal_days = 0
    soft_vol_entries = 0
    soft_vol_streak = 0
    impulse_veto_signal_days = 0
    impulse_veto_blocks = 0
    impulse_veto_score1_blocks = 0
    impulse_veto_score2_blocks = 0
    impulse_veto_run = 0
    precision_veto_signal_days = 0
    precision_veto_blocks = 0
    precision_veto_trend_blocks = 0
    precision_veto_mom_blocks = 0
    precision_veto_score1_blocks = 0
    precision_veto_fresh_blocks = 0
    precision_veto_run = 0
    reentry_grace_signal_days = 0
    reentry_grace_blocks = 0
    reentry_grace_mom_blocks = 0
    reentry_grace_trend_blocks = 0
    reentry_grace_score2_blocks = 0
    reentry_grace_run = 0
    micro_rescue_signal_days = 0
    micro_rescue_blocks = 0
    micro_rescue_score1_pullback_blocks = 0
    micro_rescue_score1_neartrend_blocks = 0
    micro_rescue_mature_mom_blocks = 0
    shadow_rescue_signal_days = 0
    shadow_rescue_blocks = 0
    shadow_rescue_score1_blocks = 0
    shadow_rescue_mature_mom_blocks = 0
    preview_overlay_signal_days = 0
    preview_overlay_blocks = 0
    preview_overlay_qld3_blocks = 0
    preview_overlay_qld2_blocks = 0
    preview_overlay_stepup_blocks = 0
    days_since_tqqq_reentry = 10**9
    last_any_veto_t = -10**9
    ordinary_delever_history: List[int] = []
    last_ordinary_delever_t = -10**9
    adaptive_hold_days = 0
    adaptive_exit_events = 0

    # Close[t] decides state at Open[t+1].
    for t in range(n - 1):
        cur = state
        nxt = cur
        staged_fired = False
        explosive_fired = False
        trough_fired = False
        if not valid[t]:
            states[t + 1] = state
            continue
        if days_since_tqqq_reentry < 10**9:
            days_since_tqqq_reentry = min(days_since_tqqq_reentry + 1, 10**9)

        # V8 explosive reversal predicates. Recovery-specific ATR is deliberately
        # independent from the normal ATR8 entry gate.
        exp_trigger_common = (
            c.explosive_rebound_mode != "off"
            and q[t] > f["QQQ_EMA5"][t]
            and f["QQQ_ROC5"][t] >= c.explosive_roc5_min
            and f["QQQ_ROC10"][t] >= c.explosive_roc10_min
            and f["QQQ_RSI14"][t] >= c.explosive_rsi_min
            and f["QQQ_MFI14"][t] >= c.explosive_mfi_min
            and f["QQQ_ATRPCT14"][t] <= c.explosive_atr_max
            and vol[t] <= c.explosive_vol_max
        )
        exp_hold_common = (
            q[t] > f["QQQ_EMA5"][t]
            and f["QQQ_ROC5"][t] >= c.explosive_hold_roc5_min
            and f["QQQ_RSI14"][t] >= max(40, c.explosive_rsi_min - 10)
            and f["QQQ_MFI14"][t] >= max(30, c.explosive_mfi_min - 10)
            and f["QQQ_ATRPCT14"][t] <= c.explosive_atr_max * 1.20
            and vol[t] <= c.explosive_vol_max * 1.15
        )
        ddr_trigger_common = (
            c.dd_rebound_mode != "off"
            and np.isfinite(crash_trough_dd)
            and (dd252[t] - crash_trough_dd) >= c.dd_rebound_min
            and q[t] > f["QQQ_EMA5"][t]
            and f["QQQ_ROC5"][t] >= c.dd_rebound_roc5_min
            and f["QQQ_RSI14"][t] >= c.dd_rebound_rsi_min
            and f["QQQ_MFI14"][t] >= c.dd_rebound_mfi_min
            and f["QQQ_ATRPCT14"][t] <= c.dd_rebound_atr_max
            and vol[t] <= c.dd_rebound_vol_max
        )
        ddr_hold_common = (
            (c.dd_rebound_mode != "off" or dd_rebound_latch > 0)
            and q[t] > f["QQQ_EMA5"][t]
            and f["QQQ_ROC5"][t] >= c.dd_rebound_hold_roc5_min
            and f["QQQ_RSI14"][t] >= max(38, c.dd_rebound_rsi_min - 10)
            and f["QQQ_MFI14"][t] >= max(28, c.dd_rebound_mfi_min - 10)
            and f["QQQ_ATRPCT14"][t] <= c.dd_rebound_atr_max * 1.25
            and vol[t] <= c.dd_rebound_vol_max * 1.20
        )

        if crash[t]:
            crash_streak += 1
            last_crash_streak = crash_streak
            days_since_crash = 0
            days_since_correction = 10**9
            crash_trough_dd = dd252[t] if not np.isfinite(crash_trough_dd) else min(crash_trough_dd, dd252[t])

            # V11 DDR2: track a separate trough and reset all recovery confirmation
            # whenever a fresh low is observed. This prevents a mid-crash bounce from
            # inheriting confirmation from an older trough.
            if c.trough_recovery_mode != "off":
                tr2_episode_crash_days += 1
                new_tr2_low = (not np.isfinite(tr2_trough_dd)) or (dd252[t] < tr2_trough_dd - 1e-12)
                if new_tr2_low:
                    tr2_trough_dd = dd252[t]
                    tr2_trough_idx = t
                    tr2_latch = 0
                    tr2_qld_armed = False
                    tr2_qld_age = -1
                    tr2_new_lows += 1
                tr2_age = t - tr2_trough_idx if tr2_trough_idx >= 0 else -1
                tr2_rebound = dd252[t] - tr2_trough_dd if np.isfinite(tr2_trough_dd) else -np.inf
                tr2_trigger = bool(
                    cur < 3
                    and tr2_age >= c.trough_age_min
                    and tr2_episode_crash_days >= c.trough_min_crash_days
                    and tr2_rebound >= c.trough_rebound_min
                    and q[t] > f["QQQ_EMA5"][t]
                    and f["QQQ_ROC5"][t] >= c.trough_roc5_min
                    and f["QQQ_ROC10"][t] >= c.trough_roc10_min
                    and f["QQQ_RSI14"][t] >= c.trough_rsi_min
                    and f["QQQ_MFI14"][t] >= c.trough_mfi_min
                    and f["QQQ_ATRPCT14"][t] <= c.trough_atr_max
                    and vol[t] <= c.trough_vol_max
                )
                tr2_hold_ok = bool(
                    tr2_age >= c.trough_age_min
                    and q[t] > f["QQQ_EMA5"][t]
                    and f["QQQ_ROC5"][t] >= c.trough_hold_roc5_min
                    and f["QQQ_ROC10"][t] >= c.trough_hold_roc10_min
                    and f["QQQ_RSI14"][t] >= max(38, c.trough_rsi_min - 10)
                    and f["QQQ_MFI14"][t] >= max(28, c.trough_mfi_min - 10)
                    and f["QQQ_ATRPCT14"][t] <= c.trough_atr_max * 1.25
                    and vol[t] <= c.trough_vol_max * 1.20
                )
                if tr2_latch > 0 and tr2_hold_ok:
                    nxt = 3
                    tr2_latch -= 1
                    trough_recovery_hold_days += 1
                    trough_fired = True
                    explosive_fired = True
                    down_count = 0
                elif tr2_qld_armed:
                    tr2_qld_age += 1
                    if c.trough_recovery_mode == "qld_split3":
                        split_promote = bool(
                            tr2_qld_age >= max(1, c.trough_confirm_days)
                            and tr2_qld_age <= max(c.trough_confirm_days, c.trough_bridge_max_days)
                            and tr2_rebound >= c.trough_promote_rebound_min
                            and q[t] > f["QQQ_EMA5"][t]
                            and f["QQQ_ROC5"][t] >= c.trough_promote_roc5_min
                            and f["QQQ_ROC10"][t] >= c.trough_promote_roc10_min
                            and f["QQQ_RSI14"][t] >= c.trough_promote_rsi_min
                            and f["QQQ_MFI14"][t] >= c.trough_promote_mfi_min
                            and f["QQQ_ATRPCT14"][t] <= c.trough_promote_atr_max
                            and vol[t] <= c.trough_promote_vol_max
                        )
                        split_hold = bool(
                            tr2_qld_age <= c.trough_bridge_max_days
                            and tr2_rebound >= max(0.03, c.trough_rebound_min - 0.06)
                            and f["QQQ_ROC10"][t] >= c.trough_hold_roc10_min
                            and f["QQQ_RSI14"][t] >= max(40, c.trough_rsi_min - 8)
                            and f["QQQ_ATRPCT14"][t] <= c.trough_atr_max * 1.25
                            and vol[t] <= c.trough_vol_max * 1.20
                        )
                        if split_promote:
                            nxt = 3
                            tr2_latch = max(0, c.trough_hold_days - 1)
                            tr2_qld_armed = False
                            trough_recovery_promotions += 1
                            trough_split_promotions += 1
                            fast_entries += int(cur < 3)
                        elif split_hold:
                            nxt = max(cur, 2)
                            trough_split_qld_days += 1
                        else:
                            tr2_qld_armed = False
                            trough_split_aborts += 1
                            nxt = 0
                    else:
                        if tr2_trigger and tr2_qld_age >= max(1, c.trough_confirm_days):
                            nxt = 3
                            tr2_latch = max(0, c.trough_hold_days - 1)
                            tr2_qld_armed = False
                            trough_recovery_promotions += 1
                            fast_entries += int(cur < 3)
                        elif tr2_qld_age <= max(1, c.trough_confirm_days) + 2:
                            nxt = max(cur, 2)
                        else:
                            tr2_qld_armed = False
                            nxt = 0
                    trough_fired = True
                    explosive_fired = True
                    down_count = 0
                elif tr2_trigger:
                    if c.trough_recovery_mode in {"qld_confirm3", "qld_split3"}:
                        nxt = max(cur, 2)
                        tr2_qld_armed = True
                        tr2_qld_age = 0
                        if nxt > cur:
                            trough_recovery_qld_entries += 1
                            trough_recovery_entries += 1
                    elif c.trough_recovery_mode == "gated3":
                        nxt = 3 if bool(entry_mask[t]) else max(cur, 2)
                        if nxt > cur:
                            trough_recovery_entries += 1
                            trough_recovery_qld_entries += int(nxt == 2)
                    else:
                        nxt = 3
                        if nxt > cur:
                            trough_recovery_entries += 1
                    tr2_latch = max(0, c.trough_hold_days - 1) if nxt == 3 else 0
                    trough_fired = True
                    explosive_fired = True
                    down_count = 0
                    if nxt > cur:
                        fast_entries += 1

            # V9 DD-rebound path gets first priority inside/post a true crash. It
            # reacts to recovery from the observed trough rather than a fixed date.
            ddr_trigger_common = (
                c.dd_rebound_mode != "off"
                and np.isfinite(crash_trough_dd)
                and (dd252[t] - crash_trough_dd) >= c.dd_rebound_min
                and q[t] > f["QQQ_EMA5"][t]
                and f["QQQ_ROC5"][t] >= c.dd_rebound_roc5_min
                and f["QQQ_RSI14"][t] >= c.dd_rebound_rsi_min
                and f["QQQ_MFI14"][t] >= c.dd_rebound_mfi_min
                and f["QQQ_ATRPCT14"][t] <= c.dd_rebound_atr_max
                and vol[t] <= c.dd_rebound_vol_max
            )
            if dd_rebound_latch > 0 and ddr_hold_common:
                nxt = 3
                dd_rebound_latch -= 1
                explosive_fired = True
                down_count = 0
            elif ddr_trigger_common:
                if c.dd_rebound_mode == "qld":
                    nxt = max(cur, 2)
                elif c.dd_rebound_mode == "gated3":
                    nxt = 3 if bool(entry_mask[t]) else max(cur, 2)
                else:
                    nxt = 3
                dd_rebound_latch = max(0, c.dd_rebound_hold_days - 1) if nxt == 3 else 0
                explosive_fired = True
                down_count = 0
                if nxt > cur:
                    fast_entries += 1
                    dd_rebound_entries += 1

            # V8 explosive path remains for exact backward-compatible anchors. Once
            # triggered, the short latch avoids an immediate return to CASH merely
            # because the 180-day trend / volatility crash flag is slow to clear.
            if not explosive_fired:
                if explosive_latch > 0 and exp_hold_common:
                    nxt = 3
                    explosive_latch -= 1
                    explosive_fired = True
                    down_count = 0
                elif (
                    c.explosive_rebound_mode != "off"
                    and crash_streak >= c.explosive_min_crash_days
                    and exp_trigger_common
                ):
                    if c.explosive_rebound_mode == "direct3":
                        nxt = 3
                    else:  # gated3: conservative control arm
                        nxt = 3 if bool(entry_mask[t]) else 2
                    explosive_latch = max(0, c.explosive_hold_days - 1) if nxt == 3 else 0
                    explosive_fired = True
                    down_count = 0
                    if nxt > cur:
                        fast_entries += 1
                        if nxt == 3:
                            explosive_entries += 1

            if not explosive_fired:
                # V7 early crash release retained for backward-compatible anchors.
                cr_common = (
                    c.crash_release_mode != "off"
                    and crash_streak >= c.crash_release_wait
                    and dd252[t] >= c.crash_release_dd_floor
                    and q[t] > f["QQQ_EMA5"][t]
                    and f["QQQ_ROC5"][t] >= c.crash_release_roc5_min
                    and f["QQQ_ROC10"][t] >= c.crash_release_roc10_min
                    and f["QQQ_RSI14"][t] >= c.crash_release_rsi_min
                    and f["QQQ_MFI14"][t] >= c.crash_release_mfi_min
                    and vol[t] <= c.crash_release_vol_max
                )
                if cr_common:
                    if c.crash_release_mode == "qld":
                        nxt = 2
                    elif c.crash_release_mode == "gated3":
                        nxt = 3 if bool(entry_mask[t]) else 2
                    else:  # direct3
                        nxt = 3
                    down_count = 0
                    if nxt > cur:
                        fast_entries += 1
                else:
                    nxt = 0
                    down_count = 0
                    if cur != 0:
                        crash_exits += 1
        else:
            crash_streak = 0
            days_since_crash = min(days_since_crash + 1, 10**9)
            fast_fired = False

            # V11 DDR2 continues for a bounded post-crash window. A fresh lower DD,
            # even on a day when the legacy crash flag briefly clears, resets confirmation.
            if c.trough_recovery_mode != "off":
                if days_since_crash > c.trough_recovery_window_days and tr2_latch <= 0:
                    tr2_trough_dd = np.nan
                    tr2_trough_idx = -1
                    tr2_episode_crash_days = 0
                    tr2_qld_armed = False
                    tr2_qld_age = -1
                elif np.isfinite(tr2_trough_dd):
                    new_tr2_low = dd252[t] < tr2_trough_dd - 1e-12
                    if new_tr2_low:
                        tr2_trough_dd = dd252[t]
                        tr2_trough_idx = t
                        tr2_latch = 0
                        tr2_qld_armed = False
                        tr2_qld_age = -1
                        tr2_new_lows += 1
                    tr2_age = t - tr2_trough_idx if tr2_trough_idx >= 0 else -1
                    tr2_rebound = dd252[t] - tr2_trough_dd
                    tr2_trigger = bool(
                        cur < 3
                        and days_since_crash <= c.trough_recovery_window_days
                        and tr2_age >= c.trough_age_min
                        and tr2_episode_crash_days >= c.trough_min_crash_days
                        and tr2_rebound >= c.trough_rebound_min
                        and q[t] > f["QQQ_EMA5"][t]
                        and f["QQQ_ROC5"][t] >= c.trough_roc5_min
                        and f["QQQ_ROC10"][t] >= c.trough_roc10_min
                        and f["QQQ_RSI14"][t] >= c.trough_rsi_min
                        and f["QQQ_MFI14"][t] >= c.trough_mfi_min
                        and f["QQQ_ATRPCT14"][t] <= c.trough_atr_max
                        and vol[t] <= c.trough_vol_max
                    )
                    tr2_hold_ok = bool(
                        tr2_age >= c.trough_age_min
                        and q[t] > f["QQQ_EMA5"][t]
                        and f["QQQ_ROC5"][t] >= c.trough_hold_roc5_min
                        and f["QQQ_ROC10"][t] >= c.trough_hold_roc10_min
                        and f["QQQ_RSI14"][t] >= max(38, c.trough_rsi_min - 10)
                        and f["QQQ_MFI14"][t] >= max(28, c.trough_mfi_min - 10)
                        and f["QQQ_ATRPCT14"][t] <= c.trough_atr_max * 1.25
                        and vol[t] <= c.trough_vol_max * 1.20
                    )
                    if tr2_latch > 0 and tr2_hold_ok:
                        nxt = 3
                        tr2_latch -= 1
                        trough_recovery_hold_days += 1
                        trough_fired = True
                        explosive_fired = True
                        down_count = 0
                    elif tr2_qld_armed:
                        tr2_qld_age += 1
                        if c.trough_recovery_mode == "qld_split3":
                            split_promote = bool(
                                tr2_qld_age >= max(1, c.trough_confirm_days)
                                and tr2_qld_age <= max(c.trough_confirm_days, c.trough_bridge_max_days)
                                and tr2_rebound >= c.trough_promote_rebound_min
                                and q[t] > f["QQQ_EMA5"][t]
                                and f["QQQ_ROC5"][t] >= c.trough_promote_roc5_min
                                and f["QQQ_ROC10"][t] >= c.trough_promote_roc10_min
                                and f["QQQ_RSI14"][t] >= c.trough_promote_rsi_min
                                and f["QQQ_MFI14"][t] >= c.trough_promote_mfi_min
                                and f["QQQ_ATRPCT14"][t] <= c.trough_promote_atr_max
                                and vol[t] <= c.trough_promote_vol_max
                            )
                            split_hold = bool(
                                tr2_qld_age <= c.trough_bridge_max_days
                                and tr2_rebound >= max(0.03, c.trough_rebound_min - 0.06)
                                and f["QQQ_ROC10"][t] >= c.trough_hold_roc10_min
                                and f["QQQ_RSI14"][t] >= max(40, c.trough_rsi_min - 8)
                                and f["QQQ_ATRPCT14"][t] <= c.trough_atr_max * 1.25
                                and vol[t] <= c.trough_vol_max * 1.20
                            )
                            if split_promote:
                                nxt = 3
                                tr2_latch = max(0, c.trough_hold_days - 1)
                                tr2_qld_armed = False
                                trough_recovery_promotions += 1
                                trough_split_promotions += 1
                                fast_entries += int(cur < 3)
                            elif split_hold:
                                nxt = max(cur, 2)
                                trough_split_qld_days += 1
                            else:
                                tr2_qld_armed = False
                                trough_split_aborts += 1
                        else:
                            if tr2_trigger and tr2_qld_age >= max(1, c.trough_confirm_days):
                                nxt = 3
                                tr2_latch = max(0, c.trough_hold_days - 1)
                                tr2_qld_armed = False
                                trough_recovery_promotions += 1
                                fast_entries += int(cur < 3)
                            elif tr2_qld_age <= max(1, c.trough_confirm_days) + 2:
                                nxt = max(cur, 2)
                            else:
                                tr2_qld_armed = False
                        trough_fired = True
                        explosive_fired = True
                        down_count = 0
                    elif tr2_trigger:
                        if c.trough_recovery_mode in {"qld_confirm3", "qld_split3"}:
                            nxt = max(cur, 2)
                            tr2_qld_armed = True
                            tr2_qld_age = 0
                            if nxt > cur:
                                trough_recovery_qld_entries += 1
                                trough_recovery_entries += 1
                        elif c.trough_recovery_mode == "gated3":
                            nxt = 3 if bool(entry_mask[t]) else max(cur, 2)
                            if nxt > cur:
                                trough_recovery_entries += 1
                                trough_recovery_qld_entries += int(nxt == 2)
                        else:
                            nxt = 3
                            if nxt > cur:
                                trough_recovery_entries += 1
                        tr2_latch = max(0, c.trough_hold_days - 1) if nxt == 3 else 0
                        trough_fired = True
                        explosive_fired = True
                        down_count = 0
                        if nxt > cur:
                            fast_entries += 1

            # V9 trough-rebound release. Keep the crash trough only for a bounded
            # generic post-crash window, then forget it completely.
            if days_since_crash > c.dd_rebound_window_days and dd_rebound_latch <= 0:
                crash_trough_dd = np.nan
            if dd_rebound_latch > 0:
                if ddr_hold_common:
                    nxt = 3
                    dd_rebound_latch -= 1
                    explosive_fired = True
                    down_count = 0
                else:
                    dd_rebound_latch = 0
            if (
                not explosive_fired
                and c.dd_rebound_mode != "off"
                and days_since_crash <= c.dd_rebound_window_days
                and ddr_trigger_common
            ):
                if c.dd_rebound_mode == "qld":
                    nxt = max(cur, 2)
                elif c.dd_rebound_mode == "gated3":
                    nxt = 3 if bool(entry_mask[t]) else max(cur, 2)
                else:
                    nxt = 3
                dd_rebound_latch = max(0, c.dd_rebound_hold_days - 1) if nxt == 3 else 0
                explosive_fired = True
                down_count = 0
                if nxt > cur:
                    fast_entries += 1
                    dd_rebound_entries += 1

            # V8 explosive latch / post-crash trigger retained for old anchors.
            # for a short period after a genuine crash; no calendar dates are used.
            if (not explosive_fired) and explosive_latch > 0:
                if exp_hold_common:
                    nxt = 3
                    explosive_latch -= 1
                    explosive_fired = True
                    down_count = 0
                else:
                    explosive_latch = 0
            if (
                not explosive_fired
                and c.explosive_rebound_mode != "off"
                and days_since_crash <= c.explosive_window_days
                and last_crash_streak >= c.explosive_min_crash_days
                and exp_trigger_common
            ):
                if c.explosive_rebound_mode == "direct3":
                    nxt = 3
                else:
                    nxt = 3 if bool(entry_mask[t]) else 2
                explosive_latch = max(0, c.explosive_hold_days - 1) if nxt == 3 else 0
                explosive_fired = True
                down_count = 0
                if nxt > cur:
                    fast_entries += 1
                    if nxt == 3:
                        explosive_entries += 1

            # V3 crash-recovery participation: retained for old seeds.
            recovery_fired = False
            if (not explosive_fired) and c.recovery_mode != "off" and days_since_crash <= c.recovery_days:
                rec_fast = (q[t] > ema_fast[t]) and (fmom[t] >= c.fast_threshold) and (mom[t] > c.recovery_mom_floor) and (vol[t] <= c.recovery_vol_max)
                rec_strong = (q[t] > f["QQQ_EMA20"][t]) and (f["QQQ_ROC20"][t] >= c.recovery_roc20_min) and (f["QQQ_RSI14"][t] >= 50) and (vol[t] <= c.recovery_vol_max)
                if c.recovery_mode == "relaxed_fast":
                    recovery_fired = bool(rec_fast)
                elif c.recovery_mode == "strong_roc":
                    recovery_fired = bool(rec_strong)
                else:
                    recovery_fired = bool(rec_fast or rec_strong)
                if recovery_fired:
                    nxt = 3
                    down_count = 0
                    if nxt > cur:
                        fast_entries += 1

            # V4 explosive post-crash rebound bypass retained for old seeds.
            rebound_fired = False
            if (not explosive_fired) and (not recovery_fired) and c.rebound_bypass_mode != "off" and days_since_crash <= c.rebound_window_days:
                rb_common = (f["QQQ_RSI14"][t] >= c.rebound_rsi_min) and (vol[t] <= c.rebound_vol_max)
                rb10 = rb_common and (q[t] > f["QQQ_EMA5"][t]) and (f["QQQ_ROC10"][t] >= c.rebound_roc10_min)
                rb20 = rb_common and (q[t] > f["QQQ_EMA10"][t]) and (f["QQQ_ROC20"][t] >= c.rebound_roc20_min)
                if c.rebound_bypass_mode == "roc10":
                    rebound_fired = bool(rb10)
                elif c.rebound_bypass_mode == "roc20":
                    rebound_fired = bool(rb20)
                else:
                    rebound_fired = bool(rb10 or rb20)
                if rebound_fired:
                    nxt = 3
                    down_count = 0
                    if nxt > cur:
                        fast_entries += 1

            # V9 ordinary-correction re-entry. Only non-crash deleveraging arms
            # days_since_correction, so this cannot weaken deep-crash protection.
            correction_fired = False
            if (
                (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and c.correction_reentry_mode != "off"
                and cur < 3
                and days_since_correction <= c.correction_window_days
            ):
                corr_common = (
                    q[t] > f[f"QQQ_EMA{c.correction_ema_days}"][t]
                    and f["QQQ_ROC5"][t] >= c.correction_roc5_min
                    and f["QQQ_ROC10"][t] >= c.correction_roc10_min
                    and f["QQQ_RSI14"][t] >= c.correction_rsi_min
                    and f["QQQ_MFI14"][t] >= c.correction_mfi_min
                    and f["QQQ_ATRPCT14"][t] <= c.correction_atr_max
                    and vol[t] <= c.correction_vol_max
                    and dd252[t] >= c.correction_dd_floor
                )
                if bool(corr_common):
                    if c.correction_reentry_mode == "qld":
                        nxt = max(cur, 2)
                    elif c.correction_reentry_mode == "gated3":
                        nxt = 3 if bool(entry_mask[t]) else max(cur, 2)
                    else:
                        nxt = 3
                    correction_fired = nxt > cur
                    if correction_fired:
                        down_count = 0
                        fast_entries += 1
                        correction_entries += 1
                        if nxt == 3:
                            correction_tqqq_entries += 1
                        # Reuse the high-priority recovery guard so lower paths do
                        # not overwrite this decision in the same session.
                        rebound_fired = True

            # V8 staged generic re-entry: after any true 3x deleveraging, recover
            # through QLD first. Promotion to TQQQ requires a second, stronger signal.
            if (
                (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and c.staged_reentry_mode != "off"
                and cur < 3
                and days_since_delever <= c.staged_window_days
            ):
                stage_common = (
                    q[t] > f[f"QQQ_EMA{c.staged_ema_days}"][t]
                    and f["QQQ_ROC10"][t] >= c.staged_roc10_min
                    and f["QQQ_RSI14"][t] >= c.staged_rsi_min
                    and f["QQQ_MFI14"][t] >= c.staged_mfi_min
                    and f["QQQ_ATRPCT14"][t] <= c.staged_atr_max
                    and vol[t] <= c.staged_vol_max
                    and dd252[t] >= c.staged_dd_floor
                )
                if bool(stage_common):
                    staged_fired = True
                    if cur < 2:
                        nxt = 2
                        down_count = 0
                        staged_entries += 1
                        fast_entries += 1
                    else:  # already QLD: wait, then require stronger promotion confirmation
                        promote_common = (
                            q[t] > f["QQQ_EMA5"][t]
                            and f["QQQ_ROC5"][t] >= c.staged_promote_roc5_min
                            and f["QQQ_ROC10"][t] >= c.staged_promote_roc10_min
                            and f["QQQ_RSI14"][t] >= c.staged_promote_rsi_min
                            and f["QQQ_MFI14"][t] >= c.staged_promote_mfi_min
                            and f["QQQ_ATRPCT14"][t] <= c.staged_promote_atr_max
                            and vol[t] <= c.staged_vol_max
                        )
                        waited_enough = staged_qld_age >= max(1, c.staged_promote_wait)
                        gate_ok = True if c.staged_reentry_mode == "qld_confirm3" else bool(entry_mask[t])
                        if waited_enough and bool(promote_common) and gate_ok:
                            nxt = 3
                            down_count = 0
                            staged_promotions += 1
                            fast_entries += 1
                        else:
                            nxt = 2
                            down_count = 0

            # V6 snapback remains as a separate fallback / anchor behavior.
            snapback_fired = False
            if (not explosive_fired) and (not recovery_fired) and (not rebound_fired) and (not staged_fired) and c.snapback_mode != "off" and cur < 3 and days_since_delever <= c.snapback_days:
                sb_common = (q[t] > f[f"QQQ_EMA{c.snapback_ema_days}"][t]) and (f["QQQ_ROC10"][t] >= c.snapback_roc10_min) and (f["QQQ_RSI14"][t] >= c.snapback_rsi_min) and (vol[t] <= c.snapback_vol_max) and (dd252[t] >= c.snapback_dd_floor)
                if bool(sb_common):
                    if c.snapback_mode == "qld":
                        nxt = max(cur, 2)
                    elif c.snapback_mode == "gated3":
                        nxt = 3 if bool(entry_mask[t]) else max(cur, 2)
                    else:  # direct3
                        nxt = 3
                    snapback_fired = nxt > cur
                    if snapback_fired:
                        down_count = 0
                        fast_entries += 1

            if (not explosive_fired) and (not recovery_fired) and (not rebound_fired) and (not staged_fired) and (not snapback_fired) and c.fast_mode != "normal" and bool(fast_base[t]):
                if c.fast_mode == "direct" and bool(entry_mask[t]):
                    nxt = 3; fast_fired = True
                elif c.fast_mode == "qld":
                    nxt = max(cur, 2); fast_fired = True
                elif c.fast_mode == "conditional":
                    # Strong confirmation -> TQQQ, otherwise capture rebound at QLD.
                    nxt = 3 if bool(entry_mask[t]) else max(cur, 2); fast_fired = True
                if fast_fired:
                    down_count = 0
                    if nxt > cur:
                        fast_entries += 1

            # V14 soft-vol boundary rescue. Promotion-only: it never delays TQQQ exits.
            # The global core vol threshold stays fixed; only QLD -> TQQQ can be rescued
            # when volatility is the sole failed core vote and extra strength checks pass.
            if c.soft_vol_mode != "off":
                crash_exempt_window = max(63, c.trough_recovery_window_days, c.dd_rebound_window_days, c.explosive_window_days)
                recent_ok = (c.soft_vol_mode == "score2_strong3") or (0 <= days_since_delever <= c.soft_vol_recent_delever_days)
                soft_vol_signal = bool(
                    cur == 2 and (not bool(crash[t])) and days_since_crash > crash_exempt_window
                    and bool(trend_ok[t]) and bool(mom_ok[t]) and (not bool(vol_ok[t]))
                    and vol[t] > c.vol_threshold and vol[t] <= c.soft_vol_max
                    and q[t] > f["QQQ_EMA20"][t]
                    and f["QQQ_ROC10"][t] >= c.soft_vol_roc10_min
                    and f["QQQ_ROC20"][t] >= c.soft_vol_roc20_min
                    and f["QQQ_ROC120"][t] >= c.soft_vol_roc120_min
                    and f["QQQ_RSI14"][t] >= c.soft_vol_rsi_min
                    and f["QQQ_MFI14"][t] >= c.soft_vol_mfi_min
                    and dd252[t] >= c.soft_vol_dd_floor
                    and bool(entry_mask[t]) and recent_ok
                )
                if soft_vol_signal:
                    soft_vol_signal_days += 1
                    soft_vol_streak += 1
                else:
                    soft_vol_streak = 0
                if (not explosive_fired) and (not recovery_fired) and (not rebound_fired) and (not staged_fired) and (not snapback_fired) and (not fast_fired) and soft_vol_signal and soft_vol_streak >= max(1, c.soft_vol_confirm_days):
                    nxt = 3
                    down_count = 0
                    soft_vol_entries += 1
                    soft_vol_streak = 0
                    fast_fired = True
            else:
                soft_vol_streak = 0

            if (not explosive_fired) and (not recovery_fired) and (not rebound_fired) and (not staged_fired) and (not snapback_fired) and (not fast_fired) and bool(bull_override[t]):
                # Strong impulse gets 3x when the ordinary TQQQ gate agrees; otherwise at least QLD.
                nxt = 3 if (bool(entry_mask[t]) or c.gate_scope == "fast_only") else max(cur, 2)
                down_count = 0
                if nxt > cur:
                    fast_entries += 1
                fast_fired = True

            if (not explosive_fired) and (not recovery_fired) and (not rebound_fired) and (not staged_fired) and (not snapback_fired) and (not fast_fired):
                # V4 medium-trend re-entry bridge before ordinary one-step state logic.
                bridge_reentry = False
                if bool(bridge_ok[t]) and c.bridge_mode in {"qld", "gated3", "direct3", "hold_gated3", "hold_direct3"}:
                    if c.bridge_mode == "qld":
                        nxt = max(cur, 2)
                    elif c.bridge_mode in {"gated3", "hold_gated3"}:
                        nxt = 3 if bool(entry_mask[t]) else max(cur, 2)
                    else:
                        nxt = 3
                    bridge_reentry = nxt > cur
                    if bridge_reentry:
                        down_count = 0
                        fast_entries += 1

                if not bridge_reentry:
                    s = int(score[t])
                    # Hold-buffer is only allowed when the base engine still has at least 2/3 bullish votes.
                    hold_bridge = (cur == 3) and (s >= 2) and bool(bridge_ok[t]) and c.bridge_mode in {"hold", "hold_gated3", "hold_direct3"}
                    if s < cur:
                        if hold_bridge:
                            nxt = 3
                            down_count = 0
                        else:
                            down_count += 1
                            required_down = c.down_confirm
                            if c.adaptive_down_mode == "tqqq_score" and cur == 3:
                                required_down = c.soft_down_confirm if s == 2 else c.hard_down_confirm
                                if s == 2 and down_count < required_down:
                                    adaptive_hold_days += 1
                            if down_count >= required_down:
                                if c.adaptive_down_mode == "tqqq_score" and cur == 3:
                                    adaptive_exit_events += 1
                                nxt = max(0, cur - 1)
                                down_count = 0
                    elif s > cur:
                        target = min(3, cur + 1)
                        if c.gate_scope == "all_up":
                            gate_ok = bool(entry_mask[t])
                        elif c.gate_scope == "tqqq_only":
                            gate_ok = (target < 3) or bool(entry_mask[t])
                        else:
                            gate_ok = True  # fast_only
                        if gate_ok:
                            nxt = target
                        down_count = 0
                    else:
                        down_count = 0

        # V15 selective impulse exit veto. Apply only to an ordinary TQQQ->QLD
        # step-down produced by the core state engine. It cannot block a crash exit,
        # and the later immediate TQQQ-exit / defense rules remain free to override it.
        if c.impulse_veto_mode != "off":
            crash_exempt_window = max(63, c.trough_recovery_window_days, c.dd_rebound_window_days, c.explosive_window_days)
            s_now = int(score[t])
            if c.impulse_veto_mode == "score1_strong":
                score_scope_ok = s_now == 1
            elif c.impulse_veto_mode == "score2_strong":
                score_scope_ok = s_now == 2
            else:
                score_scope_ok = s_now in (1, 2)
            impulse_signal = bool(
                cur == 3 and nxt == 2
                and (not bool(crash[t])) and days_since_crash > crash_exempt_window
                and bool(entry_mask[t])
                and bool(vol_ok[t]) and score_scope_ok
                and q[t] > f[f"QQQ_EMA{c.impulse_veto_ema_days}"][t]
                and f["QQQ_ROC5"][t] >= c.impulse_veto_roc5_min
                and f["QQQ_ROC10"][t] >= c.impulse_veto_roc10_min
                and f["QQQ_ROC20"][t] >= c.impulse_veto_roc20_min
                and f["QQQ_RSI14"][t] >= c.impulse_veto_rsi_min
                and f["QQQ_MFI14"][t] >= c.impulse_veto_mfi_min
                and f["QQQ_ATRPCT14"][t] <= c.impulse_veto_atr_max
                and vol[t] <= c.impulse_veto_vol_max
                and dd252[t] >= c.impulse_veto_dd_floor
                and (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and (not staged_fired) and (not snapback_fired) and (not fast_fired)
            )
            if impulse_signal:
                impulse_veto_signal_days += 1
                if impulse_veto_run < max(1, c.impulse_veto_max_days):
                    nxt = 3
                    # Keep the ordinary down-confirm almost armed. If weakness persists
                    # tomorrow and the veto is no longer eligible, the strategy can exit
                    # immediately rather than restarting a fresh multi-day grace period.
                    down_count = max(0, c.down_confirm - 1)
                    impulse_veto_run += 1
                    impulse_veto_blocks += 1
                    impulse_veto_score1_blocks += int(s_now == 1)
                    impulse_veto_score2_blocks += int(s_now == 2)
                    last_any_veto_t = t
                else:
                    impulse_veto_run = 0
            else:
                impulse_veto_run = 0
        else:
            impulse_veto_run = 0

        # V16 precision transition veto. It is layered on top of the proven V15 anchor
        # and can only block an additional ordinary TQQQ->QLD exit. Consecutive stacking
        # with a V15 veto is forbidden, so the combined grace period stays narrow.
        if c.precision_veto_mode != "off":
            crash_exempt_window = max(63, c.trough_recovery_window_days, c.dd_rebound_window_days, c.explosive_window_days)
            s_now = int(score[t])
            if c.precision_veto_mode == "trend_near":
                structure_ok = bool(s_now == 2 and (not trend_ok[t]) and mom_ok[t] and vol_ok[t] and trend_gap[t] >= c.precision_veto_trend_gap_floor)
            elif c.precision_veto_mode == "mom_near":
                structure_ok = bool(s_now == 2 and trend_ok[t] and (not mom_ok[t]) and vol_ok[t] and mom_gap[t] >= c.precision_veto_mom_gap_floor)
            elif c.precision_veto_mode == "score1_near":
                structure_ok = bool(s_now == 1 and (not trend_ok[t]) and (not mom_ok[t]) and vol_ok[t]
                                    and trend_gap[t] >= c.precision_veto_trend_gap_floor
                                    and mom_gap[t] >= c.precision_veto_mom_gap_floor)
            else:  # fresh_reentry
                structure_ok = bool(s_now in (1,2) and s_now >= c.precision_veto_score_min and vol_ok[t]
                                    and days_since_tqqq_reentry <= c.precision_veto_recent_reentry_days
                                    and trend_gap[t] >= c.precision_veto_trend_gap_floor
                                    and mom_gap[t] >= c.precision_veto_mom_gap_floor)
            precision_signal = bool(
                cur == 3 and nxt == 2 and (t - last_any_veto_t) > 1
                and (not bool(crash[t])) and days_since_crash > crash_exempt_window
                and bool(entry_mask[t]) and structure_ok
                and q[t] > f[f"QQQ_EMA{c.precision_veto_ema_days}"][t]
                and f["QQQ_ROC5"][t] >= c.precision_veto_roc5_min
                and f["QQQ_ROC10"][t] >= c.precision_veto_roc10_min
                and f["QQQ_ROC20"][t] >= c.precision_veto_roc20_min
                and f["QQQ_RSI14"][t] >= c.precision_veto_rsi_min
                and f["QQQ_MFI14"][t] >= c.precision_veto_mfi_min
                and f["QQQ_ATRPCT14"][t] <= c.precision_veto_atr_max
                and vol[t] <= c.precision_veto_vol_max
                and dd252[t] >= c.precision_veto_dd_floor
                and (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and (not staged_fired) and (not snapback_fired) and (not fast_fired)
            )
            if precision_signal:
                precision_veto_signal_days += 1
                if precision_veto_run < max(1, c.precision_veto_max_days):
                    nxt = 3
                    down_count = max(0, c.down_confirm - 1)
                    precision_veto_run += 1
                    precision_veto_blocks += 1
                    precision_veto_trend_blocks += int(c.precision_veto_mode == "trend_near")
                    precision_veto_mom_blocks += int(c.precision_veto_mode == "mom_near")
                    precision_veto_score1_blocks += int(c.precision_veto_mode == "score1_near")
                    precision_veto_fresh_blocks += int(c.precision_veto_mode == "fresh_reentry")
                    last_any_veto_t = t
                else:
                    precision_veto_run = 0
            else:
                precision_veto_run = 0
        else:
            precision_veto_run = 0

        # V17 recent re-entry grace. This intentionally excludes score-1. It can only
        # block one ordinary TQQQ->QLD step-down after a recent TQQQ restoration when
        # exactly one core vote is missing (or generic score-2 mode). Unlike V16, short
        # ROC/EMA need not be strongly positive and the post-crash exclusion is searched.
        if c.reentry_grace_mode != "off":
            s_now = int(score[t])
            grace_branch = ""
            if c.reentry_grace_mode == "mom_reentry":
                structure_ok = bool(s_now == 2 and trend_ok[t] and (not mom_ok[t]) and vol_ok[t]
                                    and mom_gap[t] >= c.reentry_grace_mom_gap_floor)
                grace_branch = "mom" if structure_ok else ""
            elif c.reentry_grace_mode == "trend_reentry":
                structure_ok = bool(s_now == 2 and (not trend_ok[t]) and mom_ok[t] and vol_ok[t]
                                    and trend_gap[t] >= c.reentry_grace_trend_gap_floor)
                grace_branch = "trend" if structure_ok else ""
            elif c.reentry_grace_mode == "split_reentry":
                mom_branch = bool(s_now == 2 and trend_ok[t] and (not mom_ok[t]) and vol_ok[t]
                                  and mom_gap[t] >= c.reentry_grace_mom_gap_floor)
                trend_branch = bool(s_now == 2 and (not trend_ok[t]) and mom_ok[t] and vol_ok[t]
                                    and trend_gap[t] >= c.reentry_grace_trend_gap_floor)
                structure_ok = bool(mom_branch or trend_branch)
                grace_branch = "mom" if mom_branch else ("trend" if trend_branch else "")
            else:  # score2_reentry
                structure_ok = bool(s_now == 2 and vol_ok[t]
                                    and trend_gap[t] >= c.reentry_grace_trend_gap_floor
                                    and mom_gap[t] >= c.reentry_grace_mom_gap_floor)
                if structure_ok:
                    grace_branch = "mom" if (trend_ok[t] and (not mom_ok[t])) else ("trend" if ((not trend_ok[t]) and mom_ok[t]) else "score2")

            if c.reentry_grace_mode == "split_reentry" and grace_branch in {"mom", "trend"}:
                branch_gate = c.reentry_grace_mom_entry_gate if grace_branch == "mom" else c.reentry_grace_trend_entry_gate
                entry_ok = bool(entry_mask[t]) if branch_gate == "pass" else True
                roc5_floor = c.reentry_grace_mom_roc5_min if grace_branch == "mom" else c.reentry_grace_trend_roc5_min
            else:
                entry_ok = bool(entry_mask[t]) if c.reentry_grace_entry_gate == "pass" else True
                roc5_floor = c.reentry_grace_roc5_min
            ema_ok = True if c.reentry_grace_ema_days <= 0 else bool(q[t] > f[f"QQQ_EMA{c.reentry_grace_ema_days}"][t])

            # V18 permits exactly one adjacent second grace day only for split-mode Trend-fail continuation.
            # All other V15/V17 veto stacking remains prohibited on adjacent sessions.
            spacing_ok = bool((t - last_any_veto_t) > 1)
            if (c.reentry_grace_mode == "split_reentry" and reentry_grace_run > 0
                    and reentry_grace_run < max(1, c.reentry_grace_max_days)
                    and grace_branch == "trend" and (t - last_any_veto_t) == 1):
                spacing_ok = True

            grace_signal = bool(
                cur == 3 and nxt == 2 and spacing_ok
                and (not bool(crash[t]))
                and days_since_crash >= c.reentry_grace_min_post_crash_days
                and c.reentry_grace_min_age <= days_since_tqqq_reentry <= c.reentry_grace_max_age
                and structure_ok and entry_ok and ema_ok
                and f["QQQ_ROC5"][t] >= roc5_floor
                and f["QQQ_ROC10"][t] >= c.reentry_grace_roc10_min
                and f["QQQ_ROC20"][t] >= c.reentry_grace_roc20_min
                and f["QQQ_RSI14"][t] >= c.reentry_grace_rsi_min
                and f["QQQ_MFI14"][t] >= c.reentry_grace_mfi_min
                and f["QQQ_ATRPCT14"][t] <= c.reentry_grace_atr_max
                and vol[t] <= c.reentry_grace_vol_max
                and dd252[t] >= c.reentry_grace_dd_floor
                and (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and (not staged_fired) and (not snapback_fired) and (not fast_fired)
            )
            if grace_signal:
                reentry_grace_signal_days += 1
                if reentry_grace_run < max(1, c.reentry_grace_max_days):
                    nxt = 3
                    down_count = max(0, c.down_confirm - 1)
                    reentry_grace_run += 1
                    reentry_grace_blocks += 1
                    if c.reentry_grace_mode == "split_reentry":
                        reentry_grace_mom_blocks += int(grace_branch == "mom")
                        reentry_grace_trend_blocks += int(grace_branch == "trend")
                    else:
                        reentry_grace_mom_blocks += int(c.reentry_grace_mode == "mom_reentry")
                        reentry_grace_trend_blocks += int(c.reentry_grace_mode == "trend_reentry")
                        reentry_grace_score2_blocks += int(c.reentry_grace_mode == "score2_reentry")
                    last_any_veto_t = t
                else:
                    reentry_grace_run = 0
            else:
                reentry_grace_run = 0
        else:
            reentry_grace_run = 0

        # V19 sparse micro-rescue layer. Additive to the proven V18 split grace.
        # It only blocks an ordinary TQQQ->QLD step-down for one session and never
        # overrides active crash/recovery logic. All branches are state-based.
        if c.micro_rescue_mode != "off":
            s_now = int(score[t])
            roc5_now = float(f["QQQ_ROC5"][t])
            roc10_now = float(f["QQQ_ROC10"][t])
            roc20_now = float(f["QQQ_ROC20"][t])
            rsi_now = float(f["QQQ_RSI14"][t])
            mfi_now = float(f["QQQ_MFI14"][t])
            pullback_branch = bool(
                s_now == 1 and (not trend_ok[t]) and (not mom_ok[t]) and vol_ok[t]
                and 1 <= days_since_tqqq_reentry <= c.micro_rescue_recent_max_age
                and trend_gap[t] >= c.micro_rescue_score1_trend_gap_floor
                and mom_gap[t] >= c.micro_rescue_score1_mom_gap_floor
                and c.micro_rescue_score1_roc5_min <= roc5_now <= c.micro_rescue_score1_roc5_max
                and roc10_now >= c.micro_rescue_score1_roc10_min
                and roc20_now >= c.micro_rescue_score1_roc20_min
                and rsi_now >= c.micro_rescue_score1_rsi_min
                and mfi_now >= c.micro_rescue_score1_mfi_min
            )
            neartrend_branch = bool(
                s_now == 1 and (not trend_ok[t]) and (not mom_ok[t]) and vol_ok[t]
                and 1 <= days_since_tqqq_reentry <= c.micro_rescue_neartrend_max_age
                and trend_gap[t] >= c.micro_rescue_neartrend_gap_floor
                and c.micro_rescue_neartrend_roc5_min <= roc5_now <= c.micro_rescue_neartrend_roc5_max
                and roc20_now >= c.micro_rescue_neartrend_roc20_min
                and rsi_now >= c.micro_rescue_neartrend_rsi_min
                and mfi_now >= c.micro_rescue_neartrend_mfi_min
            )
            mature_branch = bool(
                s_now == 2 and trend_ok[t] and (not mom_ok[t]) and vol_ok[t]
                and c.micro_rescue_mature_min_age <= days_since_tqqq_reentry <= c.micro_rescue_mature_max_age
                and mom_gap[t] >= c.micro_rescue_mature_mom_gap_floor
                and trend_gap[t] >= c.micro_rescue_mature_trend_gap_floor
                and roc5_now >= c.micro_rescue_mature_roc5_min
                and rsi_now >= c.micro_rescue_mature_rsi_min
                and mfi_now >= c.micro_rescue_mature_mfi_min
            )
            if c.micro_rescue_mode == "score1_pullback":
                micro_branch = "score1_pullback" if pullback_branch else ""
            elif c.micro_rescue_mode == "score1_neartrend":
                micro_branch = "score1_neartrend" if neartrend_branch else ""
            elif c.micro_rescue_mode == "mature_mom":
                micro_branch = "mature_mom" if mature_branch else ""
            else:  # combo
                micro_branch = ("score1_pullback" if pullback_branch else
                                ("score1_neartrend" if neartrend_branch else
                                 ("mature_mom" if mature_branch else "")))
            micro_signal = bool(
                cur == 3 and nxt == 2 and micro_branch and (t - last_any_veto_t) > 1
                and (not bool(crash[t])) and days_since_crash >= c.micro_rescue_min_post_crash_days
                and f["QQQ_ATRPCT14"][t] <= c.micro_rescue_atr_max
                and vol[t] <= c.micro_rescue_vol_max and dd252[t] >= c.micro_rescue_dd_floor
                and (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and (not staged_fired) and (not snapback_fired) and (not fast_fired)
            )
            if micro_signal:
                micro_rescue_signal_days += 1
                nxt = 3
                down_count = max(0, c.down_confirm - 1)
                micro_rescue_blocks += 1
                micro_rescue_score1_pullback_blocks += int(micro_branch == "score1_pullback")
                micro_rescue_score1_neartrend_blocks += int(micro_branch == "score1_neartrend")
                micro_rescue_mature_mom_blocks += int(micro_branch == "mature_mom")
                last_any_veto_t = t

        # Optional immediate TQQQ-only step-down. Keeps QQQ/QLD state engine intact.
        if cur == 3 and nxt == 3 and c.tqqq_exit_mode != "off":
            texit = False
            if c.tqqq_exit_mode == "ma20_macd":
                texit = (f["TQQQ_CLOSE"][t] < f["TQQQ_SMA20"][t]) and (f["TQQQ_MACD_HIST"][t] < 0)
            elif c.tqqq_exit_mode == "qqq50_macd":
                texit = (f["QQQ_CLOSE"][t] < f["QQQ_SMA50"][t]) and (f["TQQQ_MACD_HIST"][t] < 0)
            elif c.tqqq_exit_mode == "risk":
                texit = f["V15_RISK_SCORE"][t] >= c.tqqq_exit_score
            if texit:
                nxt = 2; defense_events += 1; down_count = 0

        # V3 narrow trailing-drawdown guard. One-step only; avoids turning into a generic stop-loss strategy.
        if cur == 3 and nxt == 3 and c.dd_guard_mode != "off":
            tdd = f[f"TQQQ_DD{c.dd_guard_days}"][t]
            guard = bool(np.isfinite(tdd) and tdd <= c.dd_guard_threshold)
            if c.dd_guard_mode == "dd_macd":
                guard = guard and (f["TQQQ_MACD_HIST"][t] < 0)
            elif c.dd_guard_mode == "dd_ema":
                guard = guard and (f["TQQQ_CLOSE"][t] < f["TQQQ_EMA20"][t])
            if guard:
                nxt = 2; defense_events += 1; down_count = 0

        # Immediate composite defense cap, applied after the normal/fast decision.
        if c.defense_mode != "off" and f["V15_RISK_SCORE"][t] >= c.defense_score_threshold:
            cap = {"cap2": 2, "cap1": 1, "cash": 0}[c.defense_mode]
            if nxt > cap:
                nxt = cap
                defense_events += 1
                down_count = 0

        # If a later defense rule overruled the explosive 3x decision, disarm the latch.
        if nxt < 3 and explosive_fired:
            explosive_latch = 0
            dd_rebound_latch = 0

        # V12 ordinary TQQQ re-promotion guard. It is intentionally applied only
        # outside the bounded post-crash context and never changes a TQQQ exit.
        # All crash/recovery paths remain exempt, keeping the two experiments orthogonal.
        if c.repromote_guard_mode != "off" and cur < 3 and nxt == 3 and not bool(crash[t]):
            crash_exempt_window = max(42, c.trough_recovery_window_days, c.dd_rebound_window_days, c.explosive_window_days)
            if days_since_crash > crash_exempt_window and (not trough_fired) and (not explosive_fired):
                lb = max(1, c.repromote_lookback_days)
                recent_delevers = sum(1 for di in ordinary_delever_history if 0 < (t - di) <= lb)
                if recent_delevers >= max(1, c.repromote_trigger_count):
                    repromote_guard_regime_days += 1
                    base3 = bool(score[t] >= 3)
                    vol_util = (vol[t] / c.vol_threshold) if c.vol_threshold > 1e-12 else float("inf")
                    strong2 = bool(
                        score[t] >= 2
                        and f["QQQ_ROC10"][t] >= c.repromote_roc10_min
                        and f["QQQ_RSI14"][t] >= c.repromote_rsi_min
                        and vol_util <= c.repromote_vol_util_max
                    )
                    if c.repromote_guard_mode == "score3_only":
                        allow_repromote = base3
                    elif c.repromote_guard_mode == "cooldown_strong":
                        gap_ok = (t - last_ordinary_delever_t) >= max(0, c.repromote_cooldown_days)
                        allow_repromote = base3 or (gap_ok and strong2)
                    else:  # score3_or_strong
                        allow_repromote = base3 or strong2
                    if allow_repromote:
                        repromote_guard_passes += 1
                    else:
                        nxt = 2
                        repromote_guard_blocks += 1
                        down_count = 0

        # V20 stateless shadow overlay. Keep the underlying state-machine decision
        # (usually QLD) but optionally execute one extra TQQQ session. Because `nxt`
        # is untouched, the following day's control logic resumes from QLD and the
        # tactical hold cannot create a multi-day path cascade.
        if c.shadow_rescue_mode != "off":
            s_now = int(score[t])
            roc5_now = float(f["QQQ_ROC5"][t])
            roc10_now = float(f["QQQ_ROC10"][t])
            roc20_now = float(f["QQQ_ROC20"][t])
            rsi_now = float(f["QQQ_RSI14"][t])
            mfi_now = float(f["QQQ_MFI14"][t])
            shadow_score1 = bool(
                s_now == 1 and (not trend_ok[t]) and (not mom_ok[t]) and vol_ok[t]
                and 1 <= days_since_tqqq_reentry <= c.shadow_recent_max_age
                and trend_gap[t] >= c.shadow_score1_trend_gap_floor
                and mom_gap[t] >= c.shadow_score1_mom_gap_floor
                and c.shadow_score1_roc5_min <= roc5_now <= c.shadow_score1_roc5_max
                and roc10_now >= c.shadow_score1_roc10_min
                and roc20_now >= c.shadow_score1_roc20_min
                and rsi_now >= c.shadow_score1_rsi_min
                and mfi_now >= c.shadow_score1_mfi_min
            )
            shadow_mature = bool(
                s_now == 2 and trend_ok[t] and (not mom_ok[t]) and vol_ok[t]
                and c.shadow_mature_min_age <= days_since_tqqq_reentry <= c.shadow_mature_max_age
                and mom_gap[t] >= c.shadow_mature_mom_gap_floor
                and trend_gap[t] >= c.shadow_mature_trend_gap_floor
                and c.shadow_mature_roc5_min <= roc5_now <= c.shadow_mature_roc5_max
                and rsi_now >= c.shadow_mature_rsi_min
                and mfi_now >= c.shadow_mature_mfi_min
            )
            if c.shadow_rescue_mode == "score1_recent":
                shadow_branch = "score1_recent" if shadow_score1 else ""
            elif c.shadow_rescue_mode == "mature_mom":
                shadow_branch = "mature_mom" if shadow_mature else ""
            else:
                shadow_branch = "score1_recent" if shadow_score1 else ("mature_mom" if shadow_mature else "")
            shadow_signal = bool(
                cur == 3 and nxt == 2 and shadow_branch
                and (t - last_any_veto_t) > 1
                and (not bool(crash[t])) and days_since_crash >= c.shadow_min_post_crash_days
                and f["QQQ_ATRPCT14"][t] <= c.shadow_atr_max
                and vol[t] <= c.shadow_vol_max and dd252[t] >= c.shadow_dd_floor
                and (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and (not staged_fired) and (not snapback_fired) and (not fast_fired)
            )
            if shadow_signal:
                shadow_rescue_signal_days += 1
                shadow_exec[t + 1] = 3
                shadow_rescue_blocks += 1
                shadow_rescue_score1_blocks += int(shadow_branch == "score1_recent")
                shadow_rescue_mature_mom_blocks += int(shadow_branch == "mature_mom")

        # V21 bidirectional stateless preview overlay. V20 held one extra TQQQ session
        # on selected 3x->2x exits. V21 adds the forward mirror: execute TQQQ for one
        # session while the control state remains/steps only to QLD. `nxt` is NEVER
        # changed, so the following decision resumes from the original control path.
        if c.preview_overlay_mode != "off" and shadow_exec[t + 1] < 3:
            s_now = int(score[t])
            roc5_now = float(f["QQQ_ROC5"][t])
            roc10_now = float(f["QQQ_ROC10"][t])
            roc20_now = float(f["QQQ_ROC20"][t])
            rsi_now = float(f["QQQ_RSI14"][t])
            mfi_now = float(f["QQQ_MFI14"][t])
            qld3 = bool(
                cur == 2 and nxt == 2 and s_now == 3
                and days_since_delever <= c.preview_qld_max_delever_age
                and roc5_now >= c.preview_qld3_roc5_min
                and roc10_now >= c.preview_qld3_roc10_min
                and roc20_now >= c.preview_qld3_roc20_min
                and rsi_now >= c.preview_qld3_rsi_min and mfi_now >= c.preview_qld3_mfi_min
            )
            qld2_structure = bool(
                s_now == 2 and vol_ok[t]
                and ((trend_ok[t] and (not mom_ok[t]) and mom_gap[t] >= c.preview_qld2_mom_gap_floor)
                     or ((not trend_ok[t]) and mom_ok[t] and trend_gap[t] >= c.preview_qld2_trend_gap_floor))
            )
            qld2 = bool(
                cur == 2 and nxt == 2 and qld2_structure
                and days_since_delever <= c.preview_qld_max_delever_age
                and roc5_now >= c.preview_qld2_roc5_min
                and roc10_now >= c.preview_qld2_roc10_min
                and roc20_now >= c.preview_qld2_roc20_min
                and rsi_now >= c.preview_qld2_rsi_min and mfi_now >= c.preview_qld2_mfi_min
            )
            stepup = bool(
                cur < 2 and nxt == 2 and s_now >= c.preview_stepup_score_min and bool(entry_mask[t])
                and roc5_now >= c.preview_stepup_roc5_min
                and roc10_now >= c.preview_stepup_roc10_min
                and roc20_now >= c.preview_stepup_roc20_min
                and rsi_now >= c.preview_stepup_rsi_min and mfi_now >= c.preview_stepup_mfi_min
            )
            if c.preview_overlay_mode == "qld_score3":
                preview_branch = "qld_score3" if qld3 else ""
            elif c.preview_overlay_mode == "qld_score2":
                preview_branch = "qld_score2" if qld2 else ""
            elif c.preview_overlay_mode == "stepup":
                preview_branch = "stepup" if stepup else ""
            else:
                preview_branch = "qld_score3" if qld3 else ("qld_score2" if qld2 else ("stepup" if stepup else ""))
            preview_signal = bool(
                preview_branch and (t - last_any_veto_t) > 1
                and (not bool(crash[t])) and days_since_crash >= c.preview_min_post_crash_days
                and f["QQQ_ATRPCT14"][t] <= c.preview_atr_max
                and vol[t] <= c.preview_vol_max and dd252[t] >= c.preview_dd_floor
                and (not explosive_fired) and (not recovery_fired) and (not rebound_fired)
                and (not staged_fired) and (not snapback_fired)
            )
            if preview_signal:
                preview_overlay_signal_days += 1
                shadow_exec[t + 1] = 3
                preview_overlay_blocks += 1
                preview_overlay_qld3_blocks += int(preview_branch == "qld_score3")
                preview_overlay_qld2_blocks += int(preview_branch == "qld_score2")
                preview_overlay_stepup_blocks += int(preview_branch == "stepup")

        # Track time since the most recent 3x -> lower-leverage transition. V9
        # also records whether that deleveraging was an ordinary non-crash correction.
        if cur == 3 and nxt < 3:
            days_since_delever = 0
            days_since_correction = 0 if not bool(crash[t]) else 10**9
            if not bool(crash[t]):
                ordinary_delever_history.append(t)
                last_ordinary_delever_t = t
                # Keep only the largest V12 lookback neighborhood; history is strategy-local.
                cutoff = t - max(120, c.repromote_lookback_days)
                if len(ordinary_delever_history) > 32:
                    ordinary_delever_history = [di for di in ordinary_delever_history if di >= cutoff]
        elif nxt == 3:
            days_since_delever = 10**9
            days_since_correction = 10**9
        else:
            if days_since_delever < 10**9:
                days_since_delever += 1
            if days_since_correction < 10**9:
                days_since_correction += 1

        # Track how many staged QLD sessions have been accumulated before promotion.
        if staged_fired and nxt == 2:
            if cur == 2 and staged_qld_age >= 1:
                staged_qld_age += 1
            else:
                staged_qld_age = 1
        else:
            staged_qld_age = -1

        if nxt != cur:
            transitions += 1
            if cur != 3 and nxt == 3:
                days_since_tqqq_reentry = 0
        state = int(nxt)
        states[t + 1] = state

    # Apply actual held exposure. V20/V21 overlays affect only execution, not
    # the underlying control-state machine used on the next decision.
    exec_states = states.copy()
    shadow_mask = shadow_exec >= 0
    exec_states[shadow_mask] = shadow_exec[shadow_mask]
    strat_ret = np.zeros(n, dtype=float)
    qret, lret, tret = f["QQQ_O2O_RET"], f["QLD_O2O_RET"], f["TQQQ_O2O_RET"]
    strat_ret[exec_states == 1] = qret[exec_states == 1]
    strat_ret[exec_states == 2] = lret[exec_states == 2]
    strat_ret[exec_states == 3] = tret[exec_states == 3]

    # Turnover cost follows actual executed exposure.
    prev = np.r_[0, exec_states[:-1]]
    changed = exec_states != prev
    sides = np.zeros(n, dtype=float)
    sides[changed] = 1.0
    both_assets = changed & (exec_states != 0) & (prev != 0)
    sides[both_assets] = 2.0
    strat_ret -= sides * cost_per_side
    strat_ret = np.clip(strat_ret, -0.999999, None)

    delever_idx = np.flatnonzero((exec_states[:-1] == 3) & (exec_states[1:] < 3))
    whipsaw_10d = 0
    whipsaw_20d = 0
    for di in delever_idx:
        if np.any(exec_states[di+1:min(len(exec_states), di+12)] == 3):
            whipsaw_10d += 1
        if np.any(exec_states[di+1:min(len(exec_states), di+22)] == 3):
            whipsaw_20d += 1

    stats = {
        "transitions": transitions,
        "crash_exits": crash_exits,
        "fast_entries": fast_entries,
        "defense_events": defense_events,
        "staged_entries": staged_entries,
        "staged_promotions": staged_promotions,
        "explosive_entries": explosive_entries,
        "correction_entries": correction_entries,
        "correction_tqqq_entries": correction_tqqq_entries,
        "dd_rebound_entries": dd_rebound_entries,
        "trough_recovery_entries": trough_recovery_entries,
        "trough_recovery_qld_entries": trough_recovery_qld_entries,
        "trough_recovery_promotions": trough_recovery_promotions,
        "trough_recovery_hold_days": trough_recovery_hold_days,
        "trough_new_lows": tr2_new_lows,
        "trough_split_qld_days": trough_split_qld_days,
        "trough_split_promotions": trough_split_promotions,
        "trough_split_aborts": trough_split_aborts,
        "repromote_guard_blocks": repromote_guard_blocks,
        "repromote_guard_passes": repromote_guard_passes,
        "repromote_guard_regime_days": repromote_guard_regime_days,
        "soft_vol_signal_days": soft_vol_signal_days,
        "soft_vol_entries": soft_vol_entries,
        "impulse_veto_signal_days": impulse_veto_signal_days,
        "impulse_veto_blocks": impulse_veto_blocks,
        "impulse_veto_score1_blocks": impulse_veto_score1_blocks,
        "impulse_veto_score2_blocks": impulse_veto_score2_blocks,
        "precision_veto_signal_days": precision_veto_signal_days,
        "precision_veto_blocks": precision_veto_blocks,
        "precision_veto_trend_blocks": precision_veto_trend_blocks,
        "precision_veto_mom_blocks": precision_veto_mom_blocks,
        "precision_veto_score1_blocks": precision_veto_score1_blocks,
        "precision_veto_fresh_blocks": precision_veto_fresh_blocks,
        "reentry_grace_signal_days": reentry_grace_signal_days,
        "reentry_grace_blocks": reentry_grace_blocks,
        "reentry_grace_mom_blocks": reentry_grace_mom_blocks,
        "reentry_grace_trend_blocks": reentry_grace_trend_blocks,
        "reentry_grace_score2_blocks": reentry_grace_score2_blocks,
        "micro_rescue_signal_days": micro_rescue_signal_days,
        "micro_rescue_blocks": micro_rescue_blocks,
        "micro_rescue_score1_pullback_blocks": micro_rescue_score1_pullback_blocks,
        "micro_rescue_score1_neartrend_blocks": micro_rescue_score1_neartrend_blocks,
        "micro_rescue_mature_mom_blocks": micro_rescue_mature_mom_blocks,
        "shadow_rescue_signal_days": shadow_rescue_signal_days,
        "shadow_rescue_blocks": shadow_rescue_blocks,
        "shadow_rescue_score1_blocks": shadow_rescue_score1_blocks,
        "shadow_rescue_mature_mom_blocks": shadow_rescue_mature_mom_blocks,
        "preview_overlay_signal_days": preview_overlay_signal_days,
        "preview_overlay_blocks": preview_overlay_blocks,
        "preview_overlay_qld3_blocks": preview_overlay_qld3_blocks,
        "preview_overlay_qld2_blocks": preview_overlay_qld2_blocks,
        "preview_overlay_stepup_blocks": preview_overlay_stepup_blocks,
        "adaptive_hold_days": adaptive_hold_days,
        "adaptive_exit_events": adaptive_exit_events,
        "tqqq_delever_events": int(len(delever_idx)),
        "whipsaw_10d": int(whipsaw_10d),
        "whipsaw_20d": int(whipsaw_20d),
        "time_cash": float(np.mean(exec_states == 0)),
        "time_qqq": float(np.mean(exec_states == 1)),
        "time_qld": float(np.mean(exec_states == 2)),
        "time_tqqq": float(np.mean(exec_states == 3)),
    }
    return strat_ret, exec_states, stats


# -----------------------------
# Metrics / search-stage evaluation
# -----------------------------

def calc_mdd(equity: np.ndarray) -> float:
    if len(equity) == 0:
        return float("nan")
    peak = np.maximum.accumulate(equity)
    dd = equity / peak - 1.0
    return float(np.min(dd))


def full_metrics(returns: np.ndarray) -> Dict[str, float]:
    if len(returns) == 0:
        return {"total_return": np.nan, "cagr": np.nan, "mdd": np.nan, "vol": np.nan, "sharpe": np.nan, "calmar": np.nan}
    logc = np.r_[0.0, np.cumsum(np.log1p(np.clip(returns, -0.999999, None)))]
    eq = np.exp(logc)
    years = max(len(returns) / TRADING_DAYS, 1 / TRADING_DAYS)
    cagr = eq[-1] ** (1 / years) - 1.0
    mdd = calc_mdd(eq)
    vol = float(np.std(returns, ddof=0) * math.sqrt(TRADING_DAYS))
    ann_mean = float(np.mean(returns) * TRADING_DAYS)
    sharpe = ann_mean / vol if vol > 1e-12 else np.nan
    calmar = cagr / abs(mdd) if mdd < -1e-12 else np.nan
    return {"total_return": float(eq[-1] - 1), "cagr": float(cagr), "mdd": float(mdd), "vol": vol, "sharpe": float(sharpe), "calmar": float(calmar)}


def rolling_return_stats(logcum: np.ndarray, bench_logcum: np.ndarray) -> Dict[str, float]:
    out: Dict[str, float] = {}
    weighted = []
    short_wins = []
    tail_p10 = []
    short_tail_p10 = []
    short_min_excess = float("inf")
    min_win = 1.0
    min_excess = float("inf")
    available = 0
    for label, h in HORIZONS:
        if len(logcum) <= h or len(bench_logcum) <= h:
            out[f"ret_win_{label}"] = np.nan
            out[f"worst_excess_{label}"] = np.nan
            continue
        sret = logcum[h:] - logcum[:-h]
        b = bench_logcum[h:] - bench_logcum[:-h]
        win_mask = sret >= b - 1e-12
        win = float(np.mean(win_mask))
        excess = sret - b
        wins = int(np.sum(win_mask))
        target_wins_79 = int(math.ceil(0.79 * len(win_mask) - 1e-12))
        out[f"ret_windows_{label}"] = int(len(win_mask))
        out[f"ret_wins_{label}"] = wins
        out[f"ret_wins_needed_79_{label}"] = max(0, target_wins_79 - wins)
        if label == "1Y":
            fail = ~win_mask
            out["ret_1y_near_miss_25bp"] = int(np.sum(fail & (excess >= -0.0025)))
            out["ret_1y_near_miss_50bp"] = int(np.sum(fail & (excess >= -0.0050)))
            out["ret_1y_near_miss_100bp"] = int(np.sum(fail & (excess >= -0.0100)))
        worst = float(np.min(excess))
        p10 = float(np.quantile(excess, 0.10))
        out[f"ret_win_{label}"] = win
        out[f"worst_excess_{label}"] = worst
        out[f"p10_excess_{label}"] = p10
        min_win = min(min_win, win)
        min_excess = min(min_excess, worst)
        weighted.append(win)
        tail_p10.append(p10)
        if label in {"6M", "1Y", "2Y", "3Y"}:
            short_wins.append(win)
            short_tail_p10.append(p10)
            short_min_excess = min(short_min_excess, worst)
        available += 1
    out["min_return_win_rate"] = min_win if available else np.nan
    out["avg_return_win_rate"] = float(np.mean(weighted)) if weighted else np.nan
    out["short_return_win_rate"] = float(np.mean(short_wins)) if short_wins else np.nan
    out["avg_p10_log_excess"] = float(np.mean(tail_p10)) if tail_p10 else np.nan
    out["short_p10_log_excess"] = float(np.mean(short_tail_p10)) if short_tail_p10 else np.nan
    out["short_worst_log_excess"] = short_min_excess if short_tail_p10 else np.nan
    out["worst_log_excess"] = min_excess if available else np.nan
    return out


def config_to_short_text(c: Config) -> str:
    gates = []
    if c.require_qqq_ma50: gates.append("Q>MA50")
    if c.require_qqq_ma200: gates.append("Q>MA200")
    if c.require_tqqq_ma20: gates.append("T>MA20")
    if c.require_tqqq_ma50: gates.append("T>MA50")
    if c.macd_gate != "none": gates.append(f"MACD:{c.macd_gate}")
    if c.rsi_asset != "none": gates.append(f"RSI:{c.rsi_asset}>={c.rsi_min}")
    if c.stoch_asset != "none": gates.append(f"STOCH:{c.stoch_asset}")
    if c.cci_asset != "none": gates.append(f"CCI:{c.cci_asset}")
    if c.boll_asset != "none": gates.append(f"BOLL:{c.boll_asset}/{c.boll_rule}")
    if c.adx_asset != "none": gates.append(f"ADX:{c.adx_asset}>={c.adx_min}")
    if c.mfi_asset != "none": gates.append(f"MFI:{c.mfi_asset}>={c.mfi_min}")
    if c.willr_asset != "none": gates.append(f"W%R:{c.willr_asset}>={c.willr_min}")
    if c.obv_asset != "none": gates.append(f"OBV:{c.obv_asset}")
    if c.donchian_asset != "none": gates.append(f"DON:{c.donchian_asset}>={c.donchian_min_pos}")
    if c.atr_asset != "none": gates.append(f"ATR:{c.atr_asset}<={c.atr_max_pct}")
    if c.roc_asset != "none": gates.append(f"ROC{c.roc_days}:{c.roc_asset}>={c.roc_min}")
    if c.defense_mode != "off": gates.append(f"DEF:{c.defense_mode}@{c.defense_score_threshold}")
    if c.tqqq_exit_mode != "off": gates.append(f"TEXIT:{c.tqqq_exit_mode}@{c.tqqq_exit_score}")
    if c.recovery_mode != "off": gates.append(f"REC:{c.recovery_mode}/{c.recovery_days}d/floor{c.recovery_mom_floor:.2f}")
    if c.bull_override_mode != "off": gates.append(f"BULL:{c.bull_override_mode}/roc{c.bull_roc20_min:.2f}")
    if c.dd_guard_mode != "off": gates.append(f"DDG:{c.dd_guard_mode}/{c.dd_guard_days}d/{c.dd_guard_threshold:.2f}")
    if c.rebound_bypass_mode != "off": gates.append(f"REB:{c.rebound_bypass_mode}/{c.rebound_window_days}d/r10{c.rebound_roc10_min:.2f}/r20{c.rebound_roc20_min:.2f}")
    if c.bridge_mode != "off": gates.append(f"BR:{c.bridge_mode}/r60{c.bridge_roc60_min:.2f}/dd{c.bridge_dd_floor:.2f}")
    if c.snapback_mode != "off": gates.append(f"SNAP:{c.snapback_mode}/{c.snapback_days}d/ema{c.snapback_ema_days}/dd{c.snapback_dd_floor:.2f}")
    if c.crash_release_mode != "off": gates.append(f"CRREL:{c.crash_release_mode}/{c.crash_release_wait}d/r5{c.crash_release_roc5_min:.2f}/r10{c.crash_release_roc10_min:.2f}")
    if c.staged_reentry_mode != "off": gates.append(f"STAGE:{c.staged_reentry_mode}/{c.staged_window_days}d/w{c.staged_promote_wait}/atr{c.staged_atr_max:.2f}")
    if c.explosive_rebound_mode != "off": gates.append(f"EXP:{c.explosive_rebound_mode}/{c.explosive_window_days}d/r5{c.explosive_roc5_min:.2f}/r10{c.explosive_roc10_min:.2f}/h{c.explosive_hold_days}")
    if c.correction_reentry_mode != "off": gates.append(f"CORR:{c.correction_reentry_mode}/{c.correction_window_days}d/r10{c.correction_roc10_min:.3f}/dd{c.correction_dd_floor:.2f}")
    if c.dd_rebound_mode != "off": gates.append(f"DDR:{c.dd_rebound_mode}/{c.dd_rebound_window_days}d/reb{c.dd_rebound_min:.2f}/r5{c.dd_rebound_roc5_min:.2f}/h{c.dd_rebound_hold_days}")
    if c.trough_recovery_mode != "off": gates.append(f"TR2:{c.trough_recovery_mode}/{c.trough_recovery_window_days}d/age{c.trough_age_min}/reb{c.trough_rebound_min:.3f}/r5{c.trough_roc5_min:.2f}/r10{c.trough_roc10_min:.2f}")
    if c.trough_recovery_mode == "qld_split3": gates.append(f"TR2PROM:w{c.trough_confirm_days}/mx{c.trough_bridge_max_days}/reb{c.trough_promote_rebound_min:.2f}/r10{c.trough_promote_roc10_min:.2f}/rsi{c.trough_promote_rsi_min}")
    if c.repromote_guard_mode != "off": gates.append(f"RPG:{c.repromote_guard_mode}/L{c.repromote_lookback_days}/N{c.repromote_trigger_count}/r10{c.repromote_roc10_min:.2f}/rsi{c.repromote_rsi_min}/vu{c.repromote_vol_util_max:.2f}")
    if c.soft_vol_mode != "off": gates.append(f"SV:{c.soft_vol_mode}/v{c.vol_threshold:.4f}-{c.soft_vol_max:.4f}/c{c.soft_vol_confirm_days}/r10{c.soft_vol_roc10_min:.2f}/r20{c.soft_vol_roc20_min:.2f}/rsi{c.soft_vol_rsi_min}")
    if c.impulse_veto_mode != "off": gates.append(f"IV:{c.impulse_veto_mode}/h{c.impulse_veto_max_days}/ema{c.impulse_veto_ema_days}/r5{c.impulse_veto_roc5_min:.2f}/r10{c.impulse_veto_roc10_min:.2f}/rsi{c.impulse_veto_rsi_min}/mfi{c.impulse_veto_mfi_min}")
    if c.precision_veto_mode != "off": gates.append(f"PV:{c.precision_veto_mode}/h{c.precision_veto_max_days}/ema{c.precision_veto_ema_days}/r5{c.precision_veto_roc5_min:.2f}/r10{c.precision_veto_roc10_min:.2f}/tg{c.precision_veto_trend_gap_floor:.3f}/mg{c.precision_veto_mom_gap_floor:.3f}/re{c.precision_veto_recent_reentry_days}")
    if c.reentry_grace_mode == "split_reentry":
        gates.append(f"RG:SPLIT/age{c.reentry_grace_min_age}-{c.reentry_grace_max_age}/pc{c.reentry_grace_min_post_crash_days}/md{c.reentry_grace_max_days}/mgate{c.reentry_grace_mom_entry_gate}/tgate{c.reentry_grace_trend_entry_gate}/mr5{c.reentry_grace_mom_roc5_min:.3f}/tr5{c.reentry_grace_trend_roc5_min:.3f}/tg{c.reentry_grace_trend_gap_floor:.3f}/mg{c.reentry_grace_mom_gap_floor:.3f}")
    elif c.reentry_grace_mode != "off":
        gates.append(f"RG:{c.reentry_grace_mode}/age{c.reentry_grace_min_age}-{c.reentry_grace_max_age}/pc{c.reentry_grace_min_post_crash_days}/eg{c.reentry_grace_entry_gate}/ema{c.reentry_grace_ema_days}/tg{c.reentry_grace_trend_gap_floor:.3f}/mg{c.reentry_grace_mom_gap_floor:.3f}")
    if c.micro_rescue_mode != "off":
        gates.append(f"MR:{c.micro_rescue_mode}/pc{c.micro_rescue_min_post_crash_days}/ra{c.micro_rescue_recent_max_age}/na{c.micro_rescue_neartrend_max_age}/ma{c.micro_rescue_mature_min_age}-{c.micro_rescue_mature_max_age}")
    if c.shadow_rescue_mode != "off":
        gates.append(f"SH:{c.shadow_rescue_mode}/pc{c.shadow_min_post_crash_days}/ra{c.shadow_recent_max_age}/ma{c.shadow_mature_min_age}-{c.shadow_mature_max_age}")
    if c.preview_overlay_mode != "off":
        gates.append(f"PV21:{c.preview_overlay_mode}/pc{c.preview_min_post_crash_days}/qa{c.preview_qld_max_delever_age}/q2r10{c.preview_qld2_roc10_min:.2f}/su{c.preview_stepup_score_min}")
    if c.adaptive_down_mode != "off": gates.append(f"HYST:score2/{c.soft_down_confirm}d/hard{c.hard_down_confirm}d")
    if c.gate_scope != "all_up": gates.append(f"SCOPE:{c.gate_scope}")
    if c.fast_mode == "conditional": gates.append("FAST:conditional")
    return ", ".join(gates) if gates else "none"


def evaluate_one_search(c: Config) -> Dict[str, Any]:
    f = _WF
    ret, states, st = simulate_strategy(c, f, _W_COST)
    rr = ret[_W_EVAL_START:-1]
    met = full_metrics(rr)
    logcum = np.r_[0.0, np.cumsum(np.log1p(rr))]
    roll = rolling_return_stats(logcum, _W_BENCH_LOG_CUM)

    # Search score strongly rewards cross-horizon return consistency and MDD improvement.
    # Exact rolling MDD is intentionally deferred to finalists for speed.
    mdd_improvement = met["mdd"] - _W_BENCH_FULL["mdd"]  # positive = better
    cagr_ratio = (1 + met["cagr"]) / max(1e-9, 1 + _W_BENCH_FULL["cagr"])
    complexity_penalty = 0.008 * max(0, c.filter_count - 3)
    dead_v13_modules = int(c.trough_recovery_mode != "off" and st.get("trough_recovery_entries", 0) == 0)
    dead_soft_vol = int(c.soft_vol_mode != "off" and st.get("soft_vol_entries", 0) == 0)
    dead_v14_modules = dead_v13_modules + dead_soft_vol
    dead_impulse_veto = int(c.impulse_veto_mode != "off" and st.get("impulse_veto_blocks", 0) == 0)
    dead_v15_modules = dead_v14_modules + dead_impulse_veto
    dead_precision_veto = int(c.precision_veto_mode != "off" and st.get("precision_veto_blocks", 0) == 0)
    dead_v16_modules = dead_v15_modules + dead_precision_veto
    dead_reentry_grace = int(c.reentry_grace_mode != "off" and st.get("reentry_grace_blocks", 0) == 0)
    dead_v17_modules = dead_v16_modules + dead_reentry_grace
    dead_micro_rescue = int(c.micro_rescue_mode != "off" and st.get("micro_rescue_blocks", 0) == 0)
    dead_v19_modules = dead_v17_modules + dead_micro_rescue
    dead_shadow_rescue = int(c.shadow_rescue_mode != "off" and st.get("shadow_rescue_blocks", 0) == 0)
    dead_v20_modules = dead_v19_modules + dead_shadow_rescue
    dead_preview_overlay = int(c.preview_overlay_mode != "off" and st.get("preview_overlay_blocks", 0) == 0)
    dead_v21_modules = dead_v20_modules + dead_preview_overlay
    dead_module_penalty = 0.025 * dead_v21_modules
    precision_block_penalty = 0.004 * max(0, st.get("precision_veto_blocks", 0) - 4)
    grace_free_blocks = 6 if c.reentry_grace_mode == "split_reentry" else 4
    reentry_grace_block_penalty = 0.004 * max(0, st.get("reentry_grace_blocks", 0) - grace_free_blocks)
    micro_rescue_block_penalty = 0.006 * max(0, st.get("micro_rescue_blocks", 0) - 4)
    shadow_rescue_block_penalty = 0.006 * max(0, st.get("shadow_rescue_blocks", 0) - 4)
    preview_overlay_block_penalty = 0.006 * max(0, st.get("preview_overlay_blocks", 0) - 6)
    turnover_penalty = 0.0005 * (st["transitions"] / max(len(rr), 1) * TRADING_DAYS)
    guard_block_penalty = 0.0008 * (st.get("repromote_guard_blocks", 0) / max(len(rr), 1) * TRADING_DAYS)
    search_score = (
        7.4 * roll.get("min_return_win_rate", 0.0)
        + 1.6 * roll.get("avg_return_win_rate", 0.0)
        + 3.8 * roll.get("short_return_win_rate", 0.0)
        + 3.0 * roll.get("ret_win_1Y", 0.0)
        + 5.0 * roll.get("ret_win_2Y", 0.0)
        - 0.003 * roll.get("ret_wins_needed_79_1Y", 0)
        + 0.0005 * min(40, roll.get("ret_1y_near_miss_25bp", 0))
        + 1.0 * max(-0.50, min(0.20, roll.get("avg_p10_log_excess", -0.50)))
        + 2.6 * max(-0.80, min(0.20, roll.get("short_p10_log_excess", -0.80)))
        + 1.8 * max(-1.60, min(0.20, roll.get("short_worst_log_excess", -1.60)))
        + 0.8 * max(-1.60, min(0.20, roll.get("worst_log_excess", -1.60)))
        + 2.0 * max(-0.5, min(0.5, mdd_improvement))
        + 0.8 * max(-0.5, min(0.5, cagr_ratio - 1.0))
        + 0.15 * max(-1.0, min(1.0, met.get("calmar", 0.0) if np.isfinite(met.get("calmar", np.nan)) else 0.0))
        - complexity_penalty
        - dead_module_penalty
        - precision_block_penalty
        - reentry_grace_block_penalty
        - micro_rescue_block_penalty
        - shadow_rescue_block_penalty
        - preview_overlay_block_penalty
        - turnover_penalty
        - guard_block_penalty
    )

    row: Dict[str, Any] = {
        "name": c.name,
        "family": c.family,
        "search_score": float(search_score),
        "config": asdict(c),
        "gates": config_to_short_text(c),
        **met,
        **roll,
        **st,
        "bench_cagr": _W_BENCH_FULL["cagr"],
        "bench_mdd": _W_BENCH_FULL["mdd"],
        "mdd_improvement": mdd_improvement,
        "cagr_excess": met["cagr"] - _W_BENCH_FULL["cagr"],
        "dead_v21_modules": dead_v21_modules,
        "dead_v20_modules": dead_v20_modules,
        "dead_v19_modules": dead_v19_modules,
        "dead_v17_modules": dead_v17_modules,
        "dead_v16_modules": dead_v16_modules,
        "dead_v15_modules": dead_v15_modules,
        "dead_v14_modules": dead_v14_modules,
        "dead_v13_modules": dead_v13_modules,
        "dead_v12_modules": dead_v13_modules,
    }
    return row


def evaluate_chunk(configs: List[Config]) -> List[Dict[str, Any]]:
    rows = []
    for c in configs:
        try:
            rows.append(evaluate_one_search(c))
        except Exception as exc:
            rows.append({"name": c.name, "family": c.family, "error": repr(exc), "config": asdict(c), "search_score": -1e9})
    return rows


# -----------------------------
# Exact rolling MDD via segment tree
# -----------------------------

class DrawdownSegmentTree:
    """Range maximum drawdown queries on log-equity in O(log N).

    A segment stores:
      max_x: max log-equity
      min_x: min log-equity
      drop: max ordered log drawdown x_i - x_j for i <= j
    """

    def __init__(self, log_equity: np.ndarray):
        x = np.asarray(log_equity, dtype=float)
        n = 1
        while n < len(x): n *= 2
        self.n = n
        self.mx = np.full(2 * n, -np.inf, dtype=float)
        self.mn = np.full(2 * n, np.inf, dtype=float)
        self.dp = np.zeros(2 * n, dtype=float)
        self.mx[n:n + len(x)] = x
        self.mn[n:n + len(x)] = x
        for i in range(n - 1, 0, -1):
            self._pull(i)

    def _pull(self, i: int) -> None:
        l, r = i * 2, i * 2 + 1
        self.mx[i] = max(self.mx[l], self.mx[r])
        self.mn[i] = min(self.mn[l], self.mn[r])
        cross = self.mx[l] - self.mn[r] if np.isfinite(self.mx[l]) and np.isfinite(self.mn[r]) else 0.0
        self.dp[i] = max(self.dp[l], self.dp[r], cross, 0.0)

    @staticmethod
    def _combine(a: Tuple[float, float, float], b: Tuple[float, float, float]) -> Tuple[float, float, float]:
        amx, amn, adp = a; bmx, bmn, bdp = b
        if not np.isfinite(amx): return b
        if not np.isfinite(bmx): return a
        return (max(amx, bmx), min(amn, bmn), max(adp, bdp, amx - bmn, 0.0))

    def query_drop(self, left: int, right: int) -> float:
        """Inclusive [left, right]. Returns max log drop."""
        l = left + self.n
        r = right + self.n + 1
        left_acc = (-np.inf, np.inf, 0.0)
        right_acc = (-np.inf, np.inf, 0.0)
        while l < r:
            if l & 1:
                left_acc = self._combine(left_acc, (self.mx[l], self.mn[l], self.dp[l])); l += 1
            if r & 1:
                r -= 1; right_acc = self._combine((self.mx[r], self.mn[r], self.dp[r]), right_acc)
            l //= 2; r //= 2
        return self._combine(left_acc, right_acc)[2]


def exact_rolling_analysis(
    c: Config,
    features: Dict[str, Any],
    eval_start: int,
    cost_bps: float,
    bench_cache: Dict[int, Dict[str, np.ndarray]],
    include_details: bool,
) -> Tuple[Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]]]:
    ret, states, st = simulate_strategy(c, features, cost_bps / 10000.0)
    rr = ret[eval_start:-1]
    dates = pd.to_datetime(features["dates"])[eval_start:eval_start + len(rr) + 1]
    log_eq = np.r_[0.0, np.cumsum(np.log1p(rr))]
    tree = DrawdownSegmentTree(log_eq)
    met = full_metrics(rr)

    summaries: List[Dict[str, Any]] = []
    details: List[Dict[str, Any]] = []
    min_ret_win = 1.0
    min_mdd_win = 1.0
    min_both_win = 1.0
    min_worst_ret_excess = float("inf")
    min_worst_mdd_improvement = float("inf")
    perfect_horizons = 0
    available_horizons = 0
    both_values: List[float] = []
    short_both_values: List[float] = []
    p10_return_excess_values: List[float] = []
    startyear_short_floors: List[float] = []
    startyear_short_all_rates: List[float] = []
    worst_horizon_label = None
    worst_horizon_windows = 0
    worst_horizon_both_wins = 0
    worst_horizon_wins_needed_79 = 0
    worst_horizon_wins_needed_795 = 0
    worst_horizon_wins_needed_80 = 0
    one_year_windows = 0
    one_year_both_wins = 0
    one_year_wins_needed_79 = 0
    one_year_wins_needed_795 = 0
    one_year_wins_needed_80 = 0
    one_year_near_miss_25bp = 0
    one_year_near_miss_50bp = 0
    one_year_near_miss_100bp = 0
    two_year_windows = 0
    two_year_both_wins = 0
    two_year_wins_needed_80 = 0
    two_year_wins_needed_81 = 0
    two_year_wins_needed_82 = 0
    two_year_near_miss_25bp = 0
    two_year_near_miss_50bp = 0
    two_year_near_miss_100bp = 0

    for label, h in HORIZONS:
        if len(rr) < h:
            continue
        nwin = len(rr) - h + 1
        strat_logret = log_eq[h:] - log_eq[:-h]
        b = bench_cache[h]
        bench_logret = b["logret"][:nwin]
        bench_mdd = b["mdd"][:nwin]

        # Reporting metric: 6M/1Y use actual period return; 2Y+ use CAGR.
        # This avoids grotesque annualization of short rebound windows while preserving monotonic comparison.
        if h <= TRADING_DAYS:
            strat_ann = np.expm1(strat_logret)
            bench_ann = np.expm1(bench_logret)
            return_metric = "Period Return"
        else:
            ann_factor = TRADING_DAYS / h
            strat_ann = np.expm1(strat_logret * ann_factor)
            bench_ann = np.expm1(bench_logret * ann_factor)
            return_metric = "CAGR"

        strat_mdd = np.empty(nwin, dtype=float)
        for s in range(nwin):
            drop = tree.query_drop(s, s + h)
            strat_mdd[s] = math.exp(-drop) - 1.0

        ret_win_mask = strat_logret >= bench_logret - 1e-12
        mdd_win_mask = strat_mdd >= bench_mdd - 1e-12
        both = ret_win_mask & mdd_win_mask
        ret_win = float(np.mean(ret_win_mask))
        mdd_win = float(np.mean(mdd_win_mask))
        both_win = float(np.mean(both))
        both_wins = int(np.sum(both))
        wins_needed_79 = max(0, int(math.ceil(0.79 * nwin - 1e-12)) - both_wins)
        wins_needed_795 = max(0, int(math.ceil(0.795 * nwin - 1e-12)) - both_wins)
        wins_needed_80 = max(0, int(math.ceil(0.80 * nwin - 1e-12)) - both_wins)
        wins_needed_81 = max(0, int(math.ceil(0.81 * nwin - 1e-12)) - both_wins)
        wins_needed_82 = max(0, int(math.ceil(0.82 * nwin - 1e-12)) - both_wins)
        ret_excess_arr = strat_ann - bench_ann
        # V13 target diagnostics: count how many exact every-start windows remain to 79%.
        if worst_horizon_label is None or both_win < min_both_win - 1e-15:
            worst_horizon_label = label
            worst_horizon_windows = nwin
            worst_horizon_both_wins = both_wins
            worst_horizon_wins_needed_79 = wins_needed_79
            worst_horizon_wins_needed_795 = wins_needed_795
            worst_horizon_wins_needed_80 = wins_needed_80
        if label == "1Y":
            one_year_windows = nwin
            one_year_both_wins = both_wins
            one_year_wins_needed_79 = wins_needed_79
            one_year_wins_needed_795 = wins_needed_795
            one_year_wins_needed_80 = wins_needed_80
            eligible_fail = (~ret_win_mask) & mdd_win_mask
            one_year_near_miss_25bp = int(np.sum(eligible_fail & (ret_excess_arr >= -0.0025)))
            one_year_near_miss_50bp = int(np.sum(eligible_fail & (ret_excess_arr >= -0.0050)))
            one_year_near_miss_100bp = int(np.sum(eligible_fail & (ret_excess_arr >= -0.0100)))
        if label == "2Y":
            two_year_windows = nwin
            two_year_both_wins = both_wins
            two_year_wins_needed_80 = wins_needed_80
            two_year_wins_needed_81 = wins_needed_81
            two_year_wins_needed_82 = wins_needed_82
            eligible_fail = (~ret_win_mask) & mdd_win_mask
            two_year_near_miss_25bp = int(np.sum(eligible_fail & (ret_excess_arr >= -0.0025)))
            two_year_near_miss_50bp = int(np.sum(eligible_fail & (ret_excess_arr >= -0.0050)))
            two_year_near_miss_100bp = int(np.sum(eligible_fail & (ret_excess_arr >= -0.0100)))
        worst_ret_excess = float(np.min(ret_excess_arr))
        p10_ret_excess = float(np.quantile(ret_excess_arr, 0.10))
        worst_mdd_improve = float(np.min(strat_mdd - bench_mdd))

        # Cohort robustness: for short/medium horizons, score the weakest full-ish start-year cohort.
        # Require >=200 daily-start windows so partial edge years do not dominate the ranking.
        if label in {"6M", "1Y", "2Y", "3Y"}:
            start_years = np.asarray([d.year for d in dates[:nwin]])
            cohort_rates = []
            for yy in np.unique(start_years):
                mask_y = start_years == yy
                if int(np.sum(mask_y)) >= 200:
                    cohort_rates.append(float(np.mean(both[mask_y])))
            if cohort_rates:
                startyear_short_floors.append(min(cohort_rates))
                startyear_short_all_rates.extend(cohort_rates)

        perfect = bool(np.all(both))
        perfect_horizons += int(perfect)
        available_horizons += 1
        min_ret_win = min(min_ret_win, ret_win)
        min_mdd_win = min(min_mdd_win, mdd_win)
        min_both_win = min(min_both_win, both_win)
        min_worst_ret_excess = min(min_worst_ret_excess, worst_ret_excess)
        min_worst_mdd_improvement = min(min_worst_mdd_improvement, worst_mdd_improve)
        both_values.append(both_win)
        p10_return_excess_values.append(p10_ret_excess)
        if label in {"6M", "1Y", "2Y", "3Y"}: short_both_values.append(both_win)

        summaries.append({
            "Model": c.name,
            "Horizon": label,
            "Trading_Days": h,
            "Windows": nwin,
            "Return_Metric": return_metric,
            "Return_Win_Rate": ret_win,
            "MDD_Win_Rate": mdd_win,
            "Both_Win_Rate": both_win,
            "Perfect_All_Windows": perfect,
            "Strategy_Median_Return": float(np.median(strat_ann)),
            "TQQQ_Median_Return": float(np.median(bench_ann)),
            "Strategy_Worst_Return": float(np.min(strat_ann)),
            "TQQQ_Worst_Return": float(np.min(bench_ann)),
            "Worst_Return_Excess": worst_ret_excess,
            "P10_Return_Excess": p10_ret_excess,
            "Strategy_Median_MDD": float(np.median(strat_mdd)),
            "TQQQ_Median_MDD": float(np.median(bench_mdd)),
            "Strategy_Worst_MDD": float(np.min(strat_mdd)),
            "TQQQ_Worst_MDD": float(np.min(bench_mdd)),
            "Worst_MDD_Improvement": worst_mdd_improve,
        })

        if include_details:
            for s in range(nwin):
                details.append({
                    "Model": c.name,
                    "Horizon": label,
                    "Start": dates[s].date(),
                    "End": dates[s + h].date(),
                    "Strategy_Return": float(strat_ann[s]),
                    "TQQQ_Return": float(bench_ann[s]),
                    "Return_Excess": float(strat_ann[s] - bench_ann[s]),
                    "Strategy_MDD": float(strat_mdd[s]),
                    "TQQQ_MDD": float(bench_mdd[s]),
                    "MDD_Improvement": float(strat_mdd[s] - bench_mdd[s]),
                    "Return_Win": bool(ret_win_mask[s]),
                    "MDD_Win": bool(mdd_win_mask[s]),
                    "Both_Win": bool(both[s]),
                })

    bench_ret = features["TQQQ_O2O_RET"][eval_start:-1]
    bmet = full_metrics(bench_ret)
    strict_all = bool(available_horizons > 0 and perfect_horizons == available_horizons)
    avg_both_win = float(np.mean(both_values)) if both_values else np.nan
    short_both_win = float(np.mean(short_both_values)) if short_both_values else np.nan
    min_p10_ret_excess = float(np.min(p10_return_excess_values)) if p10_return_excess_values else np.nan
    min_startyear_short_win = float(np.min(startyear_short_floors)) if startyear_short_floors else np.nan
    p20_startyear_short_win = float(np.quantile(startyear_short_all_rates, 0.20)) if startyear_short_all_rates else np.nan
    if startyear_short_all_rates:
        cohort_sorted = np.sort(np.asarray(startyear_short_all_rates, dtype=float))
        bottom3_startyear_short_win = float(np.mean(cohort_sorted[:min(3, len(cohort_sorted))]))
        bottom20_n = max(1, int(math.ceil(0.20 * len(cohort_sorted))))
        bottom20_startyear_short_win = float(np.mean(cohort_sorted[:bottom20_n]))
    else:
        bottom3_startyear_short_win = np.nan
        bottom20_startyear_short_win = np.nan

    # V18 exact score: preserve the V17 frontier, then reward progress toward 79.5%/80%
    # while keeping Worst-Horizon MDD Win >=97%.
    # Bottom-20% >=21% remains an important secondary target.
    final_score = (
        16.5 * min_both_win
        + 6.0 * min_ret_win
        + 3.6 * min_mdd_win
        + 1.2 * (avg_both_win if np.isfinite(avg_both_win) else 0.0)
        + 4.5 * (short_both_win if np.isfinite(short_both_win) else 0.0)
        + 2.4 * (bottom3_startyear_short_win if np.isfinite(bottom3_startyear_short_win) else 0.0)
        + 4.6 * (bottom20_startyear_short_win if np.isfinite(bottom20_startyear_short_win) else 0.0)
        + 1.2 * (p20_startyear_short_win if np.isfinite(p20_startyear_short_win) else 0.0)
        + 2.2 * max(-0.60, min(0.15, min_p10_ret_excess if np.isfinite(min_p10_ret_excess) else -0.60))
        + 1.3 * max(-2.00, min(0.15, min_worst_ret_excess))
        + 1.7 * max(-0.5, min(0.5, met["mdd"] - bmet["mdd"]))
        + 0.8 * max(-0.5, min(0.5, met["cagr"] - bmet["cagr"]))
        + 6.8 * max(-0.10, min(0.10, min_worst_mdd_improvement))
        + 0.35 * (perfect_horizons / max(available_horizons, 1))
        - 0.015 * max(0, c.filter_count - 3)
    )
    bottom20_safe = bottom20_startyear_short_win if np.isfinite(bottom20_startyear_short_win) else 0.0
    final_score -= 60.0 * max(0.0, 0.79 - min_both_win)
    final_score -= 36.0 * max(0.0, 0.97 - min_mdd_win)
    final_score -= 8.0 * max(0.0, 0.21 - bottom20_safe)
    dead_v13_modules = int(c.trough_recovery_mode != "off" and st.get("trough_recovery_entries", 0) == 0)
    dead_soft_vol = int(c.soft_vol_mode != "off" and st.get("soft_vol_entries", 0) == 0)
    dead_v14_modules = dead_v13_modules + dead_soft_vol
    dead_impulse_veto = int(c.impulse_veto_mode != "off" and st.get("impulse_veto_blocks", 0) == 0)
    dead_v15_modules = dead_v14_modules + dead_impulse_veto
    dead_precision_veto = int(c.precision_veto_mode != "off" and st.get("precision_veto_blocks", 0) == 0)
    dead_v16_modules = dead_v15_modules + dead_precision_veto
    dead_reentry_grace = int(c.reentry_grace_mode != "off" and st.get("reentry_grace_blocks", 0) == 0)
    dead_v17_modules = dead_v16_modules + dead_reentry_grace
    dead_micro_rescue = int(c.micro_rescue_mode != "off" and st.get("micro_rescue_blocks", 0) == 0)
    dead_v19_modules = dead_v17_modules + dead_micro_rescue
    dead_shadow_rescue = int(c.shadow_rescue_mode != "off" and st.get("shadow_rescue_blocks", 0) == 0)
    dead_v20_modules = dead_v19_modules + dead_shadow_rescue
    dead_preview_overlay = int(c.preview_overlay_mode != "off" and st.get("preview_overlay_blocks", 0) == 0)
    dead_v21_modules = dead_v20_modules + dead_preview_overlay
    final_score -= 0.06 * dead_v21_modules
    # Prefer a few surgical extra blocks over broad exit hysteresis.
    final_score -= 0.03 * max(0, st.get("precision_veto_blocks", 0) - 4)
    grace_free_blocks = 6 if c.reentry_grace_mode == "split_reentry" else 4
    final_score -= 0.03 * max(0, st.get("reentry_grace_blocks", 0) - grace_free_blocks)
    final_score -= 0.05 * max(0, st.get("micro_rescue_blocks", 0) - 4)
    final_score -= 0.05 * max(0, st.get("shadow_rescue_blocks", 0) - 4)
    final_score -= 0.05 * max(0, st.get("preview_overlay_blocks", 0) - 6)
    joint_79_97 = bool(min_both_win >= 0.79 and min_mdd_win >= 0.97)
    joint_gap_79_97 = max(0.0, 0.79 - min_both_win) + max(0.0, 0.97 - min_mdd_win)
    joint_795_97 = bool(min_both_win >= 0.795 and min_mdd_win >= 0.97)
    joint_80_97 = bool(min_both_win >= 0.80 and min_mdd_win >= 0.97)
    joint_81_97 = bool(min_both_win >= 0.81 and min_mdd_win >= 0.97)
    joint_82_97 = bool(min_both_win >= 0.82 and min_mdd_win >= 0.97)
    if joint_79_97:
        final_score += 0.80
    if joint_795_97:
        final_score += 0.65
    if joint_80_97:
        final_score += 1.30
    if joint_81_97:
        final_score += 1.50
    if joint_82_97:
        final_score += 1.75
    if strict_all:
        final_score += 100.0

    row = {
        "name": c.name,
        "family": c.family,
        "final_score": float(final_score),
        "strict_all_windows_both_win": strict_all,
        "perfect_horizons": perfect_horizons,
        "available_horizons": available_horizons,
        "min_both_win_rate": min_both_win if available_horizons else np.nan,
        "min_return_win_rate": min_ret_win if available_horizons else np.nan,
        "min_mdd_win_rate": min_mdd_win if available_horizons else np.nan,
        "avg_both_win_rate": avg_both_win,
        "short_both_win_rate": short_both_win,
        "min_startyear_short_win_rate": min_startyear_short_win,
        "bottom3_startyear_short_win_rate": bottom3_startyear_short_win,
        "bottom20_startyear_short_win_rate": bottom20_startyear_short_win,
        "p20_startyear_short_win_rate": p20_startyear_short_win,
        "target_both_79": bool(min_both_win >= 0.79),
        "target_mdd_97": bool(min_mdd_win >= 0.97),
        "target_joint_79_97": joint_79_97,
        "target_joint_795_97": joint_795_97,
        "target_joint_80_97": joint_80_97,
        "target_joint_81_97": joint_81_97,
        "target_joint_82_97": joint_82_97,
        "joint_gap_79_97": float(joint_gap_79_97),
        "core_targets_passed": int(min_both_win >= 0.79) + int(min_mdd_win >= 0.97),
        "target_bottom20_21": bool(np.isfinite(bottom20_startyear_short_win) and bottom20_startyear_short_win >= 0.21),
        "targets_passed": int(min_both_win >= 0.79) + int(min_mdd_win >= 0.97) + int(np.isfinite(bottom20_startyear_short_win) and bottom20_startyear_short_win >= 0.21),
        "worst_horizon_label": worst_horizon_label,
        "worst_horizon_windows": worst_horizon_windows,
        "worst_horizon_both_wins": worst_horizon_both_wins,
        "worst_horizon_wins_needed_79": worst_horizon_wins_needed_79,
        "worst_horizon_wins_needed_795": worst_horizon_wins_needed_795,
        "worst_horizon_wins_needed_80": worst_horizon_wins_needed_80,
        "one_year_windows": one_year_windows,
        "one_year_both_wins": one_year_both_wins,
        "one_year_wins_needed_79": one_year_wins_needed_79,
        "one_year_wins_needed_795": one_year_wins_needed_795,
        "one_year_wins_needed_80": one_year_wins_needed_80,
        "one_year_near_miss_25bp": one_year_near_miss_25bp,
        "one_year_near_miss_50bp": one_year_near_miss_50bp,
        "one_year_near_miss_100bp": one_year_near_miss_100bp,
        "two_year_windows": two_year_windows,
        "two_year_both_wins": two_year_both_wins,
        "two_year_wins_needed_80": two_year_wins_needed_80,
        "two_year_wins_needed_81": two_year_wins_needed_81,
        "two_year_wins_needed_82": two_year_wins_needed_82,
        "two_year_near_miss_25bp": two_year_near_miss_25bp,
        "two_year_near_miss_50bp": two_year_near_miss_50bp,
        "two_year_near_miss_100bp": two_year_near_miss_100bp,
        "dead_v21_modules": dead_v21_modules,
        "dead_v20_modules": dead_v20_modules,
        "dead_v19_modules": dead_v19_modules,
        "dead_v17_modules": dead_v17_modules,
        "dead_v16_modules": dead_v16_modules,
        "dead_v15_modules": dead_v15_modules,
        "dead_v14_modules": dead_v14_modules,
        "dead_v13_modules": dead_v13_modules,
        "dead_v12_modules": dead_v13_modules,
        "min_p10_return_excess": min_p10_ret_excess,
        "worst_return_excess": min_worst_ret_excess if available_horizons else np.nan,
        "worst_mdd_improvement": min_worst_mdd_improvement if available_horizons else np.nan,
        "cagr": met["cagr"],
        "mdd": met["mdd"],
        "calmar": met["calmar"],
        "sharpe": met["sharpe"],
        "total_return": met["total_return"],
        "bench_cagr": bmet["cagr"],
        "bench_mdd": bmet["mdd"],
        "cagr_excess": met["cagr"] - bmet["cagr"],
        "mdd_improvement": met["mdd"] - bmet["mdd"],
        "gates": config_to_short_text(c),
        "config": asdict(c),
        **st,
    }
    return row, summaries, details


def build_benchmark_cache(features: Dict[str, Any], eval_start: int) -> Dict[int, Dict[str, np.ndarray]]:
    rr = features["TQQQ_O2O_RET"][eval_start:-1]
    log_eq = np.r_[0.0, np.cumsum(np.log1p(np.clip(rr, -0.999999, None)))]
    tree = DrawdownSegmentTree(log_eq)
    cache: Dict[int, Dict[str, np.ndarray]] = {}
    for _, h in HORIZONS:
        if len(rr) < h:
            continue
        nwin = len(rr) - h + 1
        logret = log_eq[h:] - log_eq[:-h]
        mdd = np.empty(nwin, dtype=float)
        for s in range(nwin):
            mdd[s] = math.exp(-tree.query_drop(s, s + h)) - 1.0
        cache[h] = {"logret": logret, "mdd": mdd}
    return cache


# -----------------------------
# Ranking / indicator contribution
# -----------------------------

def choose_finalists(search_rows: List[Dict[str, Any]], n: int) -> List[Config]:
    good = [r for r in search_rows if "error" not in r and np.isfinite(r.get("search_score", np.nan))]
    good.sort(key=lambda x: x["search_score"], reverse=True)
    selected: List[Dict[str, Any]] = []
    names = set()

    # Always preserve the key baselines/seeds if present.
    priority_names = {
        "ROBUST_ORIGINAL", "ROBUST_NO_FAST", "ROBUST_FAST_TO_QLD",
        "ROBUST_V15_ENTRY", "ROBUST_V15_ENTRY_FAST_QLD", "ROBUST_RISK_CAP2_8", "ROBUST_TIGHT_COMBO",
        "V1_RND_000171", "V1_RND_000171_TQQQ_ONLY", "V1_RND_004288", "V1_RND_005848",
        "V2_FOCUS_031316", "V2_FOCUS_018911", "V2_FOCUS_025256", "V2_FOCUS_027205", "V2_FOCUS_031870", "V2_031316_RECOVERY",
        "V4_FOCUS_204252", "V5_FOCUS_366726", "V5_FOCUS_267771", "V6_366726_2OF3", "V6_366726_ADAPT", "V6_366726_SNAP",
        "V6_FOCUS_107803", "V7_107803_SNAP_DD25", "V7_107803_CR_Q", "V7_FOCUS_326611",
        "V8_STAGE_BALANCED", "V8_EXPLOSIVE", "V8_COMBO_BALANCED", "V8_COMBO_PLUS_SNAP",
        "V8_FOCUS_211857", "V8_FOCUS_291951",
        "V9_FOCUS_128967", "V9_FOCUS_147025", "V9_FOCUS_207661",
        "V10_HYST_STABLE_S3", "V10_HYST_STABLE_S4", "V10_HYST_STABLE_S5", "V10_HYST_STABLE_S6",
        "V10_DDR_R085_RSI48_R506", "V10_DDR_R090_RSI48_R506", "V10_DDR_R095_RSI48_R506", "V10_DDR_R100_RSI48_R506",
        "V11_CORE_211857", "V11_TR2_D3_A3_R090", "V11_TR2_D3_A4_R090", "V11_TR2_D3_A5_R095", "V11_TR2_D3_A6_R095",
        "V11_TR2_D3_A4_R095_R10P2", "V11_TR2_G3_A4_R095", "V11_TR2_Q3_A3_R090_C1", "V11_TR2_Q3_A4_R095_C1", "V11_TR2_Q3_A4_R095_C2", "V11_TR2_Q3_A5_R095_C2",
        "V11_FOCUS_000365", "V11_FOCUS_003898", "V11_FOCUS_012077",
        "V12_SPLIT_BASE", "V12_SPLIT_EARLY_Q", "V12_SPLIT_STRONG3", "V12_SPLIT_PATIENT",
        "V15_FOCUS_000979", "V15_FOCUS_006433", "V17_FOCUS_471542",
        "V17_RG_MOM_A17_CORE", "V17_RG_MOM_A20_LOOSE", "V17_RG_MOM_A13_TIGHT", "V17_RG_S2_A20_PASS",
        "V18_SPLIT_DAY1", "V18_SPLIT_TREND2", "V18_FOCUS_011262",
        "V19_MICRO_S1_PULLBACK", "V19_MICRO_S1_NEARTREND", "V19_MICRO_MATURE_MOM", "V19_MICRO_COMBO",
        "V19_FOCUS_001037", "V20_SHADOW_S1_RECENT", "V20_SHADOW_MATURE_MOM", "V20_SHADOW_COMBO", "V20_FOCUS_SHADOW_COMBO",
        "V21_PREVIEW_QLD3", "V21_PREVIEW_QLD2", "V21_PREVIEW_STEPUP", "V21_PREVIEW_COMBO", "V21_PREVIEW_CANARY",
    }
    for r in good:
        if (r["name"] in priority_names or r["name"] == "V12_FOCUS_011962" or r.get("family") in {"v13_core_grid", "v15_impulse_probe", "v16_precision_probe", "v17_reentry_probe", "v18_split_probe", "v19_micro_probe", "v20_shadow_probe", "v21_preview_probe"}) and r["name"] not in names:
            selected.append(r); names.add(r["name"])

    # V18 force-lane: exact-MDD validate high-1Y-win candidates even when their aggregate
    # search score is slightly lower. This prevents the V17 2,906-win rough candidates from
    # being discarded before exact BOTH/MDD analysis.
    one_y_lane = sorted(good, key=lambda x: (x.get("ret_win_1Y", 0.0), x.get("min_return_win_rate", 0.0), x.get("search_score", -1e9)), reverse=True)
    lane_n = max(24, n // 3)
    for r in one_y_lane[:lane_n]:
        if len(selected) >= n: break
        if r["name"] not in names:
            selected.append(r); names.add(r["name"])

    # V20 bottleneck lane: after V19, 2Y is the worst horizon. Force high-2Y
    # candidates into exact rolling-MDD analysis even if aggregate score is lower.
    two_y_lane = sorted(good, key=lambda x: (x.get("ret_win_2Y", 0.0), x.get("ret_win_1Y", 0.0), x.get("search_score", -1e9)), reverse=True)
    lane2_n = max(36, n // 3)
    for r in two_y_lane[:lane2_n]:
        if len(selected) >= n: break
        if r["name"] not in names:
            selected.append(r); names.add(r["name"])

    # Then best search results, with a small diversity rule by gate signature.
    sigs = set()
    for r in good:
        if len(selected) >= max(n, len(selected)) and len(selected) >= n:
            break
        if r["name"] in names:
            continue
        cfg = Config(**r["config"])
        sig = (cfg.fast_mode, cfg.gate_scope, cfg.defense_mode, cfg.tqqq_exit_mode, cfg.staged_reentry_mode, cfg.explosive_rebound_mode, cfg.correction_reentry_mode, cfg.dd_rebound_mode, cfg.trough_recovery_mode, cfg.repromote_guard_mode, cfg.soft_vol_mode, cfg.impulse_veto_mode, cfg.precision_veto_mode, cfg.reentry_grace_mode, cfg.micro_rescue_mode, cfg.shadow_rescue_mode, cfg.preview_overlay_mode, cfg.adaptive_down_mode, cfg.snapback_mode, cfg.macd_gate, cfg.rsi_asset, cfg.stoch_asset, cfg.boll_asset, cfg.adx_asset)
        if sig in sigs and len(selected) < n // 2:
            continue
        sigs.add(sig)
        selected.append(r); names.add(r["name"])
        if len(selected) >= n:
            break

    # If priority seeds pushed count above n, keep them all; otherwise n exactly.
    return [Config(**r["config"]) for r in selected]


def indicator_stats(final_rows: List[Dict[str, Any]], top_n: int = 10) -> pd.DataFrame:
    rows = sorted(final_rows, key=lambda x: x["final_score"], reverse=True)[:top_n]
    indicators = [
        ("QQQ > MA50", lambda c: c.require_qqq_ma50),
        ("QQQ > MA200", lambda c: c.require_qqq_ma200),
        ("TQQQ > MA20", lambda c: c.require_tqqq_ma20),
        ("TQQQ > MA50", lambda c: c.require_tqqq_ma50),
        ("MACD", lambda c: c.macd_gate != "none"),
        ("RSI", lambda c: c.rsi_asset != "none"),
        ("Stochastic", lambda c: c.stoch_asset != "none"),
        ("CCI", lambda c: c.cci_asset != "none"),
        ("Bollinger", lambda c: c.boll_asset != "none"),
        ("ADX/DMI", lambda c: c.adx_asset != "none"),
        ("MFI", lambda c: c.mfi_asset != "none"),
        ("Williams %R", lambda c: c.willr_asset != "none"),
        ("OBV", lambda c: c.obv_asset != "none"),
        ("Donchian", lambda c: c.donchian_asset != "none"),
        ("ATR", lambda c: c.atr_asset != "none"),
        ("ROC", lambda c: c.roc_asset != "none"),
        ("Risk Defense", lambda c: c.defense_mode != "off"),
        ("Fast->QLD", lambda c: c.fast_mode == "qld"),
        ("Fast conditional", lambda c: c.fast_mode == "conditional"),
        ("TQQQ-only gates", lambda c: c.gate_scope == "tqqq_only"),
        ("Fast-only gates", lambda c: c.gate_scope == "fast_only"),
        ("TQQQ immediate exit", lambda c: c.tqqq_exit_mode != "off"),
        ("Rebound bypass", lambda c: c.rebound_bypass_mode != "off"),
        ("Medium-trend bridge", lambda c: c.bridge_mode != "off"),
        ("RMA gate 2-of-3", lambda c: c.gate_logic == "rma_2of3"),
        ("Adaptive down confirm", lambda c: c.adaptive_down_mode != "off"),
        ("Generic snapback", lambda c: c.snapback_mode != "off"),
        ("V7 crash release", lambda c: c.crash_release_mode != "off"),
        ("V8 staged QLD->TQQQ", lambda c: c.staged_reentry_mode != "off"),
        ("V8 explosive rebound", lambda c: c.explosive_rebound_mode != "off"),
        ("Correction re-entry (legacy/control)", lambda c: c.correction_reentry_mode != "off"),
        ("Legacy V9/V10 DDR", lambda c: c.dd_rebound_mode != "off"),
        ("V11/V12 trough recovery", lambda c: c.trough_recovery_mode != "off"),
        ("V12 split QLD promotion", lambda c: c.trough_recovery_mode == "qld_split3"),
        ("V14 soft-vol re-promotion", lambda c: c.soft_vol_mode != "off"),
        ("V15 selective impulse exit veto", lambda c: c.impulse_veto_mode != "off"),
        ("V16 precision transition veto", lambda c: c.precision_veto_mode != "off"),
        ("V17 recent re-entry grace", lambda c: c.reentry_grace_mode != "off"),
        ("V12 ordinary re-promotion guard", lambda c: c.repromote_guard_mode != "off"),
        ("Fast disabled", lambda c: c.fast_mode == "normal"),
        ("Down confirm=1", lambda c: c.down_confirm == 1),
    ]
    out = []
    for label, fn in indicators:
        active_scores = []
        active_both = []
        count = 0
        for r in rows:
            c = Config(**r["config"])
            if fn(c):
                count += 1
                active_scores.append(r["final_score"])
                active_both.append(r["min_both_win_rate"])
        out.append({
            "Indicator_or_Module": label,
            "TopN_Active_Count": count,
            "TopN_Active_Rate": count / max(len(rows), 1),
            "Avg_Final_Score_When_Active": float(np.mean(active_scores)) if active_scores else np.nan,
            "Avg_Min_Both_Win_Rate_When_Active": float(np.mean(active_both)) if active_both else np.nan,
        })
    return pd.DataFrame(out).sort_values(["TopN_Active_Count", "Avg_Final_Score_When_Active"], ascending=[False, False])


def build_cohort_floor_table(rolling_details: pd.DataFrame) -> pd.DataFrame:
    """Aggregate detailed 6M-3Y rolling windows by start year; no year is hardcoded."""
    if rolling_details is None or rolling_details.empty:
        return pd.DataFrame()
    d = rolling_details.copy()
    d["Start"] = pd.to_datetime(d["Start"])
    d["Start_Year"] = d["Start"].dt.year
    d = d[d["Horizon"].isin(["6M","1Y","2Y","3Y"])].copy()
    if d.empty:
        return pd.DataFrame()
    g = d.groupby(["Model","Horizon","Start_Year"], as_index=False).agg(
        Windows=("Both_Win","size"),
        Return_Win_Rate=("Return_Win","mean"),
        MDD_Win_Rate=("MDD_Win","mean"),
        Both_Win_Rate=("Both_Win","mean"),
        Avg_Return_Excess=("Return_Excess","mean"),
        Worst_Return_Excess=("Return_Excess","min"),
        Avg_MDD_Improvement=("MDD_Improvement","mean"),
        Worst_MDD_Improvement=("MDD_Improvement","min"),
    )
    return g.sort_values(["Model","Both_Win_Rate","Horizon","Start_Year"], ascending=[True,True,True,True]).reset_index(drop=True)


def build_reentry_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 5) -> pd.DataFrame:
    """Write generic crash/correction context rows for diagnosis of blocked re-entry.

    Rows cover bounded post-crash and post-correction windows for the selected models.
    This intentionally avoids any calendar-specific event selection.
    """
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"])
    q = f["QQQ_CLOSE"]
    dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        _, states, _ = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        sma = f[f"QQQ_SMA{c.sma_days}"]
        mom = f[f"QQQ_ROC{c.mom_days}"]
        vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        crash = valid & (q <= sma) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        entry_mask = entry_mask_for_config(c, f)
        days_crash = 10**9
        days_corr = 10**9
        trough = np.nan
        tr2_trough = np.nan
        tr2_trough_idx = -1
        tr2_episode_crash_days = 0
        for t in range(len(q)-1):
            cur, nxt = int(states[t]), int(states[t+1])
            if crash[t]:
                days_crash = 0
                days_corr = 10**9
                trough = dd252[t] if not np.isfinite(trough) else min(trough, dd252[t])
                tr2_episode_crash_days += 1
                if (not np.isfinite(tr2_trough)) or dd252[t] < tr2_trough - 1e-12:
                    tr2_trough = dd252[t]; tr2_trough_idx = t
            else:
                days_crash = min(days_crash + 1, 10**9)
                if np.isfinite(tr2_trough) and days_crash <= c.trough_recovery_window_days and dd252[t] < tr2_trough - 1e-12:
                    tr2_trough = dd252[t]; tr2_trough_idx = t
                if days_corr < 10**9:
                    days_corr += 1
            corr_delever = (cur == 3 and nxt < 3 and not bool(crash[t]))
            crash_delever = (cur == 3 and nxt < 3 and bool(crash[t]))
            if corr_delever:
                days_corr = 0
            if nxt == 3:
                days_corr = 10**9
            if days_crash > 63 and not crash[t]:
                trough = np.nan
            if days_crash > c.trough_recovery_window_days and not crash[t]:
                tr2_trough = np.nan; tr2_trough_idx = -1; tr2_episode_crash_days = 0
            in_crash_ctx = bool(crash[t]) or days_crash <= 63
            in_corr_ctx = days_corr <= 30
            state_change = cur != nxt
            if not (in_crash_ctx or in_corr_ctx or state_change):
                continue
            dd_rebound = (dd252[t] - trough) if np.isfinite(trough) else np.nan
            old_exp_parts = [
                q[t] > f["QQQ_EMA5"][t],
                f["QQQ_ROC5"][t] >= c.explosive_roc5_min,
                f["QQQ_ROC10"][t] >= c.explosive_roc10_min,
                f["QQQ_RSI14"][t] >= c.explosive_rsi_min,
                f["QQQ_MFI14"][t] >= c.explosive_mfi_min,
                f["QQQ_ATRPCT14"][t] <= c.explosive_atr_max,
                vol[t] <= c.explosive_vol_max,
            ]
            tr2_age = (t - tr2_trough_idx) if tr2_trough_idx >= 0 else None
            tr2_rebound = (dd252[t] - tr2_trough) if np.isfinite(tr2_trough) else np.nan
            tr2_parts = [
                tr2_age is not None and tr2_age >= c.trough_age_min,
                tr2_episode_crash_days >= c.trough_min_crash_days,
                np.isfinite(tr2_rebound) and tr2_rebound >= c.trough_rebound_min,
                q[t] > f["QQQ_EMA5"][t],
                f["QQQ_ROC5"][t] >= c.trough_roc5_min,
                f["QQQ_ROC10"][t] >= c.trough_roc10_min,
                f["QQQ_RSI14"][t] >= c.trough_rsi_min,
                f["QQQ_MFI14"][t] >= c.trough_mfi_min,
                f["QQQ_ATRPCT14"][t] <= c.trough_atr_max,
                vol[t] <= c.trough_vol_max,
            ]
            tr2_promote_parts = [
                np.isfinite(tr2_rebound) and tr2_rebound >= c.trough_promote_rebound_min,
                q[t] > f["QQQ_EMA5"][t],
                f["QQQ_ROC5"][t] >= c.trough_promote_roc5_min,
                f["QQQ_ROC10"][t] >= c.trough_promote_roc10_min,
                f["QQQ_RSI14"][t] >= c.trough_promote_rsi_min,
                f["QQQ_MFI14"][t] >= c.trough_promote_mfi_min,
                f["QQQ_ATRPCT14"][t] <= c.trough_promote_atr_max,
                vol[t] <= c.trough_promote_vol_max,
            ]
            ddr_parts = [
                np.isfinite(dd_rebound) and dd_rebound >= c.dd_rebound_min,
                q[t] > f["QQQ_EMA5"][t],
                f["QQQ_ROC5"][t] >= c.dd_rebound_roc5_min,
                f["QQQ_RSI14"][t] >= c.dd_rebound_rsi_min,
                f["QQQ_MFI14"][t] >= c.dd_rebound_mfi_min,
                f["QQQ_ATRPCT14"][t] <= c.dd_rebound_atr_max,
                vol[t] <= c.dd_rebound_vol_max,
            ]
            corr_parts = [
                q[t] > f[f"QQQ_EMA{c.correction_ema_days}"][t],
                f["QQQ_ROC5"][t] >= c.correction_roc5_min,
                f["QQQ_ROC10"][t] >= c.correction_roc10_min,
                f["QQQ_RSI14"][t] >= c.correction_rsi_min,
                f["QQQ_MFI14"][t] >= c.correction_mfi_min,
                f["QQQ_ATRPCT14"][t] <= c.correction_atr_max,
                vol[t] <= c.correction_vol_max,
                dd252[t] >= c.correction_dd_floor,
            ]
            context = "BOTH" if in_crash_ctx and in_corr_ctx else "POST_CRASH" if in_crash_ctx else "CORRECTION"
            if state_change:
                context += "+STATE_CHANGE"
            rows.append({
                "Model": c.name, "Date": dates[t].date(), "Context": context,
                "State_From": STATE_NAME[cur], "State_To": STATE_NAME[nxt],
                "Crash_Flag": bool(crash[t]), "Days_Since_Crash": days_crash if days_crash < 10**8 else None,
                "Days_Since_Correction": days_corr if days_corr < 10**8 else None,
                "Correction_Delever": corr_delever, "Crash_Delever": crash_delever,
                "QQQ_DD252": float(dd252[t]), "Crash_Trough_DD": float(trough) if np.isfinite(trough) else np.nan,
                "DD_Rebound_From_Trough": float(dd_rebound) if np.isfinite(dd_rebound) else np.nan,
                "TR2_Trough_DD": float(tr2_trough) if np.isfinite(tr2_trough) else np.nan,
                "TR2_Trough_Age": tr2_age, "TR2_Episode_Crash_Days": tr2_episode_crash_days,
                "TR2_Rebound_From_Trough": float(tr2_rebound) if np.isfinite(tr2_rebound) else np.nan,
                "QQQ_VOL": float(vol[t]), "QQQ_ROC5": float(f["QQQ_ROC5"][t]), "QQQ_ROC10": float(f["QQQ_ROC10"][t]),
                "QQQ_ROC20": float(f["QQQ_ROC20"][t]), "QQQ_RSI14": float(f["QQQ_RSI14"][t]),
                "QQQ_MFI14": float(f["QQQ_MFI14"][t]), "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
                "Entry_Gate": bool(entry_mask[t]),
                "Old_EXP_All_Pass": bool(all(old_exp_parts)) if c.explosive_rebound_mode != "off" else None,
                "DDR_All_Pass": bool(all(ddr_parts)) if c.dd_rebound_mode != "off" else None,
                "TR2_Age_OK": bool(tr2_parts[0]) if c.trough_recovery_mode != "off" else None,
                "TR2_CrashDays_OK": bool(tr2_parts[1]) if c.trough_recovery_mode != "off" else None,
                "TR2_Rebound_OK": bool(tr2_parts[2]) if c.trough_recovery_mode != "off" else None,
                "TR2_ROC5_OK": bool(tr2_parts[4]) if c.trough_recovery_mode != "off" else None,
                "TR2_ROC10_OK": bool(tr2_parts[5]) if c.trough_recovery_mode != "off" else None,
                "TR2_RSI_OK": bool(tr2_parts[6]) if c.trough_recovery_mode != "off" else None,
                "TR2_All_Pass": bool(all(tr2_parts)) if c.trough_recovery_mode != "off" else None,
                "TR2_Promote_Rebound_OK": bool(tr2_promote_parts[0]) if c.trough_recovery_mode == "qld_split3" else None,
                "TR2_Promote_ROC5_OK": bool(tr2_promote_parts[2]) if c.trough_recovery_mode == "qld_split3" else None,
                "TR2_Promote_ROC10_OK": bool(tr2_promote_parts[3]) if c.trough_recovery_mode == "qld_split3" else None,
                "TR2_Promote_RSI_OK": bool(tr2_promote_parts[4]) if c.trough_recovery_mode == "qld_split3" else None,
                "TR2_Promote_MFI_OK": bool(tr2_promote_parts[5]) if c.trough_recovery_mode == "qld_split3" else None,
                "TR2_Promote_ATR_OK": bool(tr2_promote_parts[6]) if c.trough_recovery_mode == "qld_split3" else None,
                "TR2_Promote_VOL_OK": bool(tr2_promote_parts[7]) if c.trough_recovery_mode == "qld_split3" else None,
                "TR2_Promote_All_Pass": bool(all(tr2_promote_parts)) if c.trough_recovery_mode == "qld_split3" else None,
                "CORR_All_Pass": bool(all(corr_parts)) if c.correction_reentry_mode != "off" else None,
            })
    return pd.DataFrame(rows)



def build_transition_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 6) -> pd.DataFrame:
    """Generic state-transition / mild-wobble audit retained for cross-version diagnosis."""
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"])
    q = f["QQQ_CLOSE"]
    dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        _, states, _ = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        sma = f[f"QQQ_SMA{c.sma_days}"]; mom = f[f"QQQ_ROC{c.mom_days}"]; vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        trend_ok = q > sma; mom_ok = mom > c.mom_threshold; vol_ok = vol < c.vol_threshold
        score = trend_ok.astype(np.int8) + mom_ok.astype(np.int8) + vol_ok.astype(np.int8)
        crash = valid & (~trend_ok) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        entry = entry_mask_for_config(c, f)
        for t in range(len(q)-1):
            if not valid[t]:
                continue
            cur, nxt = int(states[t]), int(states[t+1])
            changed = cur != nxt
            mild_wobble = bool(c.adaptive_down_mode != "off" and cur == 3 and score[t] == 2 and not crash[t])
            if not (changed or mild_wobble):
                continue
            if changed:
                event = "TQQQ_DELEVER" if cur == 3 and nxt < 3 else "TQQQ_REENTER" if cur < 3 and nxt == 3 else "STATE_CHANGE"
            else:
                event = "HYST_MILD_HOLD"
            rows.append({
                "Model": c.name, "Date": dates[t].date(), "Event": event,
                "State_From": STATE_NAME[cur], "State_To": STATE_NAME[nxt], "State_Changed": changed,
                "Base_Score": int(score[t]), "Trend_OK": bool(trend_ok[t]), "Mom_OK": bool(mom_ok[t]), "Vol_OK": bool(vol_ok[t]),
                "Crash_Flag": bool(crash[t]), "Entry_Gate": bool(entry[t]),
                "Adaptive_Mode": c.adaptive_down_mode, "Soft_Down_Confirm": c.soft_down_confirm, "Hard_Down_Confirm": c.hard_down_confirm,
                "QQQ_DD252": float(dd252[t]), "QQQ_VOL": float(vol[t]),
                "QQQ_ROC5": float(f["QQQ_ROC5"][t]), "QQQ_ROC10": float(f["QQQ_ROC10"][t]), "QQQ_ROC120": float(mom[t]),
                "QQQ_RSI14": float(f["QQQ_RSI14"][t]), "QQQ_MFI14": float(f["QQQ_MFI14"][t]), "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
            })
    return pd.DataFrame(rows)


def build_tqqq_reentry_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 6) -> pd.DataFrame:
    """Audit ordinary TQQQ delever/re-entry cycles without changing strategy logic.

    This is intentionally diagnostic only: it records re-entry gaps, base score, volatility
    utilization, and whether a TQQQ re-entry happens within 10/20 sessions of a prior delever.
    No year/event is hardcoded.
    """
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"])
    q = f["QQQ_CLOSE"]
    dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        _, states, _ = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        sma = f[f"QQQ_SMA{c.sma_days}"]
        mom = f[f"QQQ_ROC{c.mom_days}"]
        vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        trend_ok = q > sma; mom_ok = mom > c.mom_threshold; vol_ok = vol < c.vol_threshold
        score = trend_ok.astype(np.int8) + mom_ok.astype(np.int8) + vol_ok.astype(np.int8)
        crash = valid & (~trend_ok) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        last_delever_t = None
        ordinary_delevers: List[int] = []
        days_crash = 10**9
        for t in range(len(q)-1):
            if crash[t]:
                days_crash = 0
            else:
                days_crash = min(days_crash + 1, 10**9)
            cur, nxt = int(states[t]), int(states[t+1])
            delever = cur == 3 and nxt < 3
            reenter = cur < 3 and nxt == 3
            if not (delever or reenter):
                continue
            gap = (t - last_delever_t) if (reenter and last_delever_t is not None) else None
            recent_guard_delevers = sum(1 for di in ordinary_delevers if 0 < (t - di) <= max(1, c.repromote_lookback_days))
            vol_util = float(vol[t] / c.vol_threshold) if np.isfinite(vol[t]) and c.vol_threshold > 0 else np.nan
            guard_strong = bool(
                valid[t] and score[t] >= 2 and f["QQQ_ROC10"][t] >= c.repromote_roc10_min
                and f["QQQ_RSI14"][t] >= c.repromote_rsi_min
                and np.isfinite(vol_util) and vol_util <= c.repromote_vol_util_max
            )
            guard_regime = bool(c.repromote_guard_mode != "off" and recent_guard_delevers >= max(1, c.repromote_trigger_count))
            if delever:
                last_delever_t = t
                if not bool(crash[t]):
                    ordinary_delevers.append(t)
                    ordinary_delevers = [di for di in ordinary_delevers if t - di <= max(120, c.repromote_lookback_days)]
            rows.append({
                "Model": c.name, "Date": dates[t].date(),
                "Event": "TQQQ_DELEVER" if delever else "TQQQ_REENTER",
                "State_From": STATE_NAME[cur], "State_To": STATE_NAME[nxt],
                "Crash_Flag": bool(crash[t]), "Base_Score": int(score[t]) if valid[t] else None,
                "Trend_OK": bool(trend_ok[t]) if valid[t] else None,
                "Mom_OK": bool(mom_ok[t]) if valid[t] else None,
                "Vol_OK": bool(vol_ok[t]) if valid[t] else None,
                "Reentry_Gap_Sessions": gap,
                "Reenter_Within_10d": bool(gap is not None and gap <= 10),
                "Reenter_Within_20d": bool(gap is not None and gap <= 20),
                "QQQ_DD252": float(dd252[t]) if np.isfinite(dd252[t]) else np.nan,
                "QQQ_VOL": float(vol[t]) if np.isfinite(vol[t]) else np.nan,
                "Vol_Threshold": float(c.vol_threshold),
                "Vol_Utilization": vol_util,
                "Near_Vol_Ceiling_90pct": bool(np.isfinite(vol[t]) and vol[t] >= 0.90 * c.vol_threshold),
                "Recent_Ordinary_Delevers": recent_guard_delevers,
                "Repromote_Guard_Mode": c.repromote_guard_mode,
                "Repromote_Guard_Regime": guard_regime,
                "Repromote_Strong_OK": guard_strong,
                "Days_Since_Crash": days_crash if days_crash < 10**8 else None,
                "QQQ_ROC5": float(f["QQQ_ROC5"][t]), "QQQ_ROC10": float(f["QQQ_ROC10"][t]),
                "QQQ_ROC120": float(f["QQQ_ROC120"][t]), "QQQ_RSI14": float(f["QQQ_RSI14"][t]),
                "QQQ_MFI14": float(f["QQQ_MFI14"][t]), "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
            })
    return pd.DataFrame(rows)



def build_impulse_veto_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 6) -> pd.DataFrame:
    """V15 diagnostic: show dates where the impulse-veto actually changed a TQQQ->QLD exit.

    Each V15 candidate is compared with an otherwise identical control whose impulse-veto
    is disabled. This keeps the audit diagnostic-only and makes the event attribution easy
    to verify from the final state arrays.
    """
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"])
    q = f["QQQ_CLOSE"]
    dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        if c.impulse_veto_mode == "off":
            continue
        _, states_v15, stats_v15 = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        control_dict = asdict(c)
        control_dict["impulse_veto_mode"] = "off"
        control = Config(**control_dict)
        _, states_control, _ = simulate_strategy(control, f, float(cost_bps) / 10000.0)

        sma = f[f"QQQ_SMA{c.sma_days}"]
        mom = f[f"QQQ_ROC{c.mom_days}"]
        vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        trend_ok = q > sma
        mom_ok = mom > c.mom_threshold
        vol_ok = vol < c.vol_threshold
        score = trend_ok.astype(np.int8) + mom_ok.astype(np.int8) + vol_ok.astype(np.int8)
        crash = valid & (~trend_ok) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        entry = entry_mask_for_config(c, f)

        days_crash = 10**9
        for t in range(len(q) - 1):
            if crash[t]:
                days_crash = 0
            else:
                days_crash = min(days_crash + 1, 10**9)
            actual_block = bool(
                int(states_control[t]) == 3 and int(states_control[t + 1]) == 2
                and int(states_v15[t]) == 3 and int(states_v15[t + 1]) == 3
            )
            if not actual_block:
                continue
            rows.append({
                "Model": c.name,
                "Date": dates[t].date(),
                "Event": "V15_IMPULSE_VETO_BLOCK",
                "Control_Next": STATE_NAME[int(states_control[t + 1])],
                "V15_Next": STATE_NAME[int(states_v15[t + 1])],
                "Mode": c.impulse_veto_mode,
                "Max_Veto_Days": c.impulse_veto_max_days,
                "Base_Score": int(score[t]) if valid[t] else None,
                "Trend_OK": bool(trend_ok[t]) if valid[t] else None,
                "Mom_OK": bool(mom_ok[t]) if valid[t] else None,
                "Vol_OK": bool(vol_ok[t]) if valid[t] else None,
                "Entry_Gate": bool(entry[t]),
                "Crash_Flag": bool(crash[t]),
                "Days_Since_Crash": days_crash if days_crash < 10**8 else None,
                "QQQ_DD252": float(dd252[t]) if np.isfinite(dd252[t]) else np.nan,
                "QQQ_VOL": float(vol[t]) if np.isfinite(vol[t]) else np.nan,
                "QQQ_ROC5": float(f["QQQ_ROC5"][t]),
                "QQQ_ROC10": float(f["QQQ_ROC10"][t]),
                "QQQ_ROC20": float(f["QQQ_ROC20"][t]),
                "QQQ_ROC120": float(f["QQQ_ROC120"][t]),
                "QQQ_RSI14": float(f["QQQ_RSI14"][t]),
                "QQQ_MFI14": float(f["QQQ_MFI14"][t]),
                "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
                "CFG_EMA_Days": c.impulse_veto_ema_days,
                "CFG_ROC5_Min": c.impulse_veto_roc5_min,
                "CFG_ROC10_Min": c.impulse_veto_roc10_min,
                "CFG_ROC20_Min": c.impulse_veto_roc20_min,
                "CFG_RSI_Min": c.impulse_veto_rsi_min,
                "CFG_MFI_Min": c.impulse_veto_mfi_min,
                "CFG_ATR_Max": c.impulse_veto_atr_max,
                "CFG_VOL_Max": c.impulse_veto_vol_max,
                "CFG_DD_Floor": c.impulse_veto_dd_floor,
                "Model_Total_Veto_Blocks": stats_v15.get("impulse_veto_blocks", 0),
            })
    return pd.DataFrame(rows)


def build_precision_veto_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 8) -> pd.DataFrame:
    """V16 diagnostic: attribute every additional precision-veto block vs the same V15 control."""
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"])
    q = f["QQQ_CLOSE"]
    dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        if c.precision_veto_mode == "off":
            continue
        ret_v16, states_v16, stats_v16 = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        control_dict = asdict(c); control_dict["precision_veto_mode"] = "off"
        control = Config(**control_dict)
        ret_ctl, states_ctl, _ = simulate_strategy(control, f, float(cost_bps) / 10000.0)
        sma = f[f"QQQ_SMA{c.sma_days}"]; mom = f[f"QQQ_ROC{c.mom_days}"]; vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        trend_ok = q > sma; mom_ok = mom > c.mom_threshold; vol_ok = vol < c.vol_threshold
        trend_gap = np.divide(q, sma, out=np.full(len(q), np.nan), where=np.isfinite(sma) & (sma != 0)) - 1.0
        mom_gap = mom - c.mom_threshold
        score = trend_ok.astype(np.int8) + mom_ok.astype(np.int8) + vol_ok.astype(np.int8)
        entry = entry_mask_for_config(c, f)
        last_reentry = -10**9
        for t in range(len(q)-1):
            if int(states_ctl[t]) != 3 and int(states_ctl[t+1]) == 3:
                last_reentry = t
            actual_block = bool(int(states_ctl[t]) == 3 and int(states_ctl[t+1]) == 2 and int(states_v16[t]) == 3 and int(states_v16[t+1]) == 3)
            if not actual_block:
                continue
            age = t - last_reentry if last_reentry > -10**8 else None
            def comp(ret, start, days):
                hi=min(len(ret), start+days)
                if hi<=start: return np.nan
                return float(np.prod(1.0+ret[start:hi])-1.0)
            start=t+1
            rows.append({
                "Model": c.name, "Date": dates[t].date(), "Event": "V16_PRECISION_VETO_BLOCK",
                "Control_Next": STATE_NAME[int(states_ctl[t+1])], "V16_Next": STATE_NAME[int(states_v16[t+1])],
                "Mode": c.precision_veto_mode, "Base_Score": int(score[t]) if valid[t] else None,
                "Trend_OK": bool(trend_ok[t]), "Mom_OK": bool(mom_ok[t]), "Vol_OK": bool(vol_ok[t]), "Entry_Gate": bool(entry[t]),
                "Trend_Gap_vs_SMA": float(trend_gap[t]), "Mom_Gap_vs_Threshold": float(mom_gap[t]),
                "Days_Since_TQQQ_Reentry": age, "QQQ_DD252": float(dd252[t]), "QQQ_VOL": float(vol[t]),
                "QQQ_ROC5": float(f["QQQ_ROC5"][t]), "QQQ_ROC10": float(f["QQQ_ROC10"][t]), "QQQ_ROC20": float(f["QQQ_ROC20"][t]),
                "QQQ_ROC120": float(f["QQQ_ROC120"][t]), "QQQ_RSI14": float(f["QQQ_RSI14"][t]),
                "QQQ_MFI14": float(f["QQQ_MFI14"][t]), "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
                "V16_minus_Control_Next1": comp(ret_v16,start,1)-comp(ret_ctl,start,1),
                "V16_minus_Control_Next5": comp(ret_v16,start,5)-comp(ret_ctl,start,5),
                "V16_minus_Control_Next10": comp(ret_v16,start,10)-comp(ret_ctl,start,10),
                "CFG_Max_Days": c.precision_veto_max_days, "CFG_EMA_Days": c.precision_veto_ema_days,
                "CFG_ROC5_Min": c.precision_veto_roc5_min, "CFG_ROC10_Min": c.precision_veto_roc10_min, "CFG_ROC20_Min": c.precision_veto_roc20_min,
                "CFG_RSI_Min": c.precision_veto_rsi_min, "CFG_MFI_Min": c.precision_veto_mfi_min,
                "CFG_ATR_Max": c.precision_veto_atr_max, "CFG_VOL_Max": c.precision_veto_vol_max, "CFG_DD_Floor": c.precision_veto_dd_floor,
                "CFG_Trend_Gap_Floor": c.precision_veto_trend_gap_floor, "CFG_Mom_Gap_Floor": c.precision_veto_mom_gap_floor,
                "CFG_Recent_Reentry_Days": c.precision_veto_recent_reentry_days, "CFG_Score_Min": c.precision_veto_score_min,
                "Model_Total_Precision_Blocks": stats_v16.get("precision_veto_blocks",0),
            })
    return pd.DataFrame(rows)


def build_v16_exit_candidate_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 4) -> pd.DataFrame:
    """List ordinary V15-control TQQQ->QLD exits so dead/missed V16 branches are immediately diagnosable."""
    rows=[]; dates=pd.to_datetime(f["dates"]); q=f["QQQ_CLOSE"]; dd252=f["QQQ_DD252"]
    for c in configs[:max(0,top_n)]:
        ctl_dict=asdict(c); ctl_dict["precision_veto_mode"]="off"; ctl=Config(**ctl_dict)
        _, states, _=simulate_strategy(ctl,f,float(cost_bps)/10000.0)
        sma=f[f"QQQ_SMA{c.sma_days}"]; mom=f[f"QQQ_ROC{c.mom_days}"]; vol=f[f"QQQ_VOL{c.vol_days}"]
        valid=np.isfinite(sma)&np.isfinite(mom)&np.isfinite(vol)&np.isfinite(dd252)
        tr=q>sma; mo=mom>c.mom_threshold; vo=vol<c.vol_threshold; score=tr.astype(np.int8)+mo.astype(np.int8)+vo.astype(np.int8)
        tgap=np.divide(q,sma,out=np.full(len(q),np.nan),where=np.isfinite(sma)&(sma!=0))-1.0; mgap=mom-c.mom_threshold
        entry=entry_mask_for_config(c,f); last_reentry=-10**9
        for t in range(len(q)-1):
            if int(states[t])!=3 and int(states[t+1])==3: last_reentry=t
            if not (int(states[t])==3 and int(states[t+1])==2 and valid[t]): continue
            age=t-last_reentry if last_reentry>-10**8 else None
            rows.append({
                "Model":c.name,"Date":dates[t].date(),"Base_Score":int(score[t]),"Trend_OK":bool(tr[t]),"Mom_OK":bool(mo[t]),"Vol_OK":bool(vo[t]),
                "Entry_Gate":bool(entry[t]),"Days_Since_TQQQ_Reentry":age,"Trend_Gap_vs_SMA":float(tgap[t]),"Mom_Gap_vs_Threshold":float(mgap[t]),
                "QQQ_DD252":float(dd252[t]),"QQQ_VOL":float(vol[t]),"QQQ_ROC5":float(f["QQQ_ROC5"][t]),"QQQ_ROC10":float(f["QQQ_ROC10"][t]),
                "QQQ_ROC20":float(f["QQQ_ROC20"][t]),"QQQ_ROC120":float(f["QQQ_ROC120"][t]),"QQQ_RSI14":float(f["QQQ_RSI14"][t]),
                "QQQ_MFI14":float(f["QQQ_MFI14"][t]),"QQQ_ATRPCT14":float(f["QQQ_ATRPCT14"][t]),
                "Next_Open_TQQQ_minus_QLD":float(f["TQQQ_O2O_RET"][t+1]-f["QLD_O2O_RET"][t+1]) if t+1<len(q) else np.nan,
            })
    return pd.DataFrame(rows)


def build_v17_reentry_grace_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 12) -> pd.DataFrame:
    """Attribute every V18 re-entry-grace block versus the same V15/V17 control."""
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"]); q=f["QQQ_CLOSE"]; dd252=f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        if c.reentry_grace_mode == "off":
            continue
        ret_v17, states_v17, stats_v17 = simulate_strategy(c, f, float(cost_bps)/10000.0)
        ctl_dict=asdict(c); ctl_dict["reentry_grace_mode"]="off"; ctl=Config(**ctl_dict)
        ret_ctl, states_ctl, _ = simulate_strategy(ctl, f, float(cost_bps)/10000.0)
        sma=f[f"QQQ_SMA{c.sma_days}"]; mom=f[f"QQQ_ROC{c.mom_days}"]; vol=f[f"QQQ_VOL{c.vol_days}"]
        valid=np.isfinite(sma)&np.isfinite(mom)&np.isfinite(vol)&np.isfinite(dd252)
        tr=q>sma; mo=mom>c.mom_threshold; vo=vol<c.vol_threshold
        score=tr.astype(np.int8)+mo.astype(np.int8)+vo.astype(np.int8)
        tgap=np.divide(q,sma,out=np.full(len(q),np.nan),where=np.isfinite(sma)&(sma!=0))-1.0
        mgap=mom-c.mom_threshold; entry=entry_mask_for_config(c,f)
        crash=valid & (~tr) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        last_reentry=-10**9; days_crash=10**9
        for t in range(len(q)-1):
            if crash[t]: days_crash=0
            else: days_crash=min(days_crash+1,10**9)
            if int(states_ctl[t])!=3 and int(states_ctl[t+1])==3: last_reentry=t
            actual=bool(int(states_ctl[t])==3 and int(states_ctl[t+1])==2 and int(states_v17[t])==3 and int(states_v17[t+1])==3)
            if not actual: continue
            age=t-last_reentry if last_reentry>-10**8 else None
            def comp(ret,start,days):
                hi=min(len(ret),start+days)
                return np.nan if hi<=start else float(np.prod(1.0+ret[start:hi])-1.0)
            start=t+1
            branch_label = ("split_mom" if (c.reentry_grace_mode == "split_reentry" and bool(tr[t]) and (not bool(mo[t])))
                            else ("split_trend" if (c.reentry_grace_mode == "split_reentry" and (not bool(tr[t])) and bool(mo[t]))
                                  else c.reentry_grace_mode))
            rows.append({
                "Model":c.name,"Date":dates[t].date(),"Event":"V18_REENTRY_GRACE_BLOCK",
                "Control_Next":STATE_NAME[int(states_ctl[t+1])],"V18_Next":STATE_NAME[int(states_v17[t+1])],
                "Mode":branch_label,"Base_Score":int(score[t]) if valid[t] else None,
                "Trend_OK":bool(tr[t]),"Mom_OK":bool(mo[t]),"Vol_OK":bool(vo[t]),"Entry_Gate":bool(entry[t]),
                "Days_Since_TQQQ_Reentry":age,"Days_Since_Crash":days_crash if days_crash<10**8 else None,
                "Trend_Gap_vs_SMA":float(tgap[t]),"Mom_Gap_vs_Threshold":float(mgap[t]),
                "QQQ_DD252":float(dd252[t]),"QQQ_VOL":float(vol[t]),"QQQ_ROC5":float(f["QQQ_ROC5"][t]),
                "QQQ_ROC10":float(f["QQQ_ROC10"][t]),"QQQ_ROC20":float(f["QQQ_ROC20"][t]),
                "QQQ_RSI14":float(f["QQQ_RSI14"][t]),"QQQ_MFI14":float(f["QQQ_MFI14"][t]),"QQQ_ATRPCT14":float(f["QQQ_ATRPCT14"][t]),
                "V18_minus_Control_Next1":comp(ret_v17,start,1)-comp(ret_ctl,start,1),
                "V18_minus_Control_Next5":comp(ret_v17,start,5)-comp(ret_ctl,start,5),
                "V18_minus_Control_Next10":comp(ret_v17,start,10)-comp(ret_ctl,start,10),
                "CFG_Age_Min":c.reentry_grace_min_age,"CFG_Age_Max":c.reentry_grace_max_age,
                "CFG_PostCrash_Min":c.reentry_grace_min_post_crash_days,"CFG_Entry_Gate":c.reentry_grace_entry_gate,
                "CFG_EMA_Days":c.reentry_grace_ema_days,"CFG_Trend_Gap_Floor":c.reentry_grace_trend_gap_floor,
                "CFG_Mom_Gap_Floor":c.reentry_grace_mom_gap_floor,"CFG_ROC5_Min":c.reentry_grace_roc5_min,
                "CFG_ROC10_Min":c.reentry_grace_roc10_min,"CFG_ROC20_Min":c.reentry_grace_roc20_min,
                "CFG_RSI_Min":c.reentry_grace_rsi_min,"CFG_MFI_Min":c.reentry_grace_mfi_min,
                "CFG_ATR_Max":c.reentry_grace_atr_max,"CFG_VOL_Max":c.reentry_grace_vol_max,"CFG_DD_Floor":c.reentry_grace_dd_floor,
                "CFG_Mom_Entry_Gate":c.reentry_grace_mom_entry_gate,"CFG_Trend_Entry_Gate":c.reentry_grace_trend_entry_gate,
                "CFG_Mom_ROC5_Min":c.reentry_grace_mom_roc5_min,"CFG_Trend_ROC5_Min":c.reentry_grace_trend_roc5_min,
                "CFG_Max_Days":c.reentry_grace_max_days,
                "Model_Total_Grace_Blocks":stats_v17.get("reentry_grace_blocks",0),
            })
    return pd.DataFrame(rows)


def build_v17_reentry_candidate_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 4) -> pd.DataFrame:
    """Expose score-2 recent-reentry ordinary exits, including post-crash age, for V18 diagnosis."""
    rows=[]; dates=pd.to_datetime(f["dates"]); q=f["QQQ_CLOSE"]; dd252=f["QQQ_DD252"]
    for c in configs[:max(0,top_n)]:
        ctl_dict=asdict(c); ctl_dict["precision_veto_mode"]="off"; ctl_dict["reentry_grace_mode"]="off"; ctl=Config(**ctl_dict)
        _,states,_=simulate_strategy(ctl,f,float(cost_bps)/10000.0)
        sma=f[f"QQQ_SMA{c.sma_days}"]; mom=f[f"QQQ_ROC{c.mom_days}"]; vol=f[f"QQQ_VOL{c.vol_days}"]
        valid=np.isfinite(sma)&np.isfinite(mom)&np.isfinite(vol)&np.isfinite(dd252)
        tr=q>sma; mo=mom>c.mom_threshold; vo=vol<c.vol_threshold; score=tr.astype(np.int8)+mo.astype(np.int8)+vo.astype(np.int8)
        tgap=np.divide(q,sma,out=np.full(len(q),np.nan),where=np.isfinite(sma)&(sma!=0))-1.0; mgap=mom-c.mom_threshold
        entry=entry_mask_for_config(c,f); crash=valid & (~tr) & (dd252<=c.crash_dd) & (vol>=c.crash_vol_threshold)
        last_reentry=-10**9; days_crash=10**9
        for t in range(len(q)-1):
            if crash[t]: days_crash=0
            else: days_crash=min(days_crash+1,10**9)
            if int(states[t])!=3 and int(states[t+1])==3: last_reentry=t
            if not (int(states[t])==3 and int(states[t+1])==2 and valid[t] and int(score[t])==2): continue
            age=t-last_reentry if last_reentry>-10**8 else None
            if age is None or age>40: continue
            fail_vote="MOM" if bool(tr[t]) and not bool(mo[t]) and bool(vo[t]) else ("TREND" if (not bool(tr[t])) and bool(mo[t]) and bool(vo[t]) else "OTHER")
            rows.append({
                "Model":c.name,"Date":dates[t].date(),"Failed_Vote":fail_vote,"Base_Score":int(score[t]),
                "Entry_Gate":bool(entry[t]),"Days_Since_TQQQ_Reentry":age,"Days_Since_Crash":days_crash if days_crash<10**8 else None,
                "Trend_Gap_vs_SMA":float(tgap[t]),"Mom_Gap_vs_Threshold":float(mgap[t]),"QQQ_DD252":float(dd252[t]),"QQQ_VOL":float(vol[t]),
                "QQQ_ROC5":float(f["QQQ_ROC5"][t]),"QQQ_ROC10":float(f["QQQ_ROC10"][t]),"QQQ_ROC20":float(f["QQQ_ROC20"][t]),
                "QQQ_RSI14":float(f["QQQ_RSI14"][t]),"QQQ_MFI14":float(f["QQQ_MFI14"][t]),"QQQ_ATRPCT14":float(f["QQQ_ATRPCT14"][t]),
                "Next_Open_TQQQ_minus_QLD":float(f["TQQQ_O2O_RET"][t+1]-f["QLD_O2O_RET"][t+1]) if t+1<len(q) else np.nan,
            })
    return pd.DataFrame(rows)


def build_v19_micro_rescue_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 12) -> pd.DataFrame:
    """Attribute each V19 micro-rescue block versus the same V18 path with micro rescue disabled."""
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"]); q = f["QQQ_CLOSE"]; dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        if c.micro_rescue_mode == "off":
            continue
        ret_v19, states_v19, stats_v19 = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        ctl_dict = asdict(c); ctl_dict["micro_rescue_mode"] = "off"; ctl = Config(**ctl_dict)
        ret_ctl, states_ctl, _ = simulate_strategy(ctl, f, float(cost_bps) / 10000.0)
        sma = f[f"QQQ_SMA{c.sma_days}"]; mom = f[f"QQQ_ROC{c.mom_days}"]; vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        tr = q > sma; mo = mom > c.mom_threshold; vo = vol < c.vol_threshold
        score = tr.astype(np.int8) + mo.astype(np.int8) + vo.astype(np.int8)
        tgap = np.divide(q, sma, out=np.full(len(q), np.nan), where=np.isfinite(sma) & (sma != 0)) - 1.0
        mgap = mom - c.mom_threshold
        entry = entry_mask_for_config(c, f)
        crash = valid & (~tr) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        last_reentry = -10**9; days_crash = 10**9
        for t in range(len(q) - 1):
            if crash[t]: days_crash = 0
            else: days_crash = min(days_crash + 1, 10**9)
            if int(states_ctl[t]) != 3 and int(states_ctl[t + 1]) == 3:
                last_reentry = t
            actual = bool(int(states_ctl[t]) == 3 and int(states_ctl[t + 1]) == 2 and int(states_v19[t]) == 3 and int(states_v19[t + 1]) == 3)
            if not actual:
                continue
            age = t - last_reentry if last_reentry > -10**8 else None
            if int(score[t]) == 2 and bool(tr[t]) and (not bool(mo[t])) and bool(vo[t]):
                branch = "mature_mom"
            elif int(score[t]) == 1 and (not bool(tr[t])) and (not bool(mo[t])) and bool(vo[t]):
                near = bool(tgap[t] >= c.micro_rescue_neartrend_gap_floor)
                recent = age is not None and age <= c.micro_rescue_recent_max_age
                branch = "score1_neartrend" if near and not recent else "score1_pullback"
            else:
                branch = c.micro_rescue_mode
            def comp(ret, start, days):
                hi = min(len(ret), start + days)
                return np.nan if hi <= start else float(np.prod(1.0 + ret[start:hi]) - 1.0)
            start = t + 1
            rows.append({
                "Model": c.name, "Date": dates[t].date(), "Event": "V19_MICRO_RESCUE_BLOCK", "Branch": branch,
                "Control_Next": STATE_NAME[int(states_ctl[t + 1])], "V19_Next": STATE_NAME[int(states_v19[t + 1])],
                "Base_Score": int(score[t]) if valid[t] else None, "Trend_OK": bool(tr[t]), "Mom_OK": bool(mo[t]), "Vol_OK": bool(vo[t]),
                "Entry_Gate": bool(entry[t]), "Days_Since_TQQQ_Reentry": age, "Days_Since_Crash": days_crash if days_crash < 10**8 else None,
                "Trend_Gap_vs_SMA": float(tgap[t]), "Mom_Gap_vs_Threshold": float(mgap[t]),
                "QQQ_DD252": float(dd252[t]), "QQQ_VOL": float(vol[t]), "QQQ_ROC5": float(f["QQQ_ROC5"][t]),
                "QQQ_ROC10": float(f["QQQ_ROC10"][t]), "QQQ_ROC20": float(f["QQQ_ROC20"][t]),
                "QQQ_RSI14": float(f["QQQ_RSI14"][t]), "QQQ_MFI14": float(f["QQQ_MFI14"][t]), "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
                "V19_minus_Control_Next1": comp(ret_v19, start, 1) - comp(ret_ctl, start, 1),
                "V19_minus_Control_Next5": comp(ret_v19, start, 5) - comp(ret_ctl, start, 5),
                "V19_minus_Control_Next10": comp(ret_v19, start, 10) - comp(ret_ctl, start, 10),
                "CFG_Mode": c.micro_rescue_mode, "CFG_PostCrash_Min": c.micro_rescue_min_post_crash_days,
                "CFG_Recent_Age_Max": c.micro_rescue_recent_max_age, "CFG_NearTrend_Age_Max": c.micro_rescue_neartrend_max_age,
                "CFG_Mature_Age_Min": c.micro_rescue_mature_min_age, "CFG_Mature_Age_Max": c.micro_rescue_mature_max_age,
                "CFG_S1_TrendGap_Floor": c.micro_rescue_score1_trend_gap_floor, "CFG_S1_MomGap_Floor": c.micro_rescue_score1_mom_gap_floor,
                "CFG_S1_ROC5_Min": c.micro_rescue_score1_roc5_min, "CFG_S1_ROC5_Max": c.micro_rescue_score1_roc5_max,
                "CFG_NearTrend_Gap_Floor": c.micro_rescue_neartrend_gap_floor,
                "CFG_Mature_MomGap_Floor": c.micro_rescue_mature_mom_gap_floor,
                "CFG_ATR_Max": c.micro_rescue_atr_max, "CFG_VOL_Max": c.micro_rescue_vol_max, "CFG_DD_Floor": c.micro_rescue_dd_floor,
                "Model_Total_Micro_Blocks": stats_v19.get("micro_rescue_blocks", 0),
            })
    return pd.DataFrame(rows)


def build_v19_nearmiss_1y(rolling_details: pd.DataFrame, best_name: str) -> pd.DataFrame:
    """List exact 1Y losses for the best path where MDD still wins; closest return misses first."""
    if rolling_details is None or rolling_details.empty:
        return pd.DataFrame()
    d = rolling_details.copy()
    required = {"Model", "Horizon", "Return_Win", "MDD_Win", "Return_Excess"}
    if not required.issubset(d.columns):
        return pd.DataFrame()
    d = d[(d["Model"] == best_name) & (d["Horizon"] == "1Y") & (~d["Return_Win"].astype(bool)) & (d["MDD_Win"].astype(bool))].copy()
    if d.empty:
        return d
    d["Abs_Return_Miss_bp"] = (-d["Return_Excess"].astype(float) * 10000.0).clip(lower=0)
    d["Near_Miss_Band"] = np.where(d["Abs_Return_Miss_bp"] <= 25, "<=25bp", np.where(d["Abs_Return_Miss_bp"] <= 50, "<=50bp", np.where(d["Abs_Return_Miss_bp"] <= 100, "<=100bp", ">100bp")))
    secondary = "Start" if "Start" in d.columns else ("Start_Date" if "Start_Date" in d.columns else "Return_Excess")
    d = d.sort_values(["Abs_Return_Miss_bp", secondary], ascending=True)
    return d.head(200).reset_index(drop=True)


def build_v20_nearmiss_2y(rolling_details: pd.DataFrame, best_name: str) -> pd.DataFrame:
    """List exact 2Y losses for the best path where MDD still wins; closest return misses first."""
    if rolling_details is None or rolling_details.empty:
        return pd.DataFrame()
    d = rolling_details.copy()
    required = {"Model", "Horizon", "Return_Win", "MDD_Win", "Return_Excess"}
    if not required.issubset(d.columns):
        return pd.DataFrame()
    d = d[(d["Model"] == best_name) & (d["Horizon"] == "2Y") & (~d["Return_Win"].astype(bool)) & (d["MDD_Win"].astype(bool))].copy()
    if d.empty:
        return d
    d["Abs_Return_Miss_bp"] = (-d["Return_Excess"].astype(float) * 10000.0).clip(lower=0)
    d["Near_Miss_Band"] = np.where(d["Abs_Return_Miss_bp"] <= 25, "<=25bp", np.where(d["Abs_Return_Miss_bp"] <= 50, "<=50bp", np.where(d["Abs_Return_Miss_bp"] <= 100, "<=100bp", ">100bp")))
    secondary = "Start" if "Start" in d.columns else ("Start_Date" if "Start_Date" in d.columns else "Return_Excess")
    return d.sort_values(["Abs_Return_Miss_bp", secondary], ascending=True).head(250).reset_index(drop=True)


def build_v20_shadow_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 12) -> pd.DataFrame:
    """Attribute stateless V20 shadow-overlay sessions versus the identical path with shadow disabled."""
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"]); q = f["QQQ_CLOSE"]; dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        if c.shadow_rescue_mode == "off":
            continue
        ret_v20, states_v20, stats_v20 = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        ctl_dict = asdict(c); ctl_dict["shadow_rescue_mode"] = "off"; ctl = Config(**ctl_dict)
        ret_ctl, states_ctl, _ = simulate_strategy(ctl, f, float(cost_bps) / 10000.0)
        sma = f[f"QQQ_SMA{c.sma_days}"]; mom = f[f"QQQ_ROC{c.mom_days}"]; vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        tr = q > sma; mo = mom > c.mom_threshold; vo = vol < c.vol_threshold
        score = tr.astype(np.int8) + mo.astype(np.int8) + vo.astype(np.int8)
        tgap = np.divide(q, sma, out=np.full(len(q), np.nan), where=np.isfinite(sma) & (sma != 0)) - 1.0
        mgap = mom - c.mom_threshold
        entry = entry_mask_for_config(c, f)
        crash = valid & (~tr) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        last_reentry = -10**9; days_crash = 10**9
        for t in range(len(q) - 1):
            if crash[t]: days_crash = 0
            else: days_crash = min(days_crash + 1, 10**9)
            if int(states_ctl[t]) != 3 and int(states_ctl[t + 1]) == 3:
                last_reentry = t
            actual = bool(int(states_ctl[t]) == 3 and int(states_ctl[t + 1]) == 2 and int(states_v20[t + 1]) == 3)
            if not actual:
                continue
            age = t - last_reentry if last_reentry > -10**8 else None
            if int(score[t]) == 2 and bool(tr[t]) and (not bool(mo[t])) and bool(vo[t]):
                branch = "mature_mom"
            else:
                branch = "score1_recent"
            def comp(ret, start, days):
                hi = min(len(ret), start + days)
                return np.nan if hi <= start else float(np.prod(1.0 + ret[start:hi]) - 1.0)
            start = t + 1
            rows.append({
                "Model": c.name, "Date": dates[t].date(), "Event": "V20_SHADOW_OVERLAY", "Branch": branch,
                "Control_Next": STATE_NAME[int(states_ctl[t + 1])], "V20_Executed_Next": STATE_NAME[int(states_v20[t + 1])],
                "Base_Score": int(score[t]) if valid[t] else None, "Trend_OK": bool(tr[t]), "Mom_OK": bool(mo[t]), "Vol_OK": bool(vo[t]),
                "Entry_Gate": bool(entry[t]), "Days_Since_TQQQ_Reentry": age, "Days_Since_Crash": days_crash if days_crash < 10**8 else None,
                "Trend_Gap_vs_SMA": float(tgap[t]), "Mom_Gap_vs_Threshold": float(mgap[t]),
                "QQQ_DD252": float(dd252[t]), "QQQ_VOL": float(vol[t]), "QQQ_ROC5": float(f["QQQ_ROC5"][t]),
                "QQQ_ROC10": float(f["QQQ_ROC10"][t]), "QQQ_ROC20": float(f["QQQ_ROC20"][t]),
                "QQQ_RSI14": float(f["QQQ_RSI14"][t]), "QQQ_MFI14": float(f["QQQ_MFI14"][t]), "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
                "V20_minus_Control_Next1": comp(ret_v20, start, 1) - comp(ret_ctl, start, 1),
                "V20_minus_Control_Next5": comp(ret_v20, start, 5) - comp(ret_ctl, start, 5),
                "V20_minus_Control_Next10": comp(ret_v20, start, 10) - comp(ret_ctl, start, 10),
                "CFG_Shadow_Mode": c.shadow_rescue_mode, "CFG_PostCrash_Min": c.shadow_min_post_crash_days,
                "CFG_Recent_Max_Age": c.shadow_recent_max_age, "CFG_Mature_Age": f"{c.shadow_mature_min_age}-{c.shadow_mature_max_age}",
                "Model_Total_Shadow_Blocks": stats_v20.get("shadow_rescue_blocks", 0),
            })
    return pd.DataFrame(rows)


def build_v21_preview_audit(configs: List[Config], f: Dict[str, Any], cost_bps: float, top_n: int = 12) -> pd.DataFrame:
    """Attribute V21 execution-only forward previews versus identical V20 path."""
    rows: List[Dict[str, Any]] = []
    dates = pd.to_datetime(f["dates"]); q = f["QQQ_CLOSE"]; dd252 = f["QQQ_DD252"]
    for c in configs[:max(0, top_n)]:
        if c.preview_overlay_mode == "off":
            continue
        ret_v21, states_v21, stats_v21 = simulate_strategy(c, f, float(cost_bps) / 10000.0)
        ctl_dict = asdict(c); ctl_dict["preview_overlay_mode"] = "off"; ctl = Config(**ctl_dict)
        ret_ctl, states_ctl, _ = simulate_strategy(ctl, f, float(cost_bps) / 10000.0)
        sma = f[f"QQQ_SMA{c.sma_days}"]; mom = f[f"QQQ_ROC{c.mom_days}"]; vol = f[f"QQQ_VOL{c.vol_days}"]
        valid = np.isfinite(sma) & np.isfinite(mom) & np.isfinite(vol) & np.isfinite(dd252)
        tr = q > sma; mo = mom > c.mom_threshold; vo = vol < c.vol_threshold
        score = tr.astype(np.int8) + mo.astype(np.int8) + vo.astype(np.int8)
        tgap = np.divide(q, sma, out=np.full(len(q), np.nan), where=np.isfinite(sma) & (sma != 0)) - 1.0
        mgap = mom - c.mom_threshold
        entry = entry_mask_for_config(c, f)
        crash = valid & (~tr) & (dd252 <= c.crash_dd) & (vol >= c.crash_vol_threshold)
        days_crash = 10**9
        for t in range(len(q) - 1):
            if crash[t]: days_crash = 0
            else: days_crash = min(days_crash + 1, 10**9)
            actual = bool(int(states_ctl[t + 1]) < 3 and int(states_v21[t + 1]) == 3)
            if not actual:
                continue
            s_now = int(score[t]) if valid[t] else -1
            if int(states_ctl[t]) < 2 and int(states_ctl[t + 1]) == 2:
                branch = "stepup"
            elif s_now == 3:
                branch = "qld_score3"
            else:
                branch = "qld_score2"
            def comp(ret, start, days):
                hi = min(len(ret), start + days)
                return np.nan if hi <= start else float(np.prod(1.0 + ret[start:hi]) - 1.0)
            start = t + 1
            rows.append({
                "Model": c.name, "Date": dates[t].date(), "Event": "V21_PREVIEW_OVERLAY", "Branch": branch,
                "Control_Executed_Next": STATE_NAME[int(states_ctl[t + 1])], "V21_Executed_Next": STATE_NAME[int(states_v21[t + 1])],
                "Base_Score": s_now if s_now >= 0 else None, "Trend_OK": bool(tr[t]), "Mom_OK": bool(mo[t]), "Vol_OK": bool(vo[t]),
                "Entry_Gate": bool(entry[t]), "Days_Since_Crash": days_crash if days_crash < 10**8 else None,
                "Trend_Gap_vs_SMA": float(tgap[t]), "Mom_Gap_vs_Threshold": float(mgap[t]),
                "QQQ_DD252": float(dd252[t]), "QQQ_VOL": float(vol[t]), "QQQ_ROC5": float(f["QQQ_ROC5"][t]),
                "QQQ_ROC10": float(f["QQQ_ROC10"][t]), "QQQ_ROC20": float(f["QQQ_ROC20"][t]),
                "QQQ_RSI14": float(f["QQQ_RSI14"][t]), "QQQ_MFI14": float(f["QQQ_MFI14"][t]), "QQQ_ATRPCT14": float(f["QQQ_ATRPCT14"][t]),
                "V21_minus_Control_Next1": comp(ret_v21, start, 1) - comp(ret_ctl, start, 1),
                "V21_minus_Control_Next5": comp(ret_v21, start, 5) - comp(ret_ctl, start, 5),
                "V21_minus_Control_Next10": comp(ret_v21, start, 10) - comp(ret_ctl, start, 10),
                "CFG_Preview_Mode": c.preview_overlay_mode, "CFG_PostCrash_Min": c.preview_min_post_crash_days,
                "CFG_QLD_Max_Age": c.preview_qld_max_delever_age,
                "Model_Total_Preview_Blocks": stats_v21.get("preview_overlay_blocks", 0),
            })
    return pd.DataFrame(rows)


def build_v21_nearmiss_2y(rolling_details: pd.DataFrame, best_name: str) -> pd.DataFrame:
    """Exact 2Y losses for the best V21 path where MDD still wins, closest misses first."""
    return build_v20_nearmiss_2y(rolling_details, best_name)


# -----------------------------
# Excel dashboard
# -----------------------------

DARK = "172033"
DARK2 = "233047"
WHITE = "FFFFFF"
GREEN = "00B050"
RED = "C00000"
ORANGE = "F4B183"
TEAL = "00A6A6"
BLUE = "4472C4"
PURPLE = "7030A0"
GRAY = "808080"
LIGHT_GREEN = "E2F0D9"
LIGHT_RED = "FCE4D6"
LIGHT_BLUE = "D9EAF7"
LIGHT_PURPLE = "E4DFEC"


def _safe(v: Any) -> Any:
    if isinstance(v, (np.floating, np.integer)): return v.item()
    if isinstance(v, (np.bool_,)): return bool(v)
    return v


def _write_table(ws, df: pd.DataFrame, start_row: int = 1, start_col: int = 1, freeze: Optional[str] = None) -> None:
    header_fill = PatternFill("solid", fgColor=DARK2)
    header_font = Font(color=WHITE, bold=True)
    imported_font = Font(color="008000")  # linked/imported analytics
    for j, col in enumerate(df.columns, start_col):
        cell = ws.cell(start_row, j, str(col))
        cell.fill = header_fill; cell.font = header_font; cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, row in enumerate(df.itertuples(index=False, name=None), start_row + 1):
        for j, v in enumerate(row, start_col):
            v = _safe(v)
            if isinstance(v, dict) or isinstance(v, list): v = json.dumps(v, ensure_ascii=False)
            ws.cell(i, j, v)
            ws.cell(i, j).font = imported_font
    if freeze: ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    # Auto width with sensible limits.
    for j, col in enumerate(df.columns, start_col):
        vals = [str(col)] + [str(_safe(x)) for x in df[col].head(250).tolist()]
        width = min(48, max(10, max(len(x) for x in vals) + 2))
        ws.column_dimensions[get_column_letter(j)].width = width


def format_percent_columns(ws, df: pd.DataFrame, start_row: int = 1, start_col: int = 1) -> None:
    pct_keywords = ["rate", "return", "mdd", "cagr", "sharpe", "calmar", "cash", "qqq", "qld", "tqqq", "improvement", "excess"]
    for j, col in enumerate(df.columns, start_col):
        lc = str(col).lower()
        if any(k in lc for k in pct_keywords) and not any(k in lc for k in ["score", "days", "count", "windows", "sharpe", "calmar"]):
            for i in range(start_row + 1, start_row + 1 + len(df)):
                ws.cell(i, j).number_format = '0.0%;[Red](0.0%);-'
        elif "score" in lc or "sharpe" in lc or "calmar" in lc:
            for i in range(start_row + 1, start_row + 1 + len(df)):
                ws.cell(i, j).number_format = '0.00'


EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_DATA_ROWS = EXCEL_MAX_ROWS - 1  # row 1 is reserved for the header


def _split_sheet_name(base_name: str, part: int) -> str:
    """Return an Excel-safe continuation sheet name (max 31 chars)."""
    if part <= 1:
        return base_name[:31]
    suffix = f"_{part:02d}"
    # Keep the familiar first sheet name; continuation sheets are shortened if needed.
    stem = base_name.replace("_WINDOWS", "_WIN")
    return f"{stem[:31-len(suffix)]}{suffix}"


def _write_split_table(
    wb,
    base_name: str,
    df: pd.DataFrame,
    *,
    freeze: Optional[str] = "A2",
    apply_percent_format: bool = True,
    max_data_rows: int = EXCEL_MAX_DATA_ROWS,
) -> List[str]:
    """Write a DataFrame across multiple worksheets without exceeding Excel's row limit.

    Excel supports at most 1,048,576 rows per worksheet. With a header in row 1,
    each sheet can therefore contain at most 1,048,575 data rows.
    """
    if max_data_rows < 1 or max_data_rows > EXCEL_MAX_DATA_ROWS:
        raise ValueError(f"max_data_rows must be between 1 and {EXCEL_MAX_DATA_ROWS:,}")

    if df is None or len(df) == 0:
        ws = wb.create_sheet(_split_sheet_name(base_name, 1))
        ws["A1"] = "No detailed rows (set --detail-top > 0)."
        ws.sheet_view.showGridLines = False
        return [ws.title]

    sheet_names: List[str] = []
    total = len(df)
    parts = (total + max_data_rows - 1) // max_data_rows
    for part in range(1, parts + 1):
        lo = (part - 1) * max_data_rows
        hi = min(part * max_data_rows, total)
        chunk = df.iloc[lo:hi]
        ws = wb.create_sheet(_split_sheet_name(base_name, part))
        _write_table(ws, chunk, freeze=freeze)
        if apply_percent_format:
            format_percent_columns(ws, chunk)
        sheet_names.append(ws.title)
        print(f"[excel] {base_name}: sheet {part}/{parts} rows {lo+1:,}-{hi:,} -> {ws.title}")
    return sheet_names


def build_excel(
    path: Path,
    final_rows: List[Dict[str, Any]],
    rolling_summary: pd.DataFrame,
    rolling_details: pd.DataFrame,
    search_top: pd.DataFrame,
    indicator_df: pd.DataFrame,
    cohort_df: pd.DataFrame,
    audit_df: pd.DataFrame,
    transition_audit_df: pd.DataFrame,
    tqqq_reentry_audit_df: pd.DataFrame,
    impulse_veto_audit_df: pd.DataFrame,
    precision_veto_audit_df: pd.DataFrame,
    v16_exit_candidate_df: pd.DataFrame,
    v17_reentry_grace_audit_df: pd.DataFrame,
    v17_reentry_candidate_df: pd.DataFrame,
    v19_micro_rescue_audit_df: pd.DataFrame,
    v19_nearmiss_1y_df: pd.DataFrame,
    v20_shadow_audit_df: pd.DataFrame,
    v20_nearmiss_2y_df: pd.DataFrame,
    v21_preview_audit_df: pd.DataFrame,
    v21_nearmiss_2y_df: pd.DataFrame,
    run_info: Dict[str, Any],
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "00_DASHBOARD"
    ws.sheet_view.showGridLines = False

    title_fill = PatternFill("solid", fgColor=DARK)
    sub_fill = PatternFill("solid", fgColor=DARK2)
    good_fill = PatternFill("solid", fgColor=LIGHT_GREEN)
    bad_fill = PatternFill("solid", fgColor=LIGHT_RED)
    teal_fill = PatternFill("solid", fgColor="DDEBF7")

    ws.merge_cells("B2:N3")
    ws["B2"] = "ROBUST TIGHT ROLLING LAB V21"
    ws["B2"].font = Font(size=22, bold=True, color=WHITE)
    ws["B2"].fill = title_fill
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    for row in ws["B2:R3"]:
        for cell in row: cell.fill = title_fill

    ws.merge_cells("B4:R4")
    ws["B4"] = "Goal: preserve the V20 80.8629% + 98.1054% frontier; add one-session forward Preview Overlay to push the 2Y bottleneck through 81%+97% without changing the control-state path"
    ws["B4"].font = Font(color=WHITE, italic=True)
    ws["B4"].fill = sub_fill
    ws["B4"].alignment = Alignment(horizontal="center")
    for row in ws["B4:R4"]:
        for cell in row: cell.fill = sub_fill

    finals = sorted(final_rows, key=lambda x: x["final_score"], reverse=True)
    best = finals[0]
    original = next((r for r in finals if r["name"] == "ROBUST_ORIGINAL"), None)

    kpis = [
        ("Best Model", best["name"], None),
        ("STRICT ALL-WINDOW PASS", "YES" if best["strict_all_windows_both_win"] else "NO", "good" if best["strict_all_windows_both_win"] else "bad"),
        ("Worst-Horizon BOTH Win", best["min_both_win_rate"], None),
        ("Worst-Horizon Return Win", best["min_return_win_rate"], None),
        ("Worst-Horizon MDD Win", best["min_mdd_win_rate"], None),
        ("Full CAGR", best["cagr"], None),
        ("Full MDD", best["mdd"], None),
        ("TQQQ Full CAGR", best["bench_cagr"], None),
        ("TQQQ Full MDD", best["bench_mdd"], None),
        ("CAGR Excess", best["cagr_excess"], None),
        ("MDD Improvement", best["mdd_improvement"], None),
        ("Bottom-3 Short Cohort", best["bottom3_startyear_short_win_rate"], None),
        ("Bottom-20% Short Cohort", best["bottom20_startyear_short_win_rate"], None),
        ("Joint 79% + 97%", "YES" if best.get("target_joint_79_97") else "NO", "good" if best.get("target_joint_79_97") else "bad"),
        ("Joint 79.5% + 97%", "YES" if best.get("target_joint_795_97") else "NO", "good" if best.get("target_joint_795_97") else None),
        ("Joint 80% + 97%", "YES" if best.get("target_joint_80_97") else "NO", "good" if best.get("target_joint_80_97") else None),
        ("Joint 81% + 97%", "YES" if best.get("target_joint_81_97") else "NO", "good" if best.get("target_joint_81_97") else None),
        ("Joint 82% + 97%", "YES" if best.get("target_joint_82_97") else "NO", "good" if best.get("target_joint_82_97") else None),
        ("2Y Wins Needed for 81%", best.get("two_year_wins_needed_81"), "good" if best.get("two_year_wins_needed_81", 1) == 0 else None),
        ("2Y Near Miss <=25bp", best.get("two_year_near_miss_25bp"), None),
        ("Worst-H Wins Needed", best.get("worst_horizon_wins_needed_80"), "good" if best.get("worst_horizon_wins_needed_80", 1) == 0 else None),
        ("Perfect Horizons", f"{best['perfect_horizons']} / {best['available_horizons']}", None),
    ]
    r0 = 6
    for i, (label, value, flag) in enumerate(kpis):
        r = r0 + i // 4 * 3
        c = 2 + (i % 4) * 3
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
        ws.cell(r, c, label); ws.cell(r, c).font = Font(bold=True, color=WHITE); ws.cell(r, c).fill = sub_fill
        ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
        ws.cell(r + 1, c, value); ws.cell(r + 1, c).font = Font(size=13, bold=True, color="000000")
        ws.cell(r + 1, c).fill = good_fill if flag == "good" else bad_fill if flag == "bad" else teal_fill
        if isinstance(value, float): ws.cell(r + 1, c).number_format = '0.0%;[Red](0.0%);-'

    # Top-model table on dashboard.
    top_cols = [
        "name", "strict_all_windows_both_win", "perfect_horizons", "final_score",
        "min_both_win_rate", "min_return_win_rate", "min_mdd_win_rate",
        "short_both_win_rate", "avg_both_win_rate", "bottom3_startyear_short_win_rate",
        "bottom20_startyear_short_win_rate", "target_both_79", "target_mdd_97", "target_joint_79_97", "target_joint_795_97", "target_joint_80_97", "target_joint_81_97", "target_joint_82_97", "joint_gap_79_97",
        "one_year_both_wins", "one_year_wins_needed_80", "two_year_both_wins", "two_year_wins_needed_80", "two_year_wins_needed_81", "two_year_wins_needed_82", "two_year_near_miss_25bp", "target_bottom20_21", "core_targets_passed", "targets_passed",
        "impulse_veto_blocks", "impulse_veto_score1_blocks", "impulse_veto_score2_blocks",
        "precision_veto_blocks", "precision_veto_trend_blocks", "precision_veto_mom_blocks", "precision_veto_score1_blocks", "precision_veto_fresh_blocks",
        "reentry_grace_blocks", "reentry_grace_mom_blocks", "reentry_grace_trend_blocks",
        "micro_rescue_blocks", "micro_rescue_score1_pullback_blocks", "micro_rescue_score1_neartrend_blocks", "micro_rescue_mature_mom_blocks",
        "shadow_rescue_blocks", "shadow_rescue_score1_blocks", "shadow_rescue_mature_mom_blocks",
        "preview_overlay_blocks", "preview_overlay_qld3_blocks", "preview_overlay_qld2_blocks", "preview_overlay_stepup_blocks",
        "dead_v21_modules", "dead_v20_modules", "dead_v19_modules",
        "worst_mdd_improvement", "cagr", "mdd", "cagr_excess", "mdd_improvement", "gates",
    ]
    top_df = pd.DataFrame([{k: r.get(k) for k in top_cols} for r in finals[:15]])
    start = r0 + int(math.ceil(len(kpis) / 4.0)) * 3 + 2
    _write_table(ws, top_df, start_row=start, start_col=2)
    format_percent_columns(ws, top_df, start_row=start, start_col=2)
    ws.freeze_panes = f"B{start}"
    ws.conditional_formatting.add(
        f"F{start+1}:H{start+len(top_df)}",
        ColorScaleRule(start_type="num", start_value=0, start_color="F8696B", mid_type="num", mid_value=0.8, mid_color="FFEB84", end_type="num", end_value=1.0, end_color="63BE7B")
    )

    # Rolling both-win chart for best model.
    bsum = rolling_summary[rolling_summary["Model"] == best["name"]].copy()
    if not bsum.empty:
        chart_start = start + len(top_df) + 3
        ws.cell(chart_start, 2, "Best Model Rolling Win Rates").font = Font(size=14, bold=True)
        for j, h in enumerate(["Horizon", "Return_Win_Rate", "MDD_Win_Rate", "Both_Win_Rate"], 2):
            ws.cell(chart_start + 1, j, h).font = Font(bold=True, color=WHITE); ws.cell(chart_start + 1, j).fill = sub_fill
        for i, row in enumerate(bsum[["Horizon", "Return_Win_Rate", "MDD_Win_Rate", "Both_Win_Rate"]].itertuples(index=False, name=None), chart_start + 2):
            for j, v in enumerate(row, 2): ws.cell(i, j, _safe(v))
            for j in [3,4,5]: ws.cell(i, j).number_format = '0.0%'
        chart = BarChart()
        chart.type = "col"; chart.style = 10
        chart.title = f"{best['name']} rolling win rates"
        chart.y_axis.title = "Win rate"; chart.x_axis.title = "Horizon"
        data = Reference(ws, min_col=3, max_col=5, min_row=chart_start + 1, max_row=chart_start + 1 + len(bsum))
        cats = Reference(ws, min_col=2, min_row=chart_start + 2, max_row=chart_start + 1 + len(bsum))
        chart.add_data(data, titles_from_data=True); chart.set_categories(cats)
        chart.height = 8; chart.width = 15
        ws.add_chart(chart, f"H{chart_start+1}")

    for col in range(2, 19): ws.column_dimensions[get_column_letter(col)].width = 16
    ws.column_dimensions["R"].width = 48

    # 01_TOP_MODELS
    top_full_rows = []
    for rank, r in enumerate(finals, 1):
        d = {k: v for k, v in r.items() if k != "config"}
        d["Rank"] = rank
        cfg = r["config"]
        for k, v in cfg.items(): d[f"CFG_{k}"] = v
        top_full_rows.append(d)
    top_full = pd.DataFrame(top_full_rows)
    ws1 = wb.create_sheet("01_TOP_MODELS")
    _write_table(ws1, top_full, freeze="A2"); format_percent_columns(ws1, top_full)

    # 02_ROLLING_SUMMARY
    ws2 = wb.create_sheet("02_ROLLING_SUMMARY")
    _write_table(ws2, rolling_summary, freeze="A2"); format_percent_columns(ws2, rolling_summary)
    # Highlight perfect rows and low both-win rows.
    if len(rolling_summary):
        headers = {c.value: c.column for c in ws2[1]}
        pcol = get_column_letter(headers["Perfect_All_Windows"])
        bcol = get_column_letter(headers["Both_Win_Rate"])
        ws2.conditional_formatting.add(f"{pcol}2:{pcol}{len(rolling_summary)+1}", FormulaRule(formula=[f'{pcol}2=TRUE'], fill=good_fill))
        ws2.conditional_formatting.add(f"{bcol}2:{bcol}{len(rolling_summary)+1}", ColorScaleRule(start_type="num", start_value=0, start_color="F8696B", mid_type="num", mid_value=0.8, mid_color="FFEB84", end_type="num", end_value=1, end_color="63BE7B"))

    # 03_ROLLING_WINDOWS
    # V13.1: detailed rolling rows can exceed Excel's hard 1,048,576-row worksheet limit.
    # Preserve ALL rows by splitting automatically across continuation sheets.
    _write_split_table(wb, "03_ROLLING_WINDOWS", rolling_details, freeze="A2", apply_percent_format=True)

    # 04_SEARCH_TOP
    ws4 = wb.create_sheet("04_SEARCH_TOP")
    _write_table(ws4, search_top, freeze="A2"); format_percent_columns(ws4, search_top)

    # 05_INDICATOR_STATS
    ws5 = wb.create_sheet("05_INDICATOR_STATS")
    _write_table(ws5, indicator_df, freeze="A2"); format_percent_columns(ws5, indicator_df)

    # 06_RUN_INFO
    ws6 = wb.create_sheet("06_RUN_INFO")
    info_df = pd.DataFrame([{"Key": k, "Value": json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v} for k, v in run_info.items()])
    _write_table(ws6, info_df, freeze="A2")

    # 07_STRATEGY_GUIDE
    ws7 = wb.create_sheet("07_STRATEGY_GUIDE")
    guide = [
        ["Item", "Meaning"],
        ["STRICT ALL-WINDOW PASS", "Every available daily-start rolling window from 6M to 10Y has BOTH return >= TQQQ and MDD >= TQQQ (less negative is better)."],
        ["Return Win Rate", "Fraction of rolling windows where strategy return >= TQQQ Buy & Hold (6M/1Y period return; 2Y+ CAGR)."],
        ["MDD Win Rate", "Fraction of rolling windows where strategy MDD is less severe than TQQQ Buy & Hold."],
        ["Both Win Rate", "Fraction where BOTH return and MDD win simultaneously."],
        ["Worst-Horizon BOTH Win", "Minimum Both Win Rate among all available horizons. This is a major ranking criterion."],
        ["Worst Return Excess", "Worst strategy return minus TQQQ return among all rolling windows (6M/1Y period return; 2Y+ CAGR). Positive is ideal."],
        ["Worst MDD Improvement", "Worst strategy MDD minus TQQQ MDD among all rolling windows. Positive is ideal."],
        ["Fast direct", "Original-style Fast Re-entry can jump directly to TQQQ if gates pass."],
        ["Fast qld", "Fast Re-entry is capped at QLD; TQQQ requires subsequent normal confirmation."],
        ["Fast normal", "Fast override disabled; only normal step-up rules are used."],
        ["V8 Staged Re-entry", "After a real TQQQ deleveraging, recovery can enter QLD first and promote to TQQQ only after a stronger second confirmation. Recovery ATR is separate from the normal ATR8 gate."],
        ["V8 Explosive Rebound", "Only after a genuine crash, extreme ROC5/ROC10 + RSI/MFI can restore TQQQ directly. A short latch prevents immediate re-cashing while the long crash flag clears."],
        ["V9 Correction Re-entry", "Armed only after a non-crash TQQQ deleveraging. Uses short momentum/RSI/MFI/ATR/vol/DD to re-enter without weakening deep-crash defense."],
        ["V9 DD-Rebound Release", "During a genuine crash, remembers the worst drawdown and can restore exposure after a generic rebound from that trough; no dates are hardcoded."],
        ["V12 Findings", "Split QLD->TQQQ crash recovery is retained; ordinary re-promotion guard is rejected because fewer whipsaws did not improve weak cohorts."],
        ["V11 Trough-Confirmed DDR2", "Historical control. A real crash arms a separate trough and every fresh low resets confirmation."],
        ["V14 Finding / V15 Anchor", "V14 real-data DEEP produced zero soft-vol signal days across all selected V14 candidates; the exact V12 stable split therefore remains the control."],
        ["V15 Selective Impulse Veto", "Only an ordinary TQQQ->QLD step-down can be delayed briefly when short ROC/RSI/MFI/EMA evidence is strongly bullish while volatility itself remains acceptable."],
        ["V15 Safety", "Crash exits are never blocked. Post-crash context is exempt, later TQQQ-specific exit/defense rules can still override the veto, and max hold is only 1-2 sessions."],
        ["V17 Achievement / V18 Control", "V17 FOCUS_471542 improved 1Y BOTH wins to 2,904/3,659, Worst BOTH 79.3659%, MDD Win 98.1054%, CAGR 51.4369%, with Full MDD unchanged at -47.5448%. V18 preserves it as the primary control."],
        ["V16 Precision Transition Veto", "Layered after V15. Separately tests trend-only near-fail, momentum-only near-fail, score-1 near-boundary rebound, and fresh-TQQQ-reentry transient weakness."],
        ["V16 Safety", "Only ordinary TQQQ->QLD exits can be blocked. Crash/post-crash context remains exempt; consecutive V15+V16 veto stacking is forbidden; broad hysteresis is penalized."],
        ["V16 Frontier Targets", "Keep joint 79%+97% as a floor; additionally report joint 79.5%+97% and 80%+97%, with exact win counts needed."],
        ["V18 Finding / Direction", "V17 showed two safe Score-2 sub-regimes with different ROC5 behavior. V18 splits Momentum-fail and Trend-fail thresholds, then forces high-1Y-win rough candidates into exact MDD validation."],
        ["V18 Split Re-entry Grace", "Momentum-fail can use a more permissive ROC5 floor with Entry Gate PASS, while Trend-fail uses a stricter ROC5 floor and may ignore Entry Gate. Optional day-2 continuation is allowed only for split-mode Trend-fail."],
        ["V18 Safety", "Active crash exits remain absolute, Score-1 remains excluded, V15/V17 adjacent stacking remains blocked except one explicitly bounded Trend-fail continuation day, and broad grace counts are penalized after six blocks."],
        ["V18 Near-miss Diagnostics", "Exact final stage still reports 1Y win count and 25/50/100bp near misses; V18 finalist selection includes a dedicated 1Y force-lane."],
        ["V19 Micro Rescue", "Preserves V18 split re-entry and tests only sparse residual exits: recent Score-1 pullbacks, Score-1 near the trend boundary, and mature Momentum-only micro-fails."],
        ["V19 Safety", "Micro Rescue can block only an ordinary TQQQ->QLD step-down for one session. Active crash/recovery branches remain absolute; excessive or dead rescue modules are penalized."],
        ["V19 Near-miss Audit", "Lists exact 1Y return losses where MDD still wins, sorted by basis-point distance. It is diagnostic only and never introduces date-specific rules."],
        ["V20 Shadow Overlay", "When an ordinary TQQQ->QLD step-down qualifies, the underlying state still moves to QLD but next-session execution alone is temporarily held at TQQQ. This prevents the rescue from cascading into a different multi-day state path."],
        ["V20 Target", "Preserve the V19 80%+97% frontier and attack the new 2Y bottleneck. Exact finals report 81%+97% and 82%+97%, 2Y win counts, and 25/50/100bp 2Y near misses."],
        ["V20 Safety", "Shadow Overlay is execution-only, one session, never overrides active crash/recovery protection, and excessive/dead shadow modules are penalized."],
        ["V21 Forward Preview", "Adds the mirror image of V20: while the control remains/steps only to QLD, a narrowly confirmed session may execute TQQQ for one session; the next decision still starts from the unchanged control state."],
        ["V21 Target", "Preserve V20 80.8629% BOTH / 98.1054% MDD Win and close the remaining five 2Y wins needed for 81%+97%; keep 82%+97% as stretch target."],
        ["V21 Safety", "Preview is stateless, one session, blocked during crash/special recovery modules, and candidates with dead or excessive preview events are penalized."],
        ["Bottom-3 Short Cohort", "Average BOTH-win rate of the three weakest qualifying start-year cohorts across 6M/1Y/2Y/3Y. No year is hardcoded."],
        ["Bottom-20% Short Cohort", "Average BOTH-win rate of the weakest 20% of qualifying start-year cohorts across 6M/1Y/2Y/3Y."],
        ["Risk Defense", "Dual-Anchor-v15-inspired composite risk score can cap exposure immediately when threshold is reached."],
        ["Execution", "Close signal -> next session open. Returns are open-to-open, with estimated transaction costs."],
        ["Important", "A perfect all-window winner may not exist. If none exists, rank by worst-horizon consistency rather than cherry-picking the best CAGR."],
    ]
    for i, row in enumerate(guide, 1):
        for j, v in enumerate(row, 1): ws7.cell(i, j, v)
    for c in ws7[1]: c.fill = sub_fill; c.font = Font(color=WHITE, bold=True)
    ws7.column_dimensions["A"].width = 28; ws7.column_dimensions["B"].width = 120
    ws7.sheet_view.showGridLines = False

    # 08_COHORT_FLOORS
    ws8 = wb.create_sheet("08_COHORT_FLOORS")
    if cohort_df is not None and len(cohort_df):
        _write_table(ws8, cohort_df, freeze="A2"); format_percent_columns(ws8, cohort_df)
    else:
        ws8["A1"] = "No cohort detail rows (increase --detail-top)."

    # 09_REENTRY_AUDIT
    ws9 = wb.create_sheet("09_REENTRY_AUDIT")
    if audit_df is not None and len(audit_df):
        _write_table(ws9, audit_df, freeze="A2"); format_percent_columns(ws9, audit_df)
    else:
        ws9["A1"] = "No re-entry audit rows."

    # 10_TRANSITION_AUDIT
    ws10 = wb.create_sheet("10_TRANSITION_AUDIT")
    if transition_audit_df is not None and len(transition_audit_df):
        _write_table(ws10, transition_audit_df, freeze="A2"); format_percent_columns(ws10, transition_audit_df)
    else:
        ws10["A1"] = "No transition audit rows."

    # 11_TQQQ_REENTRY_AUDIT
    ws11 = wb.create_sheet("11_TQQQ_REENTRY_AUDIT")
    if tqqq_reentry_audit_df is not None and len(tqqq_reentry_audit_df):
        _write_table(ws11, tqqq_reentry_audit_df, freeze="A2"); format_percent_columns(ws11, tqqq_reentry_audit_df)
    else:
        ws11["A1"] = "No TQQQ re-entry audit rows."

    # 12_V15_IMPULSE_AUDIT
    ws12 = wb.create_sheet("12_V15_IMPULSE_AUDIT")
    if impulse_veto_audit_df is not None and len(impulse_veto_audit_df):
        _write_table(ws12, impulse_veto_audit_df, freeze="A2"); format_percent_columns(ws12, impulse_veto_audit_df)
    else:
        ws12["A1"] = "No V15 impulse-veto blocks among audited models."

    # 13_V16_PRECISION_AUDIT
    ws13 = wb.create_sheet("13_V16_PRECISION_AUDIT")
    if precision_veto_audit_df is not None and len(precision_veto_audit_df):
        _write_table(ws13, precision_veto_audit_df, freeze="A2"); format_percent_columns(ws13, precision_veto_audit_df)
    else:
        ws13["A1"] = "No V16 precision-veto blocks among audited models."

    # 14_V16_EXIT_CANDIDATES
    ws14 = wb.create_sheet("14_V16_EXIT_CANDIDATES")
    if v16_exit_candidate_df is not None and len(v16_exit_candidate_df):
        _write_table(ws14, v16_exit_candidate_df, freeze="A2"); format_percent_columns(ws14, v16_exit_candidate_df)
    else:
        ws14["A1"] = "No ordinary V15-control TQQQ->QLD exits found."

    # 15_V18_REENTRY_AUDIT
    ws15 = wb.create_sheet("15_V18_REENTRY_AUDIT")
    if v17_reentry_grace_audit_df is not None and len(v17_reentry_grace_audit_df):
        _write_table(ws15, v17_reentry_grace_audit_df, freeze="A2"); format_percent_columns(ws15, v17_reentry_grace_audit_df)
    else:
        ws15["A1"] = "No V18 re-entry-grace blocks among audited models."

    # 16_V18_REENTRY_CANDIDATES
    ws16 = wb.create_sheet("16_V18_REENTRY_CANDIDATES")
    if v17_reentry_candidate_df is not None and len(v17_reentry_candidate_df):
        _write_table(ws16, v17_reentry_candidate_df, freeze="A2"); format_percent_columns(ws16, v17_reentry_candidate_df)
    else:
        ws16["A1"] = "No recent score-2 ordinary exits found."

    # 17_V19_MICRO_AUDIT
    ws17 = wb.create_sheet("17_V19_MICRO_AUDIT")
    if v19_micro_rescue_audit_df is not None and len(v19_micro_rescue_audit_df):
        _write_table(ws17, v19_micro_rescue_audit_df, freeze="A2"); format_percent_columns(ws17, v19_micro_rescue_audit_df)
    else:
        ws17["A1"] = "No V19 micro-rescue blocks among audited models."

    # 18_V19_NEARMISS_1Y
    ws18 = wb.create_sheet("18_V19_NEARMISS_1Y")
    if v19_nearmiss_1y_df is not None and len(v19_nearmiss_1y_df):
        _write_table(ws18, v19_nearmiss_1y_df, freeze="A2"); format_percent_columns(ws18, v19_nearmiss_1y_df)
    else:
        ws18["A1"] = "No best-model 1Y return-loss / MDD-win detail rows available."

    # 19_V20_SHADOW_AUDIT
    ws19 = wb.create_sheet("19_V20_SHADOW_AUDIT")
    if v20_shadow_audit_df is not None and len(v20_shadow_audit_df):
        _write_table(ws19, v20_shadow_audit_df, freeze="A2"); format_percent_columns(ws19, v20_shadow_audit_df)
    else:
        ws19["A1"] = "No V20 shadow-overlay executions among audited models."

    # 20_V20_NEARMISS_2Y
    ws20 = wb.create_sheet("20_V20_NEARMISS_2Y")
    if v20_nearmiss_2y_df is not None and len(v20_nearmiss_2y_df):
        _write_table(ws20, v20_nearmiss_2y_df, freeze="A2"); format_percent_columns(ws20, v20_nearmiss_2y_df)
    else:
        ws20["A1"] = "No best-model 2Y return-loss / MDD-win detail rows available."

    # 21_V21_PREVIEW_AUDIT
    ws21 = wb.create_sheet("21_V21_PREVIEW_AUDIT")
    if v21_preview_audit_df is not None and len(v21_preview_audit_df):
        _write_table(ws21, v21_preview_audit_df, freeze="A2"); format_percent_columns(ws21, v21_preview_audit_df)
    else:
        ws21["A1"] = "No V21 preview-overlay executions among audited models."

    # 22_V21_NEARMISS_2Y
    ws22 = wb.create_sheet("22_V21_NEARMISS_2Y")
    if v21_nearmiss_2y_df is not None and len(v21_nearmiss_2y_df):
        _write_table(ws22, v21_nearmiss_2y_df, freeze="A2"); format_percent_columns(ws22, v21_nearmiss_2y_df)
    else:
        ws22["A1"] = "No best-model V21 2Y return-loss / MDD-win detail rows available."

    # Workbook-wide styling details.
    for sh in wb.worksheets:
        sh.sheet_view.showGridLines = False
        sh.auto_filter.ref = sh.dimensions if sh.max_row > 1 and sh.title != "00_DASHBOARD" else sh.auto_filter.ref

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# -----------------------------
# Main
# -----------------------------

def main() -> int:
    args = parse_args()
    t0 = time.time()
    workers = max(1, min(int(args.workers), os.cpu_count() or int(args.workers)))
    if workers != args.workers:
        print(f"[cpu] requested={args.workers}, available={os.cpu_count()}, using={workers}")
    else:
        print(f"[cpu] using {workers} worker processes")

    data = load_data(args.data_dir, args.period, args.self_test)
    features = build_features(data)
    n = len(features["dates"])
    # Common warmup protects against MA220 + DD252 + signal->next-open shift.
    eval_start = 260
    if n - eval_start - 1 < HORIZONS[0][1]:
        raise RuntimeError("Not enough post-warmup data for even 6M rolling analysis.")
    available_h = [label for label, h in HORIZONS if n - eval_start - 1 >= h]
    print(f"[data] common rows={n:,}, evaluation rows={n-eval_start-1:,}, rolling horizons={available_h}")

    seeds = seeded_configs()
    focus_share = min(0.9999, max(0.0, float(args.focus_share)))
    focus_n = int(max(0, args.trials) * focus_share)
    broad_n = max(0, args.trials) - focus_n
    focus_cfgs = generate_focus_configs(focus_n, args.seed)
    random_cfgs = generate_random_configs(broad_n, args.seed + 9001)
    configs = seeds + focus_cfgs + random_cfgs
    print(f"[search] candidates={len(configs):,} (seeded={len(seeds)}, focus={len(focus_cfgs):,}, broad={len(random_cfgs):,})")

    chunks = [configs[i:i + max(1, args.chunk_size)] for i in range(0, len(configs), max(1, args.chunk_size))]
    rows: List[Dict[str, Any]] = []

    # Windows-safe multiprocessing: initializer copies feature arrays once per process.
    if workers == 1:
        init_worker(features, args.cost_bps, eval_start)
        for ci, chunk in enumerate(chunks, 1):
            rows.extend(evaluate_chunk(chunk))
            if ci % max(1, len(chunks) // 20) == 0 or ci == len(chunks):
                print(f"[search] {ci}/{len(chunks)} chunks complete")
    else:
        with ProcessPoolExecutor(max_workers=workers, initializer=init_worker, initargs=(features, args.cost_bps, eval_start)) as ex:
            futs = [ex.submit(evaluate_chunk, chunk) for chunk in chunks]
            done = 0
            for fut in as_completed(futs):
                rows.extend(fut.result())
                done += 1
                if done % max(1, len(futs) // 20) == 0 or done == len(futs):
                    print(f"[search] {done}/{len(futs)} chunks complete")

    errors = [r for r in rows if "error" in r]
    if errors:
        print(f"[warning] {len(errors)} candidate errors; first={errors[0].get('error')}")
    good = [r for r in rows if "error" not in r]
    good.sort(key=lambda x: x["search_score"], reverse=True)
    print(f"[search] completed. valid={len(good):,}")
    if not good:
        raise RuntimeError("No valid candidates.")

    # Keep search-stage table before exact analysis.
    search_keep = good[:max(args.top_search, args.finalists * 3)]
    search_cols = [
        "name", "family", "search_score", "cagr", "mdd", "cagr_excess", "mdd_improvement",
        "min_return_win_rate", "avg_return_win_rate", "short_return_win_rate", "avg_p10_log_excess", "worst_log_excess", "transitions",
        "staged_entries", "staged_promotions", "explosive_entries",
        "correction_entries", "correction_tqqq_entries", "dd_rebound_entries",
        "trough_recovery_entries", "trough_recovery_qld_entries", "trough_recovery_promotions", "trough_recovery_hold_days", "trough_new_lows",
        "trough_split_qld_days", "trough_split_promotions", "trough_split_aborts",
        "repromote_guard_blocks", "repromote_guard_passes", "repromote_guard_regime_days",
        "soft_vol_signal_days", "soft_vol_entries",
        "impulse_veto_signal_days", "impulse_veto_blocks", "impulse_veto_score1_blocks", "impulse_veto_score2_blocks",
        "precision_veto_signal_days", "precision_veto_blocks", "precision_veto_trend_blocks", "precision_veto_mom_blocks", "precision_veto_score1_blocks", "precision_veto_fresh_blocks",
        "reentry_grace_signal_days", "reentry_grace_blocks", "reentry_grace_mom_blocks", "reentry_grace_trend_blocks", "reentry_grace_score2_blocks",
        "micro_rescue_signal_days", "micro_rescue_blocks", "micro_rescue_score1_pullback_blocks", "micro_rescue_score1_neartrend_blocks", "micro_rescue_mature_mom_blocks",
        "shadow_rescue_signal_days", "shadow_rescue_blocks", "shadow_rescue_score1_blocks", "shadow_rescue_mature_mom_blocks",
        "preview_overlay_signal_days", "preview_overlay_blocks", "preview_overlay_qld3_blocks", "preview_overlay_qld2_blocks", "preview_overlay_stepup_blocks",
        "adaptive_hold_days", "adaptive_exit_events", "tqqq_delever_events", "whipsaw_10d", "whipsaw_20d", "dead_v21_modules", "dead_v20_modules", "dead_v19_modules", "dead_v17_modules", "dead_v16_modules", "dead_v15_modules", "dead_v14_modules", "dead_v13_modules", "ret_wins_needed_79_1Y", "ret_1y_near_miss_25bp", "ret_1y_near_miss_50bp", "ret_1y_near_miss_100bp",
        "time_cash", "time_qqq", "time_qld", "time_tqqq", "gates",
    ] + [f"ret_win_{label}" for label, _ in HORIZONS]
    search_df = pd.DataFrame([{k: r.get(k) for k in search_cols} for r in search_keep])

    finalists = choose_finalists(good, max(1, args.finalists))
    print(f"[final] exact every-start rolling MDD analysis for {len(finalists)} finalists ...")
    bench_cache = build_benchmark_cache(features, eval_start)

    # Exact stage is intentionally sequential: each finalist uses a fast O(log N) range-drawdown tree,
    # and avoids multiplying memory by 12 while generating detailed rows.
    final_rows: List[Dict[str, Any]] = []
    all_sum: List[Dict[str, Any]] = []
    all_det: List[Dict[str, Any]] = []
    # Detail candidates determined later by preliminary exact order; collect details for all only if small.
    prelim: List[Tuple[Dict[str, Any], List[Dict[str, Any]], Config]] = []
    for i, c in enumerate(finalists, 1):
        row, summ, _ = exact_rolling_analysis(c, features, eval_start, args.cost_bps, bench_cache, include_details=False)
        prelim.append((row, summ, c))
        print(
            f"[final] {i:02d}/{len(finalists):02d} {c.name:28s} "
            f"min BOTH={row['min_both_win_rate']:.1%} min RET={row['min_return_win_rate']:.1%} "
            f"min MDD={row['min_mdd_win_rate']:.1%} strict={'YES' if row['strict_all_windows_both_win'] else 'NO'}"
        )

    prelim.sort(key=lambda x: x[0]["final_score"], reverse=True)
    detail_names = {x[0]["name"] for x in prelim[:max(0, args.detail_top)]}
    for row, summ, c in prelim:
        final_rows.append(row); all_sum.extend(summ)
        if c.name in detail_names and args.detail_top > 0:
            _, _, det = exact_rolling_analysis(c, features, eval_start, args.cost_bps, bench_cache, include_details=True)
            all_det.extend(det)

    rolling_summary = pd.DataFrame(all_sum)
    rolling_details = pd.DataFrame(all_det)
    ind_df = indicator_stats(final_rows, top_n=min(10, len(final_rows)))
    cohort_df = build_cohort_floor_table(rolling_details)
    ranked_final = sorted(final_rows, key=lambda x: x["final_score"], reverse=True)
    audit_n = max(0, min(int(args.audit_top), len(ranked_final)))
    audit_cfgs = [Config(**r["config"]) for r in ranked_final[:audit_n]]
    audit_df = build_reentry_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    transition_audit_df = build_transition_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    tqqq_reentry_audit_df = build_tqqq_reentry_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    impulse_veto_audit_df = build_impulse_veto_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    precision_veto_audit_df = build_precision_veto_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    v16_exit_candidate_df = build_v16_exit_candidate_audit(audit_cfgs, features, args.cost_bps, top_n=min(4, len(audit_cfgs)))
    v17_reentry_grace_audit_df = build_v17_reentry_grace_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    v17_reentry_candidate_df = build_v17_reentry_candidate_audit(audit_cfgs, features, args.cost_bps, top_n=min(4, len(audit_cfgs)))
    v19_micro_rescue_audit_df = build_v19_micro_rescue_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    v20_shadow_audit_df = build_v20_shadow_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    v21_preview_audit_df = build_v21_preview_audit(audit_cfgs, features, args.cost_bps, top_n=len(audit_cfgs))
    best_exact_name = ranked_final[0]["name"] if ranked_final else ""
    v19_nearmiss_1y_df = build_v19_nearmiss_1y(rolling_details, best_exact_name)
    v20_nearmiss_2y_df = build_v20_nearmiss_2y(rolling_details, best_exact_name)
    v21_nearmiss_2y_df = build_v21_nearmiss_2y(rolling_details, best_exact_name)

    outdir = Path(args.output_dir)
    outdir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx = outdir / f"ROBUST_TIGHT_V21_DASHBOARD_{stamp}.xlsx"
    run_info = {
        "Script_Version": SCRIPT_VERSION,
        "Run_Time": datetime.now().isoformat(timespec="seconds"),
        "Python": sys.version,
        "Platform": platform.platform(),
        "CPU_Count": os.cpu_count(),
        "Workers": workers,
        "Trials": args.trials,
        "Focus_Share": focus_share,
        "Candidates_Total": len(configs),
        "Finalists": len(finalists),
        "Detail_Top": args.detail_top,
        "Audit_Top": args.audit_top,
        "Seed": args.seed,
        "Cost_bps_per_side": args.cost_bps,
        "Data_Start": str(pd.to_datetime(features["dates"][0]).date()),
        "Data_End": str(pd.to_datetime(features["dates"][-1]).date()),
        "Evaluation_Start": str(pd.to_datetime(features["dates"][eval_start]).date()),
        "Evaluation_Rows": n - eval_start - 1,
        "Available_Horizons": available_h,
        "Objective": "V21: preserve V20_SHADOW_COMBO / 80.8629% BOTH + 98.1054% MDD Win; keep V20 stateless exit-hold overlay; add stateless forward TQQQ preview while control remains/steps to QLD; target the remaining five 2Y wins for 81%+97% and then 82%+97%; crash/recovery exits remain absolute; no calendar-specific rules.",
        "Execution": "Close signal -> next session Open; open-to-open return; adjusted prices; transaction costs applied.",
    }
    build_excel(xlsx, final_rows, rolling_summary, rolling_details, search_df, ind_df, cohort_df, audit_df, transition_audit_df, tqqq_reentry_audit_df, impulse_veto_audit_df, precision_veto_audit_df, v16_exit_candidate_df, v17_reentry_grace_audit_df, v17_reentry_candidate_df, v19_micro_rescue_audit_df, v19_nearmiss_1y_df, v20_shadow_audit_df, v20_nearmiss_2y_df, v21_preview_audit_df, v21_nearmiss_2y_df, run_info)

    if args.save_csv:
        pd.DataFrame([{k: v for k, v in r.items() if k != "config"} for r in final_rows]).to_csv(outdir / f"top_models_{stamp}.csv", index=False, encoding="utf-8-sig")
        rolling_summary.to_csv(outdir / f"rolling_summary_{stamp}.csv", index=False, encoding="utf-8-sig")

    best = sorted(final_rows, key=lambda x: x["final_score"], reverse=True)[0]
    elapsed = time.time() - t0
    print("\n" + "=" * 78)
    print("BEST RESULT")
    print("=" * 78)
    print(f"Model                  : {best['name']}")
    print(f"STRICT all-window pass : {'YES' if best['strict_all_windows_both_win'] else 'NO'}")
    print(f"Worst-horizon BOTH win : {best['min_both_win_rate']:.2%}")
    print(f"Worst-horizon RET win  : {best['min_return_win_rate']:.2%}")
    print(f"Worst-horizon MDD win  : {best['min_mdd_win_rate']:.2%}")
    print(f"Bottom-20% short cohort: {best['bottom20_startyear_short_win_rate']:.2%}")
    print(f"V21 base target passes  : {best['targets_passed']} / 3")
    print(f"Joint 79% + 97%         : {'YES' if best.get('target_joint_79_97') else 'NO'}")
    print(f"Joint 79.5% + 97%       : {'YES' if best.get('target_joint_795_97') else 'NO'}")
    print(f"Joint 80% + 97%         : {'YES' if best.get('target_joint_80_97') else 'NO'}")
    print(f"Joint 81% + 97%         : {'YES' if best.get('target_joint_81_97') else 'NO'}")
    print(f"Joint 82% + 97%         : {'YES' if best.get('target_joint_82_97') else 'NO'}")
    print(f"2Y wins / need 81%      : {best.get('two_year_both_wins', 0)} / +{best.get('two_year_wins_needed_81', 0)}")
    print(f"Worst-H wins needed 79/79.5/80: {best.get('worst_horizon_wins_needed_79', 0)} / {best.get('worst_horizon_wins_needed_795', 0)} / {best.get('worst_horizon_wins_needed_80', 0)}")
    print(f"Strategy CAGR / MDD    : {best['cagr']:.2%} / {best['mdd']:.2%}")
    print(f"TQQQ B&H CAGR / MDD    : {best['bench_cagr']:.2%} / {best['bench_mdd']:.2%}")
    print(f"Gates                  : {best['gates']}")
    print(f"Excel dashboard        : {xlsx.resolve()}")
    print(f"Elapsed seconds        : {elapsed:,.1f}")
    print("=" * 78)
    return 0


if __name__ == "__main__":
    freeze_support()
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n[stopped] Interrupted by user.")
        raise SystemExit(130)
    except Exception as exc:
        print("\n[ERROR]", exc)
        traceback.print_exc()
        raise SystemExit(1)
