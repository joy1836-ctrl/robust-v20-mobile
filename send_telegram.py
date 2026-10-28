#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Send the daily ROBUST V20 signal to Telegram.

Secrets (GitHub repo -> Settings -> Secrets and variables -> Actions):
  TELEGRAM_BOT_TOKEN   token from @BotFather            (never paste it into a chat or commit it)
  TELEGRAM_CHAT_ID     your chat id
Optional:
  ACCOUNT_JSON         {"cash":10000,"QQQ":0,"QLD":0,"TQQQ":100}  -> adds an order plan for your account
  APP_URL              link to the mobile app, appended to the message

  python send_telegram.py --dry-run     # print the message, send nothing
"""
from __future__ import annotations
import argparse, html, json, math, os, sys, urllib.error, urllib.parse, urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

KST = timezone(timedelta(hours=9))
BUF, FEE = 0.005, 0.0005   # cash buffer 0.5%, expected cost 0.05% (strategy manual)


def plan(sig, acct):
    """Order plan from the manual: fresh transition -> sell non-target, size target; HOLD -> only non-target sells."""
    px, tgt, fresh = sig["closes"], sig["target_asset"], sig["guide"]["fresh"]
    hold = {a: int(acct.get(a, 0) or 0) for a in ("QQQ", "QLD", "TQQQ")}
    cash = float(acct.get("cash", 0) or 0)
    total = cash + sum(hold[a] * px[a] for a in hold)
    orders, proceeds, spend = [], 0.0, 0.0
    for a in hold:
        if a != tgt and hold[a] > 0:
            orders.append(("매도", a, hold[a], px[a])); proceeds += hold[a] * px[a]
    if tgt != "CASH":
        if fresh:
            want = math.floor(total * (1 - BUF) / (px[tgt] * (1 + FEE)))
            d = want - hold[tgt]
            if d > 0: orders.append(("매수", tgt, d, px[tgt])); spend = d * px[tgt]
            elif d < 0: orders.append(("매도", tgt, -d, px[tgt])); proceeds += -d * px[tgt]
        elif sig["indicators"]["entry_gate_pass"] and cash > 0:
            avail = cash + proceeds
            q = math.floor(avail * (1 - BUF) / (px[tgt] * (1 + FEE)))
            if q > 0: orders.append(("매수", tgt, q, px[tgt])); spend = q * px[tgt]
    left = cash + proceeds - spend - (proceeds + spend) * FEE
    return orders, total, cash, left


def build_message(sig, acct=None, app_url=""):
    e = html.escape
    g, ind = sig["guide"], sig["indicators"]
    gen = ""
    if sig.get("generated_utc"):
        gen = datetime.strptime(sig["generated_utc"], "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc).astimezone(KST).strftime("%Y-%m-%d %H:%M")
    L = []
    if sig["integrity"].get("synthetic_data"):
        L.append("⚠️ <b>[테스트 데이터] 매매에 쓰지 마세요</b>")
    L += ["<b>ROBUST V20</b>", f"<b>{e(g['status'])}</b>",
          f"신호 기준: {sig['asof_close']} / 다음 실행일: {sig['next_exec_date']}",
          f"유효기간 UTC: {sig['valid_until_utc']}", "",
          f"<b>현재 목표: {e(sig['target_asset'])}</b>"]
    L += [f"{a}: {e(t)}" for a, t in g["assets"].items()]
    if g["note"]:
        L.append(e(g["note"]))
    L += ["", f"전략 상태전이 (참고): {sig['prev_asset']} -> {sig['target_asset']} / {sig['signal']}"]
    if sig["signal"] == "HOLD":
        L.append("<i>HOLD는 상태전이가 없다는 뜻이며, 지속 매매 안내를 중단하지 않습니다.</i>")
    L += ["", f"Core {ind['core_score']}/3 · 게이트 {'OPEN' if ind['entry_gate_pass'] else 'CLOSED'} · 급락 {'발동' if ind['crash'] else '정상'}",
          f"RSI {ind['rsi14']:.1f} · MFI {ind['mfi14']:.1f} · ATR {ind['atr_pct14']*100:.2f}% · VOL40 {ind['vol40']*100:.1f}% · DD252 {ind['dd252']*100:.1f}%"]
    L += ["", "<b>입력 계좌 기준 주문안</b>"]
    if acct:
        orders, total, cash, left = plan(sig, acct)
        if orders:
            L += [f"{s} {a} {q:,}주 @ ${p:,.2f} (≈${q*p:,.2f})" for s, a, q, p in orders]
        else:
            L.append("입력 잔고 기준 주문안 없음")
            L.append("매매 신호가 종료됐다는 뜻은 아닙니다.")
        L.append(f"평가액 ${total:,.2f} / 현금 ${cash:,.2f} / 예상 잔여 ${left:,.2f}")
    else:
        L.append("계좌 미설정: 앱에서 보유수량을 입력하면 수량이 계산됩니다.")
    L += ["", "<i>실제 체결·입금 후 잔고를 수정하세요. 자동 주문·자동 체결 처리는 하지 않습니다.\n"
          "주문안은 종가 기준 추정치이며 실제 시가·수수료·매수가능액을 다시 확인하세요.</i>"]
    if gen: L.append(f"데이터 생성(KST): {gen}")
    if app_url: L.append(app_url)
    return "\n".join(L)


def send(token, chat, text):
    req = urllib.request.Request(f"https://api.telegram.org/bot{token}/sendMessage",
        data=urllib.parse.urlencode({"chat_id": chat, "text": text, "parse_mode": "HTML", "disable_web_page_preview": "true"}).encode())
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            body = json.loads(r.read().decode())
    except urllib.error.HTTPError as ex:   # show Telegram's reason (never prints the token)
        reason = ex.read().decode("utf-8", "replace")
        hints = {401: "토큰이 틀렸거나 폐기됨 (TELEGRAM_BOT_TOKEN 확인)",
                 400: "chat id가 틀렸거나 봇과 대화 시작 전 (TELEGRAM_CHAT_ID 확인, 봇에게 먼저 메시지 보내기)",
                 403: "봇이 차단됐거나 대화 시작 전", 404: "토큰 형식 오류 (TELEGRAM_BOT_TOKEN 확인)"}
        print(f"[telegram] HTTP {ex.code}: {reason}\n[힌트] {hints.get(ex.code, '')}", file=sys.stderr)
        import re
        shape = "정상 형태" if re.fullmatch(r"\d{6,12}:[A-Za-z0-9_-]{30,}", token) else "형태 이상(숫자:영문 형식이 아님)"
        print(f"[진단] 토큰 길이={len(token)}자, 콜론 개수={token.count(':')}, 형태={shape}, chat_id 자릿수={len(str(chat))}", file=sys.stderr)
        raise SystemExit(1)
    if not body.get("ok"):
        raise RuntimeError(body)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--signal", default="docs/signal.json")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    sig = json.loads(Path(a.signal).read_text(encoding="utf-8"))
    acct = None
    if os.environ.get("ACCOUNT_JSON"):
        try: acct = json.loads(os.environ["ACCOUNT_JSON"])
        except Exception: print("[warn] ACCOUNT_JSON is not valid JSON; ignored", file=sys.stderr)
    msg = build_message(sig, acct, os.environ.get("APP_URL", ""))
    if a.dry_run:
        print(msg); return 0
    token, chat = (os.environ.get("TELEGRAM_BOT_TOKEN") or "").strip(), (os.environ.get("TELEGRAM_CHAT_ID") or "").strip()
    if not token or not chat:
        print("TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID not set; skipping send"); return 0
    if token.lower().startswith("bot"): token = token[3:]   # tolerate a pasted "bot" prefix
    send(token, chat, msg)
    print("sent")
    return 0


if __name__ == "__main__":
    sys.exit(main())
