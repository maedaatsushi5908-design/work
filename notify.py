"""
通知モジュール

対応している通知先:
  - LINE Messaging API (pushメッセージ)
  - Discord Webhook
  - Slack Webhook
  - メール (SMTP / Gmail)

環境変数 (.env ファイル) で設定する。
"""

from __future__ import annotations

import os
import smtplib
from email.mime.text import MIMEText
from typing import Optional

import requests


def _env(key: str) -> Optional[str]:
    return os.environ.get(key) or None


def _post_json(url: str, payload: dict, headers: dict = {}) -> bool:
    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=10)
        return resp.status_code < 300
    except Exception as e:
        print(f"  通知失敗: {e}")
        return False


# ──────────────────────────────────────────
# LINE Messaging API
# ──────────────────────────────────────────

def notify_line(message: str) -> bool:
    """
    LINE Messaging API でプッシュ通知を送る。

    必要な環境変数:
      LINE_CHANNEL_ACCESS_TOKEN  チャンネルアクセストークン
      LINE_USER_ID               送り先のユーザーID（またはグループID）
    """
    token = _env("LINE_CHANNEL_ACCESS_TOKEN")
    user_id = _env("LINE_USER_ID")
    if not token or not user_id:
        return False

    url = "https://api.line.me/v2/bot/message/push"
    payload = {
        "to": user_id,
        "messages": [{"type": "text", "text": message}],
    }
    return _post_json(url, payload, headers={"Authorization": f"Bearer {token}"})


# ──────────────────────────────────────────
# Discord Webhook
# ──────────────────────────────────────────

def notify_discord(message: str) -> bool:
    """
    Discord Webhook で通知する。

    必要な環境変数:
      DISCORD_WEBHOOK_URL  ウェブフックURL
    """
    url = _env("DISCORD_WEBHOOK_URL")
    if not url:
        return False
    return _post_json(url, {"content": message})


# ──────────────────────────────────────────
# Slack Incoming Webhook
# ──────────────────────────────────────────

def notify_slack(message: str) -> bool:
    """
    Slack Incoming Webhook で通知する。

    必要な環境変数:
      SLACK_WEBHOOK_URL  ウェブフックURL
    """
    url = _env("SLACK_WEBHOOK_URL")
    if not url:
        return False
    return _post_json(url, {"text": message})


# ──────────────────────────────────────────
# メール (Gmail / SMTP)
# ──────────────────────────────────────────

def notify_email(message: str, subject: str = "【kenmo スキャン】本日のシグナル") -> bool:
    """
    SMTP メールで通知する。

    必要な環境変数:
      MAIL_SMTP_HOST  例: smtp.gmail.com
      MAIL_SMTP_PORT  例: 587
      MAIL_USER       送信元アドレス
      MAIL_PASSWORD   送信元パスワード（Gmailはアプリパスワード）
      MAIL_TO         宛先アドレス
    """
    host = _env("MAIL_SMTP_HOST")
    port = int(_env("MAIL_SMTP_PORT") or 587)
    user = _env("MAIL_USER")
    password = _env("MAIL_PASSWORD")
    to = _env("MAIL_TO")
    if not all([host, user, password, to]):
        return False

    msg = MIMEText(message, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = to
    try:
        with smtplib.SMTP(host, port) as smtp:
            smtp.starttls()
            smtp.login(user, password)
            smtp.send_message(msg)
        return True
    except Exception as e:
        print(f"  メール送信失敗: {e}")
        return False


# ──────────────────────────────────────────
# 一括送信
# ──────────────────────────────────────────

def send_all(message: str) -> None:
    """設定されている全通知先へ送信する。"""
    sent = False

    if _env("LINE_CHANNEL_ACCESS_TOKEN"):
        ok = notify_line(message)
        print(f"  LINE: {'✅ 送信成功' if ok else '❌ 送信失敗'}")
        sent = True

    if _env("DISCORD_WEBHOOK_URL"):
        ok = notify_discord(message)
        print(f"  Discord: {'✅ 送信成功' if ok else '❌ 送信失敗'}")
        sent = True

    if _env("SLACK_WEBHOOK_URL"):
        ok = notify_slack(message)
        print(f"  Slack: {'✅ 送信成功' if ok else '❌ 送信失敗'}")
        sent = True

    if _env("MAIL_SMTP_HOST"):
        ok = notify_email(message)
        print(f"  メール: {'✅ 送信成功' if ok else '❌ 送信失敗'}")
        sent = True

    if not sent:
        print("  ⚠️  通知先が設定されていません（.env を確認してください）")
