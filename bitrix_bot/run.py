"""Один проход: непрочитанные сообщения AI Агент и ответ в тот же чат."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from bitrix_bot.answer import (
    OWN_USER_ID,
    classify,
    comment_body,
    render,
    select_pending,
    should_answer,
)
from bitrix_bot.client import Bitrix, BitrixError, webhook_from_settings

ROOT = Path(__file__).resolve().parents[1]
SETTINGS = ROOT / "common" / "Тендерная воронка" / "settings.local.json"
CHAT_FILE = Path(__file__).resolve().parent / "chat.json"


def load_bot_dialog() -> str | None:
    if not CHAT_FILE.exists():
        return None
    with CHAT_FILE.open(encoding="utf-8") as handle:
        payload = json.load(handle)
    dialog_id = payload.get("dialog_id") or ""
    return str(dialog_id) or None


def dialogs_to_check(recent: list[dict], bot_dialog_id: str | None) -> list[str]:
    found: list[str] = []
    if bot_dialog_id:
        found.append(bot_dialog_id)
    for item in recent:
        if not item.get("counter") and not item.get("unread"):
            continue
        dialog_id = str(item.get("id") or "")
        if dialog_id and dialog_id not in found:
            found.append(dialog_id)
    return found


def reply_one(bitrix: Bitrix, dialog_id: str, bot_dialog_id: str | None) -> str:
    """Возвращает replied, skipped или failed."""
    messages = bitrix.messages(dialog_id)
    pending = select_pending(messages, OWN_USER_ID)
    if pending is None:
        bitrix.mark_read(dialog_id)
        return "skipped"
    text = pending.get("text") or ""
    if not should_answer(text, dialog_id, bot_dialog_id):
        bitrix.mark_read(dialog_id)
        return "skipped"
    intent = classify(text)
    deal = bitrix.find_deal(intent.query) if intent.query else None
    comment = bitrix.latest_comment(str(deal["ID"])) if deal else ""
    if intent.kind == "calc" and deal:
        author = bitrix.user_name(str(pending.get("author_id")))
        bitrix.add_deal_comment(str(deal["ID"]), comment_body(author, text, deal))
    bitrix.send(dialog_id, render(intent, deal, comment))
    bitrix.mark_read(dialog_id)
    return "replied"


def main() -> int:
    try:
        bitrix = Bitrix(webhook_from_settings(SETTINGS))
        bot_dialog_id = load_bot_dialog()
        dialogs = dialogs_to_check(bitrix.recent(), bot_dialog_id)
        replied = 0
        for dialog_id in dialogs:
            if reply_one(bitrix, dialog_id, bot_dialog_id) == "replied":
                replied += 1
        print(f"проверено чатов: {len(dialogs)}, ответов: {replied}")
        return 0
    except (BitrixError, OSError, json.JSONDecodeError) as error:
        print(f"бот остановился: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
