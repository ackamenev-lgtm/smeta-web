"""Один раз создаёт открытый чат «Вопросы к ИИ» и записывает его id."""

from __future__ import annotations

import json
from pathlib import Path

from bitrix_bot.answer import BOT_CHAT_TITLE, HELP
from bitrix_bot.client import Bitrix, webhook_from_settings
from bitrix_bot.run import CHAT_FILE, SETTINGS

WELCOME = "Пишите сюда или мне в личные сообщения.\n" + HELP


def existing_dialog(bitrix: Bitrix) -> str | None:
    payload = bitrix.call("im.search.chat.list.json", {"FIND": BOT_CHAT_TITLE})
    for chat in payload.get("result") or []:
        title = chat.get("title") or chat.get("name") or ""
        if title == BOT_CHAT_TITLE:
            chat_id = chat.get("id") or chat.get("chat_id")
            if chat_id:
                return f"chat{chat_id}" if not str(chat_id).startswith("chat") else str(chat_id)
    return None


def main() -> None:
    if CHAT_FILE.exists():
        print(f"чат уже записан в {CHAT_FILE}")
        return
    bitrix = Bitrix(webhook_from_settings(SETTINGS))
    dialog_id = existing_dialog(bitrix)
    if dialog_id is None:
        payload = bitrix.call(
            "im.chat.add.json",
            {
                "USERS[]": [1],
                "TYPE": "OPEN",
                "TITLE": BOT_CHAT_TITLE,
                "DESCRIPTION": "Стадия сделки, папка документов и запрос на расчёт.",
                "COLOR": "AZURE",
                "MESSAGE": WELCOME,
            },
        )
        chat_id = payload.get("result")
        dialog_id = f"chat{chat_id}"
        print(f"создан {dialog_id}")
    else:
        print(f"уже есть {dialog_id}")
    CHAT_FILE.write_text(
        json.dumps({"dialog_id": dialog_id, "title": BOT_CHAT_TITLE}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(CHAT_FILE)


if __name__ == "__main__":
    main()
