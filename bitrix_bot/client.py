"""Вызовы входящего вебхука AI Агент. Секрет в журнал не пишется."""

from __future__ import annotations

import json
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from bitrix_bot.answer import OWN_USER_ID


class BitrixError(RuntimeError):
    pass


class Bitrix:
    def __init__(self, webhook_url: str, timeout: int = 40):
        self.webhook_url = webhook_url.rstrip("/") + "/"
        self.timeout = timeout

    def call(self, method: str, params: dict | None = None) -> dict:
        data = urlencode(params or {}, doseq=True).encode()
        request = Request(self.webhook_url + method, data=data, method="POST")
        try:
            with urlopen(request, timeout=self.timeout) as response:
                payload = json.load(response)
        except HTTPError as error:
            raw = error.read().decode("utf-8", "replace")
            raise BitrixError(f"{method} ответил {error.code}: {raw[:240]}") from error
        if payload.get("error"):
            raise BitrixError(
                f"{method}: {payload.get('error')} {payload.get('error_description') or ''}".strip()
            )
        return payload

    def recent(self) -> list[dict]:
        payload = self.call("im.recent.get.json")
        result = payload.get("result") or []
        return result if isinstance(result, list) else []

    def messages(self, dialog_id: str, limit: int = 30) -> list[dict]:
        payload = self.call(
            "im.dialog.messages.get.json",
            {"DIALOG_ID": dialog_id, "LIMIT": limit},
        )
        result = payload.get("result") or {}
        return list(result.get("messages") or [])

    def send(self, dialog_id: str, text: str) -> None:
        self.call("im.message.add.json", {"DIALOG_ID": dialog_id, "MESSAGE": text})

    def mark_read(self, dialog_id: str) -> None:
        self.call("im.dialog.read.json", {"DIALOG_ID": dialog_id})

    def user_name(self, user_id: str) -> str:
        payload = self.call("im.user.get.json", {"ID": user_id})
        user = payload.get("result") or {}
        name = " ".join(
            part
            for part in (user.get("first_name") or user.get("name"), user.get("last_name"))
            if part
        )
        return name.strip() or f"сотрудник {user_id}"

    def find_deal(self, query: str) -> dict | None:
        by_number = self.call(
            "crm.deal.list.json",
            {
                "filter[CATEGORY_ID]": 6,
                "filter[=UF_CRM_TP_11]": query,
                "order[ID]": "DESC",
                "select[]": ["ID", "TITLE", "STAGE_ID", "CATEGORY_ID", "UF_CRM_TP_11", "UF_CRM_TP_10"],
                "start": 0,
            },
        )
        rows = by_number.get("result") or []
        if rows:
            return rows[0]
        if not query.isdigit():
            return None
        payload = self.call("crm.deal.get.json", {"id": query})
        deal = payload.get("result") or None
        if not deal or str(deal.get("CATEGORY_ID")) != "6":
            return None
        return deal

    def latest_comment(self, deal_id: str) -> str:
        payload = self.call(
            "crm.timeline.comment.list.json",
            {
                "filter[ENTITY_ID]": deal_id,
                "filter[ENTITY_TYPE]": "deal",
                "order[ID]": "DESC",
                "start": 0,
            },
        )
        rows = payload.get("result") or []
        if not rows:
            return ""
        return rows[0].get("COMMENT") or ""

    def add_deal_comment(self, deal_id: str, text: str) -> None:
        self.call(
            "crm.timeline.comment.add.json",
            {
                "fields[ENTITY_ID]": deal_id,
                "fields[ENTITY_TYPE]": "deal",
                "fields[COMMENT]": text,
            },
        )


def webhook_from_settings(settings_path) -> str:
    with open(settings_path, encoding="utf-8") as handle:
        payload = json.load(handle)
    url = (payload.get("env") or {}).get("BITRIX24_WEBHOOK_URL") or ""
    if not url:
        raise BitrixError("В настройках нет BITRIX24_WEBHOOK_URL")
    if f"/rest/{OWN_USER_ID}/" not in url:
        raise BitrixError("Вебхук в настройках не принадлежит сотруднику AI Агент")
    return url
