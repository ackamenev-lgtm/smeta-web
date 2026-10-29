"""Разбор сообщения сотрудника и текст ответа по сделке."""

from __future__ import annotations

import re
from dataclasses import dataclass
from html import unescape

OWN_USER_ID = "316"
BOT_CHAT_TITLE = "Вопросы к ИИ"

STAGES = {
    "C6:NEW": "Первичный отбор",
    "C6:UC_0AH6PM": "ИИ не может скачать документы",
    "C6:PREPARATION": "Анализ и расчет",
    "C6:PREPAYMENT_INVOICE": "Решение об участии",
    "C6:EXECUTING": "Сбор документов",
    "C6:FINAL_INVOICE": "Заявка подана",
    "C6:UC_U4AIN4": "Торги (переторжка)",
    "C6:UC_UHPNOY": "Подведение итогов",
    "C6:UC_FUW8S1": "Не подались (отказ от участия)",
    "C6:UC_EDRIFD": "Жалоба УФАС",
    "C6:WON": "Контракт заключен",
    "C6:LOSE": "Невыигранные торги 2026",
}

HELP = (
    "Я AI Агент. Напишите номер сделки или номер закупки.\n"
    "Стадия 1930\n"
    "Где документы 4660147\n"
    "Посчитай лот 4660147\n"
    "Ответ приходит в этот чат в течение нескольких минут."
)

_NUMBER = re.compile(r"\d{3,}(?:-\d+)*")
_TAG = re.compile(r"<[^>]+>")
_HREF = re.compile(r"""href=["'](https?://[^"']+)["']""", re.I)
_URL = re.compile(r"https?://[^\s<>'\"]+")


@dataclass(frozen=True)
class Intent:
    kind: str
    query: str | None


def classify(text: str) -> Intent:
    raw = text or ""
    lowered = _plain(raw).lower()
    query = _query(lowered)
    if _is_calc(lowered):
        return Intent("calc", query)
    if _has_any(lowered, ("где документ", "папк", "конкурсн")):
        return Intent("docs", query)
    if _has_any(lowered, ("стади", "этап", "где сделк", "какая сделк")):
        return Intent("stage", query)
    if "номер" in lowered:
        return Intent("number", query)
    if query:
        return Intent("all", query)
    return Intent("help", None)


def should_answer(text: str, dialog_id: str, bot_dialog_id: str | None) -> bool:
    if bot_dialog_id and str(dialog_id) == str(bot_dialog_id):
        return True
    if not str(dialog_id).startswith("chat"):
        return True
    lowered = (text or "").lower()
    if f"[user={OWN_USER_ID}" in lowered:
        return True
    if "ai агент" in lowered:
        return True
    return False


def select_pending(messages: list[dict], own_id: str = OWN_USER_ID) -> dict | None:
    """Последнее чужое сообщение после нашей последней реплики."""
    last_own = -1
    for message in messages or []:
        if str(message.get("author_id")) == str(own_id):
            last_own = max(last_own, int(message.get("id") or 0))
    pending: list[dict] = []
    for message in messages or []:
        author = str(message.get("author_id"))
        message_id = int(message.get("id") or 0)
        if author in (str(own_id), "0", "None", ""):
            continue
        text = (message.get("text") or "").strip()
        if message_id > last_own and text:
            pending.append(message)
    if not pending:
        return None
    pending.sort(key=lambda item: int(item["id"]))
    return pending[-1]


def render(intent: Intent, deal: dict | None, comment_html: str = "") -> str:
    if intent.kind == "help" or (intent.kind != "calc" and not intent.query):
        return HELP
    if deal is None:
        return f"Не нашёл в воронке «Тендеры» сделку или закупку с номером {intent.query}."
    if intent.kind == "calc":
        card = _card(deal, comment_html, ("stage", "number", "docs"))
        return (
            "Запрос на расчёт принял и записал его в сделку. "
            "Считать лот будет агент по регламенту, результат появится в карточке сделки.\n"
            + card
        )
    parts = {
        "stage": ("stage",),
        "docs": ("docs",),
        "number": ("number",),
        "all": ("stage", "number", "docs"),
    }[intent.kind]
    return _card(deal, comment_html, parts)


def comment_body(author: str, text: str, deal: dict) -> str:
    safe_author = _plain(author) or "сотрудник"
    safe_text = _plain(text)
    title = _plain(deal.get("TITLE") or "")
    return (
        f"Запрос на расчёт из чата AI Агент.<br>"
        f"Кто написал: {safe_author}<br>"
        f"Сделка: {deal.get('ID')} {title}<br>"
        f"Текст: {safe_text}"
    )


def _card(deal: dict, comment_html: str, parts: tuple[str, ...]) -> str:
    lines = [f"Сделка {deal.get('ID')}. {_plain(deal.get('TITLE') or '')}".strip()]
    if "stage" in parts:
        stage_id = str(deal.get("STAGE_ID") or "")
        lines.append("Стадия: " + STAGES.get(stage_id, stage_id or "не указана"))
    if "number" in parts:
        number = deal.get("UF_CRM_TP_11") or "не указан"
        lines.append(f"Номер закупки: {number}")
    if "docs" in parts:
        lines.extend(_doc_lines(comment_html, deal))
    return "\n".join(lines)


def _doc_lines(comment_html: str, deal: dict) -> list[str]:
    lines: list[str] = []
    folders = [url for url in _urls(comment_html) if "focus.php" in url or "/disk/" in url]
    if folders:
        lines.append("Папка: " + folders[0])
    else:
        lines.append("Папки на диске в комментарии сделки нет.")
    for row in _comment_rows(comment_html):
        lowered = row.lower()
        if any(word in lowered for word in ("документац", "расчет", "расчёт", "не скачан", "пусто")):
            lines.append(row)
    platform = deal.get("UF_CRM_TP_10") or ""
    if platform:
        lines.append("Площадка: " + platform)
    return lines


def _comment_rows(comment_html: str) -> list[str]:
    text = re.sub(r"(?i)<br\s*/?>", "\n", comment_html or "")
    text = unescape(_TAG.sub("", text))
    rows = []
    for row in text.splitlines():
        clean = " ".join(row.split())
        if clean:
            rows.append(clean)
    return rows


def _urls(comment_html: str) -> list[str]:
    found = _HREF.findall(comment_html or "") + _URL.findall(comment_html or "")
    unique: list[str] = []
    for url in found:
        url = url.rstrip(".,)")
        if url not in unique:
            unique.append(url)
    return unique


def _query(lowered: str) -> str | None:
    numbers = _NUMBER.findall(lowered)
    if not numbers:
        return None
    for marker in ("сделк", "закупк", "лот", "номер"):
        position = lowered.find(marker)
        if position < 0:
            continue
        tail = _NUMBER.search(lowered[position:])
        if tail:
            return tail.group(0)
    return max(numbers, key=len)


def _is_calc(lowered: str) -> bool:
    return _has_any(lowered, ("посчита", "расчет", "расчёт", "расчит", "скачай документ"))


def _has_any(text: str, parts: tuple[str, ...]) -> bool:
    return any(part in text for part in parts)


def _plain(value: str) -> str:
    text = re.sub(r"\[/?[^\]]+\]", "", value or "")
    return " ".join(text.split())
