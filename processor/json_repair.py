"""
Восстановление обрезанного JSON из ответов Claude API.
"""
import json
import re


def repair_json(raw: str) -> dict:
    """
    Пытается восстановить обрезанный JSON несколькими способами.
    Raises ValueError если не удалось.
    """
    raw = raw.strip().replace('```json', '').replace('```', '').strip()

    # 1. Сначала пробуем как есть
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # 2. Убираем trailing запятые
    fixed = re.sub(r',(\s*[}\]])', r'\1', raw)
    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    # 3. Обрезанный JSON — закрываем открытые скобки
    def close_truncated(s):
        stack = []
        in_str = False
        escape = False
        result = []
        for i, ch in enumerate(s):
            if escape:
                escape = False
                result.append(ch)
                continue
            if ch == '\\' and in_str:
                escape = True
                result.append(ch)
                continue
            if ch == '"' and not in_str:
                in_str = True
                result.append(ch)
                continue
            if ch == '"' and in_str:
                in_str = False
                result.append(ch)
                continue
            if in_str:
                result.append(ch)
                continue
            if ch in '{[':
                stack.append(ch)
            elif ch in '}]':
                if stack:
                    stack.pop()
            result.append(ch)

        # Закрываем незакрытые строки
        if in_str:
            result.append('"')

        # Закрываем незакрытые скобки в обратном порядке
        while stack:
            opener = stack.pop()
            if opener == '{':
                result.append('}')
            else:
                result.append(']')

        return ''.join(result)

    closed = close_truncated(fixed)
    # Убираем trailing запятые снова после закрытия
    closed = re.sub(r',(\s*[}\]])', r'\1', closed)
    try:
        return json.loads(closed)
    except json.JSONDecodeError:
        pass

    # 4. Ищем последний валидный объект
    for end in range(len(raw), 0, -100):
        chunk = raw[:end]
        chunk = re.sub(r',(\s*[}\]])', r'\1', chunk)
        try:
            return json.loads(chunk + '}' * 10)
        except json.JSONDecodeError:
            continue

    raise ValueError(f"Не удалось восстановить JSON (первые 200 символов): {raw[:200]}")
