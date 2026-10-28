"""База цен материалов: искать здесь до интернета, пополнять из лотов.

    python tools/prices.py search "извещатель дымовой ИП 212"     # что уже есть
    python tools/prices.py top outbox/<лот>/суммы.json            # кому нужны 2 источника
    python tools/prices.py merge                                  # сложить outbox/*/Цены.json в базу
    python tools/prices.py import-xlsx                            # разово: забрать цены из бюджетов лотов

Цены.json лота — список объектов:
{"name","unit","shop","url","price_vat","date","note","vat_included": true}
price_vat — цена как на сайте. vat_included=false, если цена уже без НДС.
"""
import argparse
import csv
import json
import re
import sys
import warnings
from datetime import date, datetime
from pathlib import Path

import openpyxl

from paths import OUTBOX, PRICE_BASE

warnings.filterwarnings("ignore")

VAT = 1.22
FRESH_DAYS = 45
FIELDS = ["name", "unit", "shop", "url", "price_vat", "price_net", "date", "note", "source"]


def norm(s):
    s = str(s or "").lower().replace("ё", "е")
    return re.findall(r"[a-zа-я]+|\d+(?:[.,]\d+)?", s)


def stem(t):
    return t[: max(4, len(t) - 2)] if len(t) > 5 and not t[0].isdigit() else t


def load():
    if not PRICE_BASE.exists():
        return []
    with PRICE_BASE.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def save(rows):
    rows.sort(key=lambda r: (r["name"].lower(), r["unit"], r["date"]), reverse=False)
    PRICE_BASE.parent.mkdir(parents=True, exist_ok=True)
    with PRICE_BASE.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, FIELDS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def parse_date(s):
    s = str(s or "")[:10]
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None


def age_days(s):
    d = parse_date(s)
    return (date.today() - d).days if d else 10**6


def key_of(r):
    return (re.sub(r"\s+", " ", r["name"].strip().lower()),
            str(r["unit"]).strip().lower(),
            str(r["shop"]).strip().lower(),
            round(float(r["price_vat"] or 0), 2))


def add_rows(base, new_rows):
    """Новая запись с той же позицией, магазином и ценой заменяет старую, если дата новее."""
    idx = {key_of(r): i for i, r in enumerate(base)}
    added = updated = 0
    for r in new_rows:
        k = key_of(r)
        i = idx.get(k)
        if i is None:
            # та же позиция и магазин — оставить более свежую цену
            same = [j for j, b in enumerate(base)
                    if b["name"].strip().lower() == r["name"].strip().lower()
                    and str(b["shop"]).strip().lower() == str(r["shop"]).strip().lower()
                    and str(b["unit"]).strip().lower() == str(r["unit"]).strip().lower()]
            if same and age_days(r["date"]) <= age_days(base[same[0]]["date"]):
                base[same[0]] = r
                updated += 1
            elif not same:
                idx[k] = len(base)
                base.append(r)
                added += 1
        elif age_days(r["date"]) < age_days(base[i]["date"]):
            base[i] = r
            updated += 1
    return added, updated


def search(rows, query, n=8):
    q = [stem(t) for t in norm(query)]
    if not q:
        return []
    scored = []
    for r in rows:
        toks = norm(r["name"])
        hit = sum(1 for s in q if any(t.startswith(s) for t in toks))
        if hit:
            scored.append((hit / len(q), -age_days(r["date"]), r))
    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return [r for _, _, r in scored[:n]]


def row_from_offer(name, unit, shop, url, price_vat, when, note, source, vat_included=True):
    price_vat = round(float(price_vat), 2)
    net = round(price_vat / VAT, 2) if vat_included else price_vat
    gross = round(net * VAT, 2) if not vat_included else price_vat
    return {
        "name": str(name).strip(), "unit": str(unit or "").strip(), "shop": str(shop or "").strip(),
        "url": str(url or "").strip(), "price_vat": gross, "price_net": net,
        "date": str(when)[:10], "note": str(note or "").strip(), "source": source,
    }


def offers_from_sheet(ws, source, when):
    data = [list(r) for r in ws.iter_rows(values_only=True)]
    header_i = None
    for i, r in enumerate(data[:15]):
        if any("позиц" in str(c or "").lower() for c in r):
            header_i = i
            break
    if header_i is None:
        return []
    h = [str(c or "").lower() for c in data[header_i]]

    def col(*needles, exclude=()):
        for i, name in enumerate(h):
            if all(n in name for n in needles) and not any(x in name for x in exclude):
                return i
        return None

    name_c = col("позиц")
    unit_c = col("ед")
    out = []
    shop1 = col("магазин 1") or col("источник") or col("магазин")
    price1 = col("цена 1") or col("цена", exclude=("б/ндс", "без ндс", "розниц"))
    url1 = col("url 1") or col("url") or col("ссылк")
    shop2, price2, url2 = col("магазин 2"), col("цена 2"), col("url 2")
    note_c = col("отмет") or col("примечан")
    vat_c = col("с ндс")
    for r in data[header_i + 1:]:
        name = r[name_c] if name_c is not None and name_c < len(r) else None
        if not name or not str(name).strip() or str(name).strip().isdigit():
            continue
        unit = r[unit_c] if unit_c is not None and unit_c < len(r) else ""
        note = r[note_c] if note_c is not None and note_c < len(r) else ""

        def cell(i):
            return r[i] if i is not None and i < len(r) else None

        pairs = [(cell(shop1), cell(price1), cell(url1))]
        if shop2 is not None:
            pairs.append((cell(shop2), cell(price2), cell(url2)))
        vat_included = True
        mark = str(cell(vat_c) or "").lower()
        if "без" in mark or "б/ндс" in mark:
            vat_included = False
        for shop, price, url in pairs:
            if not isinstance(price, (int, float)) or price <= 0:
                continue
            out.append(row_from_offer(name, unit, shop, url, price, when, note, source, vat_included))
    return out


def import_xlsx():
    found = []
    for x in OUTBOX.glob("*/*.xlsx"):
        if x.name.startswith("~$") or "Реестр" in x.name:
            continue
        when = date.fromtimestamp(x.stat().st_mtime).isoformat()
        reg = x.parent / "Строка реестра.json"
        if reg.exists():
            when = json.loads(reg.read_text(encoding="utf-8")).get("Дата расчёта", when)
        try:
            wb = openpyxl.load_workbook(x, data_only=False, read_only=True)
        except Exception:
            continue
        for ws in wb.worksheets:
            if "источник" in ws.title.lower() or ws.title.lower().startswith("цены"):
                found += offers_from_sheet(ws, x.parent.name, when)
    return found


def from_json(path, source):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    out = []
    for r in data:
        if not r.get("price_vat"):
            continue
        out.append(row_from_offer(
            r.get("name"), r.get("unit"), r.get("shop"), r.get("url"), r["price_vat"],
            r.get("date") or date.today().isoformat(), r.get("note"), source,
            r.get("vat_included", True)))
    return out


def cmd_top(path):
    items = json.loads(Path(path).read_text(encoding="utf-8"))
    items = [i for i in items if i.get("sum")]
    total = sum(i["sum"] for i in items) or 1
    items.sort(key=lambda i: i["sum"], reverse=True)
    cum = 0
    print("позиция | сумма | доля | что делать")
    for i in items:
        share = i["sum"] / total
        equip = any(w in i["name"].lower() for w in ("оборудован", "установк", "ворот", "щит"))
        if cum < 0.80 or share >= 0.05 or equip:
            flag = "два источника" + (", запросить КП" if share >= 0.05 or equip else "")
        else:
            flag = "один источник, база или аналог с пометкой «оценка»"
        cum += share
        print(f"{i['name']} | {i['sum']:.0f} | {share:.0%} | {flag}")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("cmd")
    p.add_argument("args", nargs="*")
    p.add_argument("-n", type=int, default=8)
    a = p.parse_args()

    if a.cmd == "search":
        rows = search(load(), " ".join(a.args), a.n)
        if not rows:
            print("в базе нет")
            return
        print("позиция | ед. | магазин | с НДС | б/НДС | дата | возраст | источник")
        for r in rows:
            days = age_days(r["date"])
            fresh = "свежая" if days <= FRESH_DAYS else f"устарела {days} дн."
            print(f"{r['name']} | {r['unit']} | {r['shop']} | {r['price_vat']} | {r['price_net']} | "
                  f"{r['date']} | {fresh} | {r['source']}")
            if r["url"]:
                print(f"    {r['url']}")
    elif a.cmd == "top":
        cmd_top(a.args[0])
    elif a.cmd == "merge":
        base = load()
        new = []
        for f in OUTBOX.glob("*/Цены.json"):
            new += from_json(f, f.parent.name)
        added, updated = add_rows(base, new)
        save(base)
        print(f"база: {len(base)} цен, добавлено {added}, обновлено {updated}")
    elif a.cmd == "import-xlsx":
        base = load()
        new = import_xlsx()
        added, updated = add_rows(base, new)
        save(base)
        print(f"из бюджетов: {len(new)} цен; база: {len(base)}, добавлено {added}, обновлено {updated}")
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
