"""Дописать утверждённые работы в конец тарификатора.

Строку в середину листов не вставляем: у «Расценки» и «Нормы и нормо-часы»
формулы завязаны на номер строки, строка норм = строка расценки + 1.
Новая позиция встаёт следом за последней, в блок «Утверждено из расчётов».
Цена НЕОСТРОЙ — формула: не ниже рынка и не ниже оплаты бригаде с запасом
25 % (у техники TEX — 15 %).

    python tools/tarif_add.py --draft outbox/<лот>/Расчёт.json
    python tools/tarif_add.py outbox/<лот>/В тарификатор.json --dry-run
    python tools/tarif_add.py outbox/<лот>/В тарификатор.json

Черновик из расчёта справочник не меняет: в нём стоит "ready": false.
После проверки цифр ключ убирают и запускают запись.
Пишет три книги: тарификатор, расценки для сотрудников, лист «Тарификатор»
в файле материалов. Лот сам эти файлы не трогает — это командный чат.
"""
import json
import re
import sys
import zipfile
from datetime import date
from decimal import Decimal, ROUND_CEILING, ROUND_HALF_UP
from html import unescape
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from paths import CACHE, MATERIALS, STAFF, TARIF

import openpyxl

APPROVED = "Утверждено из расчётов"
MARGIN = Decimal("0.25")
TEX_KEEP = Decimal("0.85")


def D(value):
    if isinstance(value, Decimal):
        return value
    if isinstance(value, bool) or value is None or value == "":
        raise ValueError("пустое число")
    if isinstance(value, int):
        return Decimal(value)
    text = str(value).strip().replace(" ", "").replace(",", ".")
    return Decimal(text)


def xround(value, digits=0):
    quantum = Decimal(1).scaleb(-digits)
    return D(value).quantize(quantum, rounding=ROUND_HALF_UP)


def xceil(value):
    return D(value).quantize(Decimal(1), rounding=ROUND_CEILING)


def num(value):
    if value is None or value == "":
        return ""
    text = format(D(value), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def xml_text(value):
    return (str(value).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def die(message):
    raise SystemExit(message)


def pick(item, *keys):
    for key in keys:
        if key in item and item[key] not in (None, ""):
            return item[key]
    return None


def col_letters(n):
    out = []
    for i in range(1, n + 1):
        x, s = i, ""
        while x:
            x, r = divmod(x - 1, 26)
            s = chr(65 + r) + s
        out.append(s)
    return out


def sheet_target(workbook_xml, rels_xml, title):
    match = re.search(rf'<sheet name="{re.escape(title)}" [^>]*r:id="(rId\d+)"', workbook_xml)
    if not match:
        die(f"в книге нет листа «{title}»")
    rel = re.search(rf'Id="{match.group(1)}" [^>]*Target="([^"]+)"', rels_xml)
    target = rel.group(1)
    return target if target.startswith("xl/") else "xl/" + target


def dimension_last(xml):
    match = re.search(r'<dimension ref="[A-Z]+\d+:[A-Z]+(\d+)"', xml)
    if not match:
        die("у листа нет dimension")
    return int(match.group(1))


def last_row_tag(xml):
    rows = re.findall(r'<row r="(\d+)"', xml)
    return int(rows[-1]) if rows else 0


def extract_row(xml, row):
    match = re.search(rf'<row r="{row}"[^>]*>.*?</row>', xml)
    if not match:
        die(f"нет строки {row}")
    return match.group(0)


def row_styles(xml, row):
    body = extract_row(xml, row)
    styles = {}
    for cell in re.finditer(rf'<c r="([A-Z]{{1,3}}){row}"([^>]*?)(?:/>|>)', body):
        style = re.search(r's="(\d+)"', cell.group(2))
        styles[cell.group(1)] = style.group(1) if style else ""
    return styles


def extend_tail(xml, old, new):
    """Продлить dimension, автофильтр и условное форматирование до новой последней строки.

    Формулы внутри листа не трогаем: там номер строки — это ссылка конкретной ячейки.
    """
    split = xml.find("</sheetData>")
    if split < 0:
        die("у листа нет sheetData")
    head, tail = xml[:split], xml[split:]

    def grow_ref(match):
        ref = match.group(2)
        grown = re.sub(rf"([A-Z]){old}$", rf"\g<1>{new}", ref)
        return match.group(1) + grown + match.group(3)

    head, n = re.subn(r'(<dimension ref=")([^"]+)(")', grow_ref, head, count=1)
    if n != 1 or f"{new}" not in re.search(r'<dimension ref="[^"]+"', head).group(0):
        die(f"не удалось продлить dimension с {old}")

    def grow_token(match):
        return match.group(1) + str(new) if int(match.group(2)) == old else match.group(0)

    tail = re.sub(r"([A-Z]{1,3})(\d+)", grow_token, tail)
    return head + tail


def must_replace(xml, old, new, label):
    if old not in xml:
        die(f"{label}: в файле нет «{old}»")
    return xml.replace(old, new)


class Strings:
    def __init__(self, xml):
        self.xml = xml
        self.map = {}
        count = 0
        for match in re.finditer(r"<si>(.*?)</si>", xml, flags=re.S):
            text = unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", match.group(1), flags=re.S)))
            self.map.setdefault(text, count)
            count += 1
        unique = int(re.search(r'uniqueCount="(\d+)"', xml).group(1))
        if unique != count:
            die(f"справочник строк разошёлся: uniqueCount {unique}, фактически {count}")
        self.unique = unique
        self.total = int(re.search(r'\bcount="(\d+)"', xml).group(1))
        self.added = []
        self.uses = 0

    def idx(self, text):
        text = "" if text is None else str(text)
        self.uses += 1
        if text in self.map:
            return self.map[text]
        index = self.unique + len(self.added)
        self.map[text] = index
        self.added.append(text)
        return index

    def render(self):
        extra = []
        for text in self.added:
            space = ' xml:space="preserve"' if text[:1].isspace() or text[-1:].isspace() else ""
            extra.append(f"<si><t{space}>{xml_text(text)}</t></si>")
        xml = self.xml.replace("</sst>", "".join(extra) + "</sst>", 1)
        xml = re.sub(r'\bcount="\d+"', f'count="{self.total + self.uses}"', xml, count=1)
        xml = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{self.unique + len(self.added)}"', xml, count=1)
        return xml


def c_num(col, row, style, value, formula=None):
    style_attr = f' s="{style}"' if style else ""
    ref = f"{col}{row}"
    if formula:
        cached = f"<v>{num(value)}</v>" if value not in (None, "") else ""
        return f'<c r="{ref}"{style_attr}><f>{xml_text(formula)}</f>{cached}</c>'
    if value in (None, ""):
        return f'<c r="{ref}"{style_attr}/>'
    return f'<c r="{ref}"{style_attr}><v>{num(value)}</v></c>'


def c_str(col, row, style, index):
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{col}{row}"{style_attr} t="s"><v>{index}</v></c>'


def c_inline(col, row, style, formula, text):
    style_attr = f' s="{style}"' if style else ""
    return (f'<c r="{col}{row}"{style_attr} t="str"><f>{xml_text(formula)}</f>'
            f"<v>{xml_text(text)}</v></c>")


def header_row(row, styles, title, strings):
    letters = list(styles)
    fill = styles.get("B") or "6"
    cells = [f'<c r="A{row}" s="{styles.get("A") or fill}" t="s"><v>{strings.idx(title)}</v></c>']
    for col in letters[1:]:
        cells.append(f'<c r="{col}{row}" s="{styles.get(col) or fill}"/>')
    return (f'<row r="{row}" spans="1:{len(letters)}" x14ac:dyDescent="0.25">'
            f'{"".join(cells)}</row>')


def price_of(market, hourly, brigade, spec):
    if market <= 0:
        return None
    parts = [market]
    if hourly > 0:
        parts.append(xceil(hourly / (1 - MARGIN)))
    if brigade > 0:
        denom = TEX_KEEP if spec == "TEX" else (1 - MARGIN)
        parts.append(xceil(brigade / denom))
    return max(parts)


def rates_price_formula(row, norms_row, market_col, norm_col, brig_col, spec):
    piece_denom = "0.85" if spec == "TEX" else "(1-$F$3)"
    return (
        f'IF(OR({market_col}{row}="",{market_col}{row}=0),"",'
        f'MAX({market_col}{row},'
        f"IF(ISNUMBER('Нормы и нормо-часы'!{norm_col}{norms_row}),"
        f"ROUNDUP('Нормы и нормо-часы'!{norm_col}{norms_row}/(1-$F$3),0),0),"
        f'IF(OR({brig_col}{row}="",{brig_col}{row}=0),0,'
        f'ROUNDUP({brig_col}{row}/{piece_denom},0))))'
    )


def norms_formulas(q, r, s_last):
    ref = "'Нормо-час по специальностям'"
    a = f"{ref}!$A$6:$A${s_last}"
    return {
        "E": f"INDEX({ref}!$B$6:$B${s_last},MATCH(D{q},{a},0))",
        "H": f'IF(OR(F{q}="",F{q}=0),"",ROUND($G$3/F{q},2))',
        "I": f"INDEX({ref}!$D$6:$D${s_last},MATCH(D{q},{a},0))",
        "J": f"INDEX({ref}!$E$6:$E${s_last},MATCH(D{q},{a},0))",
        "K": f'IF(OR(F{q}="",F{q}=0,I{q}=0),"",ROUND(F{q}*I{q},0))',
        "L": f'IF(OR(F{q}="",F{q}=0,J{q}=0),"",ROUND(F{q}*J{q},0))',
        "M": f"Расценки!L{r}",
        "N": f"Расценки!M{r}",
        "O": f'IF(OR(K{q}="",M{q}="",M{q}=0),"",K{q}/M{q}-1)',
        "P": f'IF(OR(L{q}="",N{q}="",N{q}=0),"",L{q}/N{q}-1)',
        "Q": f'IF(O{q}="","—",IF(O{q}<-$C$3,"нормо-часы",IF(O{q}>$C$3,"расценка","равно")))',
        "R": f'IF(P{q}="","—",IF(P{q}<-$C$3,"нормо-часы",IF(P{q}>$C$3,"расценка","равно")))',
        "S": f'IF(OR(F{q}="",F{q}=0,M{q}=""),"",ROUND(M{q}/F{q},0))',
        "T": f'IF(OR(F{q}="",F{q}=0,N{q}=""),"",ROUND(N{q}/F{q},0))',
        "U": f"Расценки!F{r}",
        "V": f"Расценки!G{r}",
        "W": f'IF(OR(K{q}="",U{q}="",U{q}=0),"",(U{q}-K{q})/U{q})',
        "X": f'IF(OR(L{q}="",V{q}="",V{q}=0),"",(V{q}-L{q})/V{q})',
        "Y": f"Расценки!H{r}",
        "Z": f"Расценки!I{r}",
        "AA": f"Расценки!J{r}",
        "AB": f"Расценки!K{r}",
        "AC": f'IF(OR(K{q}="",Y{q}="",Y{q}=0),"",(Y{q}-K{q})/Y{q})',
        "AD": f'IF(OR(L{q}="",Z{q}="",Z{q}=0),"",(Z{q}-L{q})/Z{q})',
        "AE": f'IF(OR(Y{q}="",Y{q}=0),"",(Y{q}-M{q})/Y{q})',
        "AF": f'IF(OR(Z{q}="",Z{q}=0),"",(Z{q}-N{q})/Z{q})',
        "AG": f'IF(OR(K{q}="",Y{q}=""),"",Y{q}-K{q})',
        "AH": f'IF(OR(L{q}="",Z{q}=""),"",Z{q}-L{q})',
    }


def better_text(gap, threshold):
    if gap is None:
        return "—"
    if gap < -threshold:
        return "нормо-часы"
    if gap > threshold:
        return "расценка"
    return "равно"


def gap(cost, base):
    if not base:
        return None
    return (cost - base) / base


def load_rates(path, sheet, code_col=0, name_col=1, spb_col=3, msk_col=4, min_row=6):
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rates = {}
    for row in book[sheet].iter_rows(min_row=min_row, max_col=max(msk_col, name_col) + 1, values_only=True):
        code = row[code_col]
        if code and row[spb_col] is not None:
            rates[str(code).strip()] = {
                "name": row[name_col],
                "h_spb": D(row[spb_col]),
                "h_msk": D(row[msk_col] if row[msk_col] is not None else row[spb_col]),
            }
    book.close()
    return rates


def existing(path):
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    arts, names, last_section = set(), set(), ""
    for row in book["Расценки"].iter_rows(min_row=5, max_col=3, values_only=True):
        art, name, unit = row
        if art is None:
            continue
        if name is None and unit is None:
            last_section = str(art).strip()
            continue
        if unit is None:
            continue
        arts.add(str(art).strip())
        names.add((str(name).strip().casefold(), str(unit).strip().casefold()))
    book.close()
    return arts, names, last_section


def staff_last_section(path):
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    last = ""
    for row in book["Расценки"].iter_rows(min_row=5, max_col=3, values_only=True):
        art, name, unit = row
        if art is not None and name is None and unit is None:
            last = str(art).strip()
    book.close()
    return last


def staff_params(path):
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    values = {}
    for row in book["Параметры"].iter_rows(min_row=3, max_col=2, values_only=True):
        if row[0]:
            values[str(row[0])] = row[1]
    book.close()
    return {
        "shift": D(values["Оплачиваемых часов в смене"]),
        "uplift": D(values["Потолок сделки: надбавка к цели"]),
        "cap": D(values["Потолок сделки: не выше доли цены КП"]),
        "band": D(values["Порог «любой способ», ±"]),
    }


def next_art(arts, section):
    if section:
        found = [int(m.group(1)) for art in arts
                 if (m := re.fullmatch(rf"{re.escape(section)}\.N(\d+)", art))]
        return f"{section}.N{max(found, default=0) + 1}"
    found = [int(m.group(1)) for art in arts if (m := re.fullmatch(r"У\.(\d+)", art))]
    return f"У.{max(found, default=0) + 1}"


def normalize(item, arts, names, rates):
    name = str(pick(item, "name", "работа", "название") or "").strip()
    unit = str(pick(item, "unit", "ед", "ед.") or "").strip()
    spec = str(pick(item, "spec", "код", "специальность") or "").strip().upper()
    if not name or not unit:
        die(f"у позиции нет названия или единицы: {item}")
    if spec not in rates:
        die(f"«{name}»: кода {spec or '—'} нет среди специальностей. "
            f"Допустимые: {', '.join(sorted(rates))}")
    key = (name.casefold(), unit.casefold())
    if key in names and not item.get("force"):
        die(f"«{name}», {unit} уже есть в тарификаторе. Если это другая работа — уточните название "
            f"или передайте \"force\": true")
    try:
        norm = D(pick(item, "norm", "норма"))
    except (ValueError, ArithmeticError):
        die(f"«{name}»: не задана норма")
    if norm < 0 or (norm == 0 and spec != "TEX"):
        die(f"«{name}»: норма должна быть больше нуля")
    try:
        market = D(pick(item, "mkt_spb", "рынок спб", "kp", "цена"))
    except (ValueError, ArithmeticError):
        die(f"«{name}»: не задан рынок СПб (mkt_spb)")
    if market <= 0:
        die(f"«{name}»: рынок СПб должен быть больше нуля")
    market_msk = pick(item, "mkt_msk", "рынок мск")
    msk_market_assumed = market_msk in (None, "")
    market_msk = market if msk_market_assumed else D(market_msk)
    if market_msk <= 0:
        die(f"«{name}»: рынок Москвы должен быть больше нуля")
    rate = rates[spec]
    hourly_spb = xround(norm * rate["h_spb"], 0) if norm else Decimal(0)
    hourly_msk = xround(norm * rate["h_msk"], 0) if norm else Decimal(0)
    given_spb = pick(item, "brig_spb", "бригада спб", "brig")
    given_msk = pick(item, "brig_msk", "бригада мск")
    brigade = hourly_spb if given_spb in (None, "") else D(given_spb)
    if given_msk in (None, ""):
        brigade_msk = hourly_msk if spec != "TEX" else brigade
        msk_brig_assumed = True
    else:
        brigade_msk = D(given_msk)
        msk_brig_assumed = False
    if brigade <= 0:
        die(f"«{name}»: оплата бригаде СПб получилась 0. Для техники укажите brig_spb — аренду за единицу")
    if brigade_msk <= 0:
        die(f"«{name}»: оплата бригаде Москвы получилась 0")
    art = str(pick(item, "art", "артикул") or "").strip()
    section = str(pick(item, "section", "раздел") or "").strip().strip(".")
    if art:
        if art in arts:
            die(f"артикул {art} уже занят")
    else:
        art = next_art(arts, section)
    arts.add(art)
    names.add(key)
    src = str(pick(item, "src", "источник", "источник цены") or "").strip()
    norm_src = str(pick(item, "norm_src", "источник нормы") or "").strip()
    if not norm_src:
        die(f"«{name}»: укажите основание нормы (norm_src)")
    if not src:
        die(f"«{name}»: укажите, откуда цена рынка (src)")
    return {
        "art": art, "name": name, "unit": unit, "spec": spec, "section": section,
        "norm": norm, "norm_src": norm_src, "src": src,
        "mkt_spb": market, "mkt_msk": market_msk,
        "brig_spb": brigade, "brig_msk": brigade_msk,
        "hourly_spb": hourly_spb, "hourly_msk": hourly_msk,
        "rate_spb": rate["h_spb"], "rate_msk": rate["h_msk"], "spec_name": rate["name"],
        "brig_given": given_spb not in (None, ""),
        "msk_market_assumed": msk_market_assumed,
        "msk_brig_assumed": msk_brig_assumed,
        "note": str(pick(item, "note", "примечание") or "").strip(),
        "price_spb": price_of(market, hourly_spb, brigade, spec),
        "price_msk": price_of(market_msk, hourly_msk, brigade_msk, spec),
    }


def note_of(item, source):
    bits = [f"{date.today().isoformat()}. Утверждено: {source}"]
    if item["section"]:
        bits.append(f"к разделу {item['section']}, строка дописана в конец, а не внутрь раздела")
    bits.append("оплата СПб задана при утверждении" if item["brig_given"]
                else "оплата СПб = норма × ставка часа")
    if item["msk_brig_assumed"] and item["spec"] != "TEX":
        bits.append("оплата Мск = норма × ставка часа Мск")
    elif item["msk_brig_assumed"]:
        bits.append("оплата Мск не задана, поставлена равной СПб")
    if item["msk_market_assumed"]:
        bits.append("рынок Москвы не задан, поставлен равным СПб — проверить")
    if item["spec"] == "TEX":
        bits.append("техника: запас цены над оплатой 15 %")
    if item["note"]:
        bits.append(item["note"])
    return ". ".join(bits)


def build_rates_row(item, row, norms_row, styles, strings):
    note = note_of(item, item["source"])
    item["note_text"] = note
    price_spb, price_msk = item["price_spb"], item["price_msk"]
    over_spb = price_spb / item["mkt_spb"] - 1
    over_msk = price_msk / item["mkt_msk"] - 1
    cells = [
        c_str("A", row, styles["A"], strings.idx(item["art"])),
        c_str("B", row, styles["B"], strings.idx(item["name"])),
        c_str("C", row, styles["C"], strings.idx(item["unit"])),
        c_str("D", row, styles["D"], strings.idx("да")),
        c_str("E", row, styles["E"], strings.idx(item["spec"])),
        c_num("F", row, styles["F"], item["mkt_spb"]),
        c_num("G", row, styles["G"], item["mkt_msk"]),
        c_num("H", row, styles["H"], price_spb,
              rates_price_formula(row, norms_row, "F", "K", "L", item["spec"])),
        c_num("I", row, styles["I"], price_msk,
              rates_price_formula(row, norms_row, "G", "L", "M", item["spec"])),
        c_num("J", row, styles["J"], over_spb, f'IF(OR(F{row}="",F{row}=0),"",H{row}/F{row}-1)'),
        c_num("K", row, styles["K"], over_msk, f'IF(OR(G{row}="",G{row}=0),"",I{row}/G{row}-1)'),
        c_num("L", row, styles["L"], item["brig_spb"]),
        c_num("M", row, styles["M"], item["brig_msk"]),
        c_num("N", row, styles["N"], (price_spb - item["brig_spb"]) / price_spb,
              f'IF(OR(H{row}="",H{row}=0),"",(H{row}-L{row})/H{row})'),
        c_num("O", row, styles["O"], (price_msk - item["brig_msk"]) / price_msk,
              f'IF(OR(I{row}="",I{row}=0),"",(I{row}-M{row})/I{row})'),
        c_num("P", row, styles["P"], price_spb - item["brig_spb"], f'IF(H{row}="","",H{row}-L{row})'),
        c_num("Q", row, styles["Q"], price_msk - item["brig_msk"], f'IF(I{row}="","",I{row}-M{row})'),
        c_num("R", row, styles["R"], None),
        c_num("S", row, styles["S"], None),
        c_inline("T", row, styles["T"], f'IF(OR(S{row}="",F{row}="",F{row}=0),"",S{row}/F{row})', ""),
        c_str("U", row, styles["U"], strings.idx(item["src"])),
        c_str("V", row, styles["V"], strings.idx(note)),
    ]
    return f'<row r="{row}" spans="1:22" x14ac:dyDescent="0.25">{"".join(cells)}</row>'


def build_norms_row(item, row, rates_row, styles, strings, s_last, threshold, day_hours):
    formulas = norms_formulas(row, rates_row, s_last)
    hourly_spb, hourly_msk = item["hourly_spb"], item["hourly_msk"]
    per_hour = xround(day_hours / item["norm"], 2) if item["norm"] else None
    gap_spb = gap(hourly_spb, item["brig_spb"]) if hourly_spb else None
    gap_msk = gap(hourly_msk, item["brig_msk"]) if hourly_msk else None
    text_spb = better_text(gap_spb, threshold)
    text_msk = better_text(gap_msk, threshold)
    market_gap_spb = gap(item["mkt_spb"], hourly_spb) if hourly_spb else None
    market_gap_msk = gap(item["mkt_msk"], hourly_msk) if hourly_msk else None
    price_gap_spb = gap(item["price_spb"], hourly_spb) if hourly_spb else None
    price_gap_msk = gap(item["price_msk"], hourly_msk) if hourly_msk else None
    margin_spb = (item["price_spb"] - item["brig_spb"]) / item["price_spb"]
    margin_msk = (item["price_msk"] - item["brig_msk"]) / item["price_msk"]
    values = {
        "H": per_hour, "I": item["rate_spb"], "J": item["rate_msk"],
        "K": hourly_spb or None, "L": hourly_msk or None,
        "M": item["brig_spb"], "N": item["brig_msk"],
        "O": gap_spb, "P": gap_msk,
        "S": xround(item["brig_spb"] / item["norm"], 0) if item["norm"] else None,
        "T": xround(item["brig_msk"] / item["norm"], 0) if item["norm"] else None,
        "U": item["mkt_spb"], "V": item["mkt_msk"],
        "W": market_gap_spb, "X": market_gap_msk,
        "Y": item["price_spb"], "Z": item["price_msk"],
        "AA": item["price_spb"] / item["mkt_spb"] - 1,
        "AB": item["price_msk"] / item["mkt_msk"] - 1,
        "AC": price_gap_spb, "AD": price_gap_msk,
        "AE": margin_spb, "AF": margin_msk,
        "AG": (item["price_spb"] - hourly_spb) if hourly_spb else None,
        "AH": (item["price_msk"] - hourly_msk) if hourly_msk else None,
    }
    string_cols = {"E", "Q", "R"}
    cells = [
        c_str("A", row, styles["A"], strings.idx(item["art"])),
        c_str("B", row, styles["B"], strings.idx(item["name"])),
        c_str("C", row, styles["C"], strings.idx(item["unit"])),
        c_str("D", row, styles["D"], strings.idx(item["spec"])),
        c_inline("E", row, styles["E"], formulas["E"], item["spec_name"] or ""),
        c_num("F", row, styles["F"], item["norm"] if item["norm"] else 0),
        c_str("G", row, styles["G"], strings.idx(item["norm_src"])),
    ]
    for col in col_letters(34)[7:]:
        formula = formulas[col]
        if col in string_cols:
            text = text_spb if col == "Q" else text_msk
            cells.append(c_inline(col, row, styles[col], formula, text))
        else:
            cells.append(c_num(col, row, styles[col], values[col], formula))
    return f'<row r="{row}" spans="1:34" x14ac:dyDescent="0.25">{"".join(cells)}</row>'


def shift_staff_row(row_xml, old, new):
    row_xml = re.sub(rf'(<row r="){old}(")', rf"\g<1>{new}\2", row_xml, count=1)
    row_xml = re.sub(rf'(r="[A-Z]{{1,3}}){old}(")', rf"\g<1>{new}\2", row_xml)

    def fix_formula(match):
        return match.group(1) + re.sub(rf"(?<!\d){old}(?!\d)", str(new), match.group(2)) + match.group(3)

    return re.sub(r"(<f[^>]*>)(.*?)(</f>)", fix_formula, row_xml)


def replace_cell(row_xml, col, row, cell):
    updated, count = re.subn(
        rf'<c r="{col}{row}"(?: [^>]*?/>|[^>]*>.*?</c>)', cell, row_xml, count=1, flags=re.S)
    if count != 1:
        die(f"в строке сотрудника нет ячейки {col}{row}")
    return updated


def set_cache(row_xml, col, row, value, text=False):
    cached = f"<v>{xml_text(value) if text else num(value)}</v>"

    def repl(match):
        body = re.sub(r"<v[^>]*/>|<v>.*?</v>", cached, match.group(2), count=1)
        if "<v" not in body:
            body += cached
        return match.group(1) + body + match.group(3)

    updated, count = re.subn(rf'(<c r="{col}{row}"[^>]*>)(.*?)(</c>)', repl, row_xml, count=1, flags=re.S)
    if count != 1:
        die(f"не обновился кэш {col}{row}")
    return updated


def deal_ceiling(brigade, price, uplift, cap):
    grown_raw = brigade * (1 + uplift)
    cap_raw = price * cap
    grown = xround(grown_raw, 0 if grown_raw < 1000 else -1)
    limited = xround(cap_raw, 0 if cap_raw < 1000 else -1)
    return max(brigade, min(grown, limited))


def how_to_pay(spec, hourly, brigade, band):
    if spec == "TEX":
        return "аренда техники — сделка/счёт"
    if hourly in (None, "") or brigade in (None, 0):
        return "—"
    if hourly < brigade * (1 - band):
        return "почасовка"
    if hourly > brigade * (1 + band):
        return "сделка"
    return "любой способ"


def expand_shared(row_xml, col, row_num, formula, cache):
    pattern = rf'<c r="{col}{row_num}"([^>]*)><f t="shared" si="\d+"/>.*?</c>'
    if not re.search(pattern, row_xml, flags=re.S):
        return row_xml

    def repl(match):
        cached = f"<v>{num(cache)}</v>" if cache != "" else "<v/>"
        return f'<c r="{col}{row_num}"{match.group(1)}><f>{xml_text(formula)}</f>{cached}</c>'

    updated, count = re.subn(pattern, repl, row_xml, count=1, flags=re.S)
    if count != 1:
        die(f"не раскрылась общая формула {col}{row_num} в файле сотрудников")
    return updated


def build_staff_row(template_xml, old, new, item, strings, params, staff_rate):
    row = shift_staff_row(template_xml, old, new)
    styles = row_styles(row, new)
    norm = item["norm"]
    hourly = xround(norm * staff_rate["h_spb"], 0) if norm else Decimal(0)
    row = replace_cell(row, "A", new, c_str("A", new, styles["A"], strings.idx(item["art"])))
    row = replace_cell(row, "B", new, c_str("B", new, styles["B"], strings.idx(item["name"])))
    row = replace_cell(row, "C", new, c_str("C", new, styles["C"], strings.idx(item["unit"])))
    row = replace_cell(row, "D", new, c_str("D", new, styles["D"], strings.idx(item["spec"])))
    row = replace_cell(row, "F", new, c_num("F", new, styles["F"], item["price_spb"]))
    row = replace_cell(row, "G", new, c_num("G", new, styles["G"], item["price_msk"]))
    row = replace_cell(row, "H", new, c_num("H", new, styles["H"], item["brig_spb"]))
    row = replace_cell(row, "J", new, c_num("J", new, styles["J"], item["brig_msk"]))
    row = replace_cell(row, "L", new, c_num("L", new, styles["L"], norm if norm else 0))
    row = replace_cell(row, "S", new, c_num("S", new, styles.get("S", ""), item["price_spb"] / item["mkt_spb"] - 1))
    row = replace_cell(row, "T", new, c_num("T", new, styles.get("T", ""), item["price_msk"] / item["mkt_msk"] - 1))
    row = expand_shared(row, "M", new, f'IF(OR(L{new}="",L{new}=0),"",1/L{new})', "")
    row = expand_shared(row, "Q", new, f'IF(OR(L{new}="",O{new}=""),"",ROUND(L{new}*O{new},0))', "")
    row = set_cache(row, "E", new, staff_rate["name"] or "", text=True)
    row = set_cache(row, "I", new, deal_ceiling(item["brig_spb"], item["price_spb"], params["uplift"], params["cap"]))
    row = set_cache(row, "K", new, deal_ceiling(item["brig_msk"], item["price_msk"], params["uplift"], params["cap"]))
    row = set_cache(row, "M", new, (1 / norm) if norm else "")
    row = set_cache(row, "N", new, (params["shift"] / norm) if norm else "")
    indexed = staff_rate.get("indexed", True)
    row = set_cache(row, "O", new, staff_rate["h_spb"] if indexed else "")
    row = set_cache(row, "P", new, staff_rate["h_msk"] if indexed else "")
    row = set_cache(row, "Q", new, (hourly if norm else 0) if staff_rate.get("indexed", True) else "")
    row = set_cache(row, "R", new, how_to_pay(item["spec"], hourly, item["brig_spb"], params["band"]), text=True)
    return row


def build_materials_row(row, styles, strings, item):
    cells = [
        c_str("A", row, styles.get("A", ""), strings.idx(item["art"])),
        c_str("B", row, styles.get("B", ""), strings.idx(item["name"])),
        c_str("C", row, styles.get("C", ""), strings.idx(item["unit"])),
        c_str("D", row, styles.get("D", ""), strings.idx(item["spec"])),
        c_num("E", row, styles.get("E", ""), item["price_spb"]),
        c_num("F", row, styles.get("F", ""), item["price_msk"]),
        c_num("G", row, styles.get("G", ""), item["brig_spb"]),
        c_num("H", row, styles.get("H", ""), item["brig_msk"]),
        c_num("I", row, styles.get("I", ""), item["price_spb"] / item["mkt_spb"] - 1),
        c_num("J", row, styles.get("J", ""), item["price_msk"] / item["mkt_msk"] - 1),
    ]
    return f'<row r="{row}" spans="1:10" x14ac:dyDescent="0.25">{"".join(cells)}</row>'


def insert_rows(xml, rows_xml):
    return xml.replace("</sheetData>", "".join(rows_xml) + "</sheetData>", 1)


def save_book(path, replacements):
    path = Path(path)
    temporary = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = replacements.get(info.filename, source.read(info.filename))
            target.writestr(info.filename, data)
    temporary.replace(path)


def clear_cache():
    if CACHE.exists():
        for cached in CACHE.glob("spr_*.pkl"):
            cached.unlink()


def prepare(payload, tarif_path):
    if payload.get("ready") is False:
        die("это черновик (\"ready\": false). Проверьте цифры, уберите ключ и запустите запись ещё раз")
    source = str(payload.get("source") or "").strip()
    if not source:
        die("укажите source: лот и кто утвердил цифры")
    items = payload.get("items")
    if not isinstance(items, list) or not items:
        die("в файле нет списка items")
    arts, names, last_section = existing(tarif_path)
    rates = load_rates(tarif_path, "Нормо-час по специальностям")
    prepared = []
    for item in items:
        if not isinstance(item, dict):
            die(f"позиция должна быть объектом, а не {item!r}")
        ready = normalize(item, arts, names, rates)
        ready["source"] = source
        prepared.append(ready)
    return prepared, last_section, rates


def apply(payload, tarif_path=TARIF, staff_path=STAFF, materials_path=MATERIALS, dry_run=False):
    prepared, last_section, rates = prepare(payload, tarif_path)
    staff_rates = load_rates(staff_path, "Почасовка — ставки и план", min_row=5)
    for code, rate in staff_rates.items():
        rate["indexed"] = True
    missing = [item["spec"] for item in prepared if item["spec"] not in staff_rates and item["spec"] != "TEX"]
    if missing:
        die("этих кодов нет в почасовке для сотрудников: " + ", ".join(sorted(set(missing))))
    staff_rates.setdefault("TEX", {
        "name": "Аренда техники / механизмы", "h_spb": D(0), "h_msk": D(0), "indexed": False,
    })
    for item in prepared:
        item["spec_name"] = item["spec_name"] or staff_rates[item["spec"]]["name"]
        own = rates[item["spec"]]
        copied = staff_rates[item["spec"]]
        if own["h_spb"] != copied["h_spb"] or own["h_msk"] != copied["h_msk"]:
            print(f"  {item['spec']}: ставка в тарификаторе {num(own['h_spb'])}/{num(own['h_msk'])}, "
                  f"в почасовке сотрудников {num(copied['h_spb'])}/{num(copied['h_msk'])}. "
                  f"Цена считается по ставке тарификатора.")
    params = staff_params(staff_path)
    need_header = last_section != APPROVED
    staff_header = staff_last_section(staff_path) != APPROVED

    with zipfile.ZipFile(tarif_path) as book:
        tarif_wb = book.read("xl/workbook.xml").decode()
        rels = book.read("xl/_rels/workbook.xml.rels").decode()
        rates_name = sheet_target(tarif_wb, rels, "Расценки")
        norms_name = sheet_target(tarif_wb, rels, "Нормы и нормо-часы")
        rates_xml = book.read(rates_name).decode()
        norms_xml = book.read(norms_name).decode()
        strings = Strings(book.read("xl/sharedStrings.xml").decode())
    rates_last, norms_last = dimension_last(rates_xml), dimension_last(norms_xml)
    if last_row_tag(rates_xml) != rates_last or last_row_tag(norms_xml) != norms_last:
        die("последняя строка тарификатора не совпала с dimension, файл лучше открыть и пересохранить")
    if norms_last != rates_last + 1:
        die(f"сбита пара строк: расценки кончаются на {rates_last}, нормы на {norms_last}. "
            f"Ожидалось, что нормы ровно на одну строку длиннее. В середину ничего не вставлялось")
    s_last = re.search(r"Нормо-час по специальностям'!\$[A-Z]\$6:\$[A-Z]\$(\d+)", norms_xml)
    if not s_last:
        die("не нашёл диапазон листа специальностей в формулах норм")
    s_last = s_last.group(1)
    threshold = D(re.search(r'<c r="C3"[^>]*>.*?<v>([^<]+)</v>', norms_xml).group(1))
    day_hours = D(re.search(r'<c r="G3"[^>]*>.*?<v>([^<]+)</v>', norms_xml).group(1))
    rates_style = row_styles(rates_xml, rates_last)
    norms_style = row_styles(norms_xml, norms_last)
    rates_header_style = row_styles(rates_xml, 5)
    norms_header_style = row_styles(norms_xml, 6)

    rates_rows, norms_rows = [], []
    rates_row, norms_row = rates_last, norms_last
    if need_header:
        rates_row += 1
        norms_row += 1
        rates_rows.append(header_row(rates_row, rates_header_style, APPROVED, strings))
        norms_rows.append(header_row(norms_row, norms_header_style, APPROVED, strings))
    for item in prepared:
        rates_row += 1
        norms_row += 1
        if norms_row != rates_row + 1:
            die("внутренняя ошибка: пара строк норм и расценок разошлась")
        rates_rows.append(build_rates_row(item, rates_row, norms_row, rates_style, strings))
        norms_rows.append(build_norms_row(
            item, norms_row, rates_row, norms_style, strings, s_last, threshold, day_hours))

    with zipfile.ZipFile(staff_path) as book:
        staff_wb = book.read("xl/workbook.xml").decode()
        rels = book.read("xl/_rels/workbook.xml.rels").decode()
        staff_sheet = sheet_target(staff_wb, rels, "Расценки")
        hour_sheet = sheet_target(staff_wb, rels, "Почасовка — ставки и план")
        staff_xml = book.read(staff_sheet).decode()
        hour_xml = book.read(hour_sheet).decode()
        staff_strings = Strings(book.read("xl/sharedStrings.xml").decode())
    staff_last = dimension_last(staff_xml)
    if last_row_tag(staff_xml) != staff_last:
        die("последняя строка файла сотрудников не совпала с dimension")
    staff_template = extract_row(staff_xml, staff_last)
    staff_header_style = row_styles(staff_xml, 5)
    staff_rows = []
    staff_row = staff_last
    if staff_header:
        staff_row += 1
        staff_rows.append(header_row(staff_row, staff_header_style, APPROVED, staff_strings))
    for item in prepared:
        staff_row += 1
        staff_rows.append(build_staff_row(
            staff_template, staff_last, staff_row, item, staff_strings, params, staff_rates[item["spec"]]))

    with zipfile.ZipFile(materials_path) as book:
        mat_wb = book.read("xl/workbook.xml").decode()
        rels = book.read("xl/_rels/workbook.xml.rels").decode()
        mat_sheet = sheet_target(mat_wb, rels, "Тарификатор")
        calc_sheet = sheet_target(mat_wb, rels, "Калькулятор")
        mat_xml = book.read(mat_sheet).decode()
        calc_xml = book.read(calc_sheet).decode()
        mat_strings = Strings(book.read("xl/sharedStrings.xml").decode())
    mat_last = dimension_last(mat_xml)
    mat_style = row_styles(mat_xml, mat_last)
    mat_rows = []
    mat_row = mat_last
    for item in prepared:
        mat_row += 1
        mat_rows.append(build_materials_row(mat_row, mat_style, mat_strings, item))

    report = []
    for item in prepared:
        report.append(item)
        print(f"{item['art']} | {item['name']} | {item['unit']} | {item['spec']} | "
              f"норма {num(item['norm'])} | бригада {num(item['brig_spb'])}/{num(item['brig_msk'])} | "
              f"рынок {num(item['mkt_spb'])}/{num(item['mkt_msk'])} | "
              f"цена НЕОСТРОЙ {num(item['price_spb'])}/{num(item['price_msk'])}")
        if item["msk_market_assumed"] or (item["msk_brig_assumed"] and item["spec"] == "TEX"):
            print("  Москва не была задана и поставлена по правилу выше — её стоит проверить")
    if dry_run:
        print(f"Проверка прошла, запись не делалась. Строк к добавлению: {len(prepared)}. "
              f"Блок «{APPROVED}» {'будет открыт' if need_header else 'уже есть'}.")
        return report

    rates_xml = extend_tail(insert_rows(rates_xml, rates_rows), rates_last, rates_row)
    norms_xml = extend_tail(insert_rows(norms_xml, norms_rows), norms_last, norms_row)
    tarif_wb = must_replace(tarif_wb, f"Расценки!$A$4:$V${rates_last}", f"Расценки!$A$4:$V${rates_row}", "фильтр расценок")
    tarif_wb = must_replace(
        tarif_wb, f"'Нормы и нормо-часы'!$A$5:$AH${norms_last}",
        f"'Нормы и нормо-часы'!$A$5:$AH${norms_row}", "фильтр норм")
    save_book(tarif_path, {
        rates_name: rates_xml.encode(), norms_name: norms_xml.encode(),
        "xl/sharedStrings.xml": strings.render().encode(), "xl/workbook.xml": tarif_wb.encode(),
    })

    staff_xml = extend_tail(insert_rows(staff_xml, staff_rows), staff_last, staff_row)
    hour_xml = must_replace(
        hour_xml, f"Расценки!$D$5:$D${staff_last}", f"Расценки!$D$5:$D${staff_row}", "счётчик специальностей")
    staff_wb = must_replace(
        staff_wb, f"Расценки!$A$4:$R${staff_last}", f"Расценки!$A$4:$R${staff_row}", "фильтр сотрудников")
    save_book(staff_path, {
        staff_sheet: staff_xml.encode(), hour_sheet: hour_xml.encode(),
        "xl/sharedStrings.xml": staff_strings.render().encode(), "xl/workbook.xml": staff_wb.encode(),
    })

    mat_xml = extend_tail(insert_rows(mat_xml, mat_rows), mat_last, mat_row)
    calc_xml = re.sub(
        rf"(\$[A-J]\$5:\$[A-J]\$){mat_last}\b", lambda match: match.group(1) + str(mat_row), calc_xml)
    mat_wb = must_replace(
        mat_wb, f"Тарификатор!$A$4:$J${mat_last}", f"Тарификатор!$A$4:$J${mat_row}", "фильтр калькулятора")
    save_book(materials_path, {
        mat_sheet: mat_xml.encode(), calc_sheet: calc_xml.encode(),
        "xl/sharedStrings.xml": mat_strings.render().encode(), "xl/workbook.xml": mat_wb.encode(),
    })
    clear_cache()
    print(f"Дописано позиций: {len(prepared)}. Они в конце блока «{APPROVED}». "
          f"Следующий расчёт увидит их после попадания справочника в рабочую ветку.")
    return report


def draft(calc_path, out_path):
    data = json.loads(Path(calc_path).read_text(encoding="utf-8"))
    works = data.get("works") or []
    items = []
    for work in works:
        if work.get("art"):
            continue
        item = {
            "name": work.get("name"),
            "unit": work.get("unit"),
            "spec": work.get("spec"),
            "norm": work.get("norm"),
            "norm_src": work.get("norm_src") or "",
            "mkt_spb": work.get("kp"),
            "src": work.get("src") or "",
            "note": "черновик из расчёта: рынок и оплату сверить перед записью",
        }
        if work.get("brig"):
            item["brig_spb"] = work["brig"]
        items.append(item)
    payload = {
        "ready": False,
        "source": data.get("project") or Path(calc_path).parent.name,
        "items": items,
    }
    destination = Path(out_path)
    destination.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Черновик, справочник не изменён. Позиций: {len(items)}. Файл: {destination}")
    print("Рынок СПб взят из цены КП расчёта. Рынок Москвы пустой: при записи он станет равным СПб, с пометкой.")
    print("Оплату, которая должна отличаться от нормы × ставка, впишите в brig_spb и brig_msk.")
    print('После проверки уберите "ready": false и выполните python tools/tarif_add.py "<файл>"')
    return destination


def main(argv):
    args = list(argv)
    dry = "--dry-run" in args
    if dry:
        args.remove("--dry-run")
    if args and args[0] == "--draft":
        if len(args) < 2:
            die("укажите Расчёт.json: python tools/tarif_add.py --draft outbox/<лот>/Расчёт.json")
        out = None
        if "-o" in args:
            out = args[args.index("-o") + 1]
        if out is None:
            out = str(Path(args[1]).with_name("В тарификатор.json"))
        draft(args[1], out)
        return
    if len(args) != 1:
        die("так: python tools/tarif_add.py <файл.json> [--dry-run]\n"
            "или: python tools/tarif_add.py --draft outbox/<лот>/Расчёт.json [-o черновик.json]")
    payload = json.loads(Path(args[0]).read_text(encoding="utf-8"))
    if isinstance(payload, list):
        payload = {"source": "", "items": payload}
    apply(payload, dry_run=dry)


if __name__ == "__main__":
    main(sys.argv[1:])
