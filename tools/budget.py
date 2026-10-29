"""Собрать бюджет, КП, карту и строку реестра из короткого Расчёт.json.

    python tools/budget.py outbox/<лот>/Расчёт.json

Агент заполняет только объёмы, артикулы и цены закупки. Ставки, отходы, наценки
и расходники подставляются из справочников. В книге бюджета итоги и связи между
листами — формулы Excel: правка объёма или расценки пересчитывает файл.
Заказчику уходит только КП: себестоимости и оплаты бригад в нём нет.

Схема Расчёт.json — в tools/README.md.
"""
import json
import sys
import warnings
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tarif
from paths import COMMON

warnings.filterwarnings("ignore")

NAVY = "1F4E79"
PALE = "D6E3F0"
YELLOW = "FFF2CC"
GREEN = "E2EFDA"
ORANGE = "FCE4D6"
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)
VAT = 1.22


def region_key(name):
    n = str(name or "").lower()
    if "моск" in n:
        return "msk"
    return "spb"


def delivery_of(name):
    n = str(name or "").lower()
    if "ленобл" in n or "ленинград" in n or "подмосков" in n:
        return 0.05
    if "петербург" in n or n.strip() in ("спб", "мск") or "москва" in n or "москв" in n:
        return 0.03
    if "спб" in n:
        return 0.03
    return 0.10


def load_refs():
    d = tarif.load()
    cats = {c["cat"]: c for c in d["cats"]}
    types = d["types"]
    return d, cats, types


def cat_of(cats, name):
    if name in cats:
        return cats[name]
    for c in cats.values():
        if name and name[:12].lower() in c["cat"].lower():
            return c
    return cats.get("Прочее") or {"waste": 0.03, "mk_sep": 0.15, "mk_turnkey": 0.25}


def tender_of(types, name):
    for t in types:
        if name and name[:18].lower() in t["type"].lower():
            return t
    for t in types:
        if name and name.lower() in t["type"].lower():
            return t
    return {"type": name, "mk_from": 0.15, "mk_to": 0.3}


def requisites():
    path = COMMON / "Реквизиты - НЕОСТРОЙ ООО.docx"
    if not path.exists():
        return {}
    doc = Document(path)
    out = {}
    if doc.tables:
        for row in doc.tables[0].rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) >= 2 and cells[0]:
                out[cells[0].rstrip(":")] = cells[1]
    return out


def money(cell):
    cell.number_format = '#,##0.00'
    cell.border = THIN


def head(ws, row, titles, fill=NAVY):
    for i, t in enumerate(titles, 1):
        c = ws.cell(row, i, t)
        c.font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = THIN
    ws.row_dimensions[row].height = 32
    ws.auto_filter.ref = None
    ws.freeze_panes = "A" + str(row + 1)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_view.showGridLines = False
    ws.oddHeader.left.text = ws.title


def widths(ws, wsizes):
    for i, w in enumerate(wsizes, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def safe(name):
    return "".join(c for c in name if c.isalnum() or c in " _-").strip()[:60] or "Объект"


def build(data, out_dir):
    spr, cats, types = load_refs()
    reg = region_key(data.get("price_region") or data.get("region"))
    reg_name = "Мск" if reg == "msk" else "СПб"
    if not data.get("price_region") and region_key(data.get("region")) == "spb" and delivery_of(data.get("region")) == 0.10:
        print(f"регион «{data.get('region')}»: цены тарификатора СПб, доставка как для региона 10 %")
    delivery = delivery_of(data.get("region"))
    months = float(data.get("months_to_purchase") or 0)
    growth = 0.0125 * max(0, months - 2)
    finance = 0 if data.get("advance_covers_materials", True) else 0.02 * float(data.get("finance_months") or 0)
    turnkey = str(data.get("form", "separate")).lower() in ("turnkey", "под ключ", "ключ")
    tender = tender_of(types, data.get("tender_type", ""))
    rash = spr["rash"]

    works = []
    for src in data.get("works", []):
        w = dict(src)
        tw = spr["works"].get(str(src.get("art") or ""))
        if tw:
            w.update(name=w.get("name") or tw["name"], unit=w.get("unit") or tw["unit"],
                     spec=tw["spec"], kp=tw[f"kp_{reg}"], brig=tw[f"brig_{reg}"],
                     norm=tw.get("norm") or 0, over=tw[f"over_{reg}"] or 0, sec=tw["sec"])
        else:
            w.setdefault("kp", 0)
            w.setdefault("brig", 0)
            w.setdefault("norm", 0)
            w.setdefault("over", 0)
            w.setdefault("spec", "")
            print(f"нет в тарификаторе: {w.get('name')} — цена КП {w['kp']}, бригада {w['brig']} (должно быть в JSON)")
        w["qty"] = float(w.get("qty") or 0)
        w["rash_pct"] = rash.get(w["spec"], {}).get("pct", 0)
        works.append(w)

    mats = []
    for src in data.get("materials", []):
        m = dict(src)
        cat = cat_of(cats, m.get("cat") or "Прочее")
        m["waste"] = 0 if m.get("service") else cat["waste"]
        pulled = False
        if m.get("work"):
            for w in works:
                if str(m["work"]) in (str(w.get("art")), w.get("name")) and w["over"] > 0.005:
                    pulled = True
        sep = m.get("service") or pulled or not turnkey
        m["markup"] = 0.15 if m.get("service") else (cat["mk_sep"] if sep else cat["mk_turnkey"])
        m["qty"] = float(m.get("qty") or 0)
        gross = float(m.get("price_vat") or 0)
        m["buy_net"] = round(gross / VAT, 4) if m.get("vat_included", True) else round(gross, 4)
        retail = m.get("retail_vat")
        m["retail_net"] = round(float(retail) / VAT, 4) if retail else None
        m["ss_unit"] = m["buy_net"] * (1 + m["waste"]) * (1 + delivery) * (1 + growth + finance)
        raw = m["ss_unit"] * (1 + m["markup"])
        m["capped"] = bool(m["retail_net"] and raw > m["retail_net"] + 0.01)
        m["kp_unit"] = min(raw, m["retail_net"]) if m["capped"] else raw
        mats.append(m)

    for w in works:
        w["mat_ss"] = sum(m["ss_unit"] * m["qty"] for m in mats
                          if str(m.get("work")) in (str(w.get("art")), w.get("name")))
        w["brig_sum"] = (w["brig"] or 0) * w["qty"]
        w["rash_sum"] = w["brig_sum"] * w["rash_pct"]
        w["direct"] = w["brig_sum"] + w["rash_sum"] + w["mat_ss"]
    indirect = data.get("indirect", [])
    indirect_sum = sum(float(i.get("sum") or 0) for i in indirect)
    direct_sum = sum(w["direct"] for w in works) or 1
    for w in works:
        w["alloc"] = indirect_sum * w["direct"] / direct_sum if works else 0
        w["kp_sum"] = (w["kp"] or 0) * w["qty"] + w["alloc"]
        w["kp_unit_full"] = w["kp_sum"] / w["qty"] if w["qty"] else 0

    mat_ss = sum(m["ss_unit"] * m["qty"] for m in mats)
    mat_kp = sum(m["kp_unit"] * m["qty"] for m in mats)
    brig = sum(w["brig_sum"] for w in works)
    rash_sum = sum(w["rash_sum"] for w in works)
    work_kp = sum(w["kp_sum"] for w in works)
    cost = mat_ss + brig + rash_sum + indirect_sum
    kp = work_kp + mat_kp
    markup = kp / cost - 1 if cost else 0
    hours = sum((w["norm"] or 0) * w["qty"] for w in works)

    proj = safe(data.get("project") or "Объект")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    budget_path = out_dir / f"Бюджет себестоимости_НЕОСТРОЙ_{proj}.xlsx"
    kp_path = out_dir / f"КП_НЕОСТРОЙ_{proj}.xlsx"
    map_path = out_dir / f"Карта реализации_НЕОСТРОЙ_{proj}.xlsx"
    letter_path = out_dir / f"Письмо-КП_НЕОСТРОЙ_{proj}.docx"

    write_budget(budget_path, data, works, mats, indirect, reg_name, delivery, growth, finance,
                 turnkey, tender, mat_ss, mat_kp, brig, rash_sum, indirect_sum, cost, kp, markup, hours)
    write_kp(kp_path, data, works, mats, kp, turnkey)
    write_map(map_path, data, works, hours, kp)
    write_letter(letter_path, data, kp)
    write_registry(out_dir, data, cost, kp, mat_ss, mat_kp, brig, work_kp)
    print(f"с/с {cost:,.0f} | КП б/НДС {kp:,.0f} | с НДС {kp * VAT:,.0f} | наценка {markup:.1%}")
    lo, hi = tender.get("mk_from") or 0, tender.get("mk_to") or 1
    if markup < 0.15:
        print("наценка ниже 15 % — нужно согласование ГД")
    elif markup < lo or markup > hi:
        print(f"наценка вне диапазона типа тендера {lo:.0%}–{hi:.0%}")
    print(budget_path)
    print(kp_path)
    print(map_path)
    print(letter_path)


def title(ws, text):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14, color=NAVY, name="Calibri")
    ws.row_dimensions[1].height = 22
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_view.showGridLines = False
    ws.oddFooter.right.text = "НЕОСТРОЙ  |  стр. &P из &N"
    ws.oddFooter.left.text = "Внутренний документ — заказчику не передавать"
    ws.sheet_view.zoomScale = 110


def write_budget(path, data, works, mats, indirect, reg_name, delivery, growth, finance,
                 turnkey, tender, mat_ss, mat_kp, brig, rash_sum, indirect_sum, cost, kp, markup, hours):
    """Все производные ячейки — формулы. Числа только во вводе, правка пересчитывает файл."""
    del growth, finance, mat_ss, mat_kp, brig, rash_sum, indirect_sum, cost, kp, markup, hours
    wb = openpyxl.Workbook()
    wb.calculation.calcMode = "auto"

    par = wb.active
    par.title = "Параметры"
    title(par, "Параметры расчёта — правка этих ячеек пересчитывает весь файл")
    head(par, 3, ["Параметр", "Значение"])
    months = float(data.get("months_to_purchase") or 0)
    fin_months = float(data.get("finance_months") or 0)
    advance = 1 if data.get("advance_covers_materials", True) else 0
    inputs = [
        (4, "Доставка", delivery, "0%"),
        (5, "Месяцев до закупки", months, "0.00"),
        (7, "Месяцев финансирования", fin_months, "0.00"),
        (8, "Материалы из аванса (1 — да)", advance, "0"),
        (10, "Ставка НДС", 0.22, "0%"),
    ]
    for row, name, val, fmt in inputs:
        par.cell(row, 1, name).border = THIN
        c = par.cell(row, 2, val)
        money(c)
        c.number_format = fmt
    derived = [
        (6, "Резерв роста", "=0.0125*MAX(0,B5-2)", "0.00%"),
        (9, "Стоимость денег", "=IF(B8=1,0,0.02*B7)", "0.00%"),
        (11, "Делитель НДС", "=1+B10", "0.00"),
        (12, "Коэффициент закупки", "=(1+B4)*(1+B6+B9)", "0.0000"),
    ]
    for row, name, formula, fmt in derived:
        par.cell(row, 1, name).border = THIN
        c = par.cell(row, 2, formula)
        money(c)
        c.number_format = fmt
    nmck = data.get("nmck")
    par.cell(13, 1, "НМЦК б/НДС").border = THIN
    c = par.cell(13, 2, float(nmck) if nmck else None)
    money(c)
    par["A15"] = (f"Регион цен {reg_name}. Форма КП: {'под ключ' if turnkey else 'раздельная'}. "
                  "Жёлтым ничего не помечено: менять можно белые ячейки ввода.")
    widths(par, [42, 18])

    mm = wb.create_sheet("Материалы", 0)
    title(mm, "Материалы")
    head(mm, 3, ["К работе", "Материал", "Категория", "Ед.", "Кол-во", "Закупка с НДС", "НДС в цене",
                 "Закупка б/НДС", "Отходы", "С/с ед.", "Наценка", "Цена до потолка", "Розница б/НДС",
                 "Цена КП ед.", "С/с итого", "КП итого", "Выше розницы", "Магазин", "Ссылка", "Дата", "Примечание"])
    mend = 3 + max(len(mats), 1)
    for i, m in enumerate(mats):
        r = 4 + i
        retail = m.get("retail_net")
        vals = [m.get("work"), m.get("name"), m.get("cat"), m.get("unit"), m["qty"],
                m.get("price_vat") or 0, 1 if m.get("vat_included", True) else 0,
                m["waste"], m["markup"], retail if retail else None,
                m.get("shop"), m.get("url"), m.get("date"), m.get("note")]
        # E qty, F price, G vat flag, I waste, K markup, M retail — ввод. Остальное формулы.
        mm.cell(r, 1, vals[0])
        mm.cell(r, 2, vals[1])
        mm.cell(r, 3, vals[2])
        mm.cell(r, 4, vals[3])
        mm.cell(r, 5, vals[4])
        mm.cell(r, 6, vals[5])
        mm.cell(r, 7, vals[6])
        mm.cell(r, 8, f"=IF(G{r}=1,F{r}/Параметры!$B$11,F{r})")
        mm.cell(r, 9, vals[7])
        mm.cell(r, 10, f"=H{r}*(1+I{r})*Параметры!$B$12")
        mm.cell(r, 11, vals[8])
        mm.cell(r, 12, f"=J{r}*(1+K{r})")
        mm.cell(r, 13, vals[9])
        mm.cell(r, 14, f'=IF(OR(M{r}="",M{r}=0),L{r},MIN(L{r},M{r}))')
        mm.cell(r, 15, f"=J{r}*E{r}")
        mm.cell(r, 16, f"=N{r}*E{r}")
        mm.cell(r, 17, f'=IF(AND(M{r}<>"",M{r}<>0,L{r}>M{r}+0.01),"ДА — снижено до розницы","")')
        mm.cell(r, 18, vals[10])
        mm.cell(r, 19, vals[11])
        mm.cell(r, 20, vals[12])
        mm.cell(r, 21, vals[13])
        for c in range(1, 22):
            cell = mm.cell(r, c)
            cell.border = THIN
            if c in (5, 6, 8, 10, 12, 13, 14, 15, 16):
                cell.number_format = '#,##0.00'
            if c in (9, 11):
                cell.number_format = "0%"
            if c == 7:
                cell.number_format = "0"
        mm.row_dimensions[r].height = 28
    mm.conditional_formatting.add(
        f"Q4:Q{mend}",
        FormulaRule(formula=['$Q4<>""'], fill=PatternFill("solid", fgColor=ORANGE)),
    )
    widths(mm, [14, 36, 24, 8, 10, 14, 12, 14, 10, 12, 10, 14, 14, 12, 14, 14, 16, 16, 28, 12, 24])

    ww = wb.create_sheet("Работы", 0)
    title(ww, "Работы")
    head(ww, 3, ["Артикул", "Позиция ВОР", "Ед.", "Объём", "Метка", "Код", "КП тарифа", "Бригада",
                 "Норма, чел.-ч", "% расходников", "Подтянута", "Оплата бригаде", "Расходники",
                 "С/с материалов", "Прямая с/с", "Косвенные", "КП работ", "КП за ед."])
    wend = 3 + max(len(works), 1)
    for i, w in enumerate(works):
        r = 4 + i
        ww.cell(r, 1, w.get("art"))
        ww.cell(r, 2, w.get("name"))
        ww.cell(r, 3, w.get("unit"))
        ww.cell(r, 4, w["qty"])
        ww.cell(r, 5, w.get("match", ""))
        ww.cell(r, 6, w.get("spec"))
        ww.cell(r, 7, w.get("kp") or 0)
        ww.cell(r, 8, w.get("brig") or 0)
        ww.cell(r, 9, w.get("norm") or 0)
        ww.cell(r, 10, w.get("rash_pct") or 0)
        ww.cell(r, 11, "да" if w["over"] > 0.005 else "")
        ww.cell(r, 12, f"=H{r}*D{r}")
        ww.cell(r, 13, f"=L{r}*J{r}")
        ww.cell(r, 14, (
            f'=IF(A{r}="",0,SUMIF(Материалы!$A$4:$A${mend},A{r},Материалы!$O$4:$O${mend}))'
            f'+IF(OR(B{r}="",B{r}=A{r}),0,SUMIF(Материалы!$A$4:$A${mend},B{r},Материалы!$O$4:$O${mend}))'
        ))
        ww.cell(r, 15, f"=L{r}+M{r}+N{r}")
        ww.cell(r, 16, f'=IF(SUM($O$4:$O${wend})=0,0,Косвенные!$C$4*O{r}/SUM($O$4:$O${wend}))')
        ww.cell(r, 17, f"=G{r}*D{r}+P{r}")
        ww.cell(r, 18, f"=IF(D{r}=0,0,Q{r}/D{r})")
        for c in range(1, 19):
            cell = ww.cell(r, c)
            cell.border = THIN
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            if c in (4, 7, 8, 9, 12, 13, 14, 15, 16, 17, 18):
                cell.number_format = "#,##0.00"
            if c == 10:
                cell.number_format = "0%"
        if w["over"] > 0.005:
            ww.cell(r, 11).fill = PatternFill("solid", fgColor=YELLOW)
        ww.row_dimensions[r].height = 30
    widths(ww, [12, 42, 8, 12, 14, 8, 12, 12, 12, 12, 12, 16, 14, 16, 14, 14, 14, 12])

    ind = wb.create_sheet("Косвенные", 0)
    title(ind, "Косвенные")
    head(ind, 3, ["Статья", "Расчёт", "Сумма"])
    if indirect:
        for i, item in enumerate(indirect):
            ind.cell(4 + i, 1, item.get("name")).border = THIN
            ind.cell(4 + i, 2, item.get("calc")).border = THIN
            c = ind.cell(4 + i, 3, float(item.get("sum") or 0))
            money(c)
        total_row = 4 + len(indirect)
        ind.cell(total_row, 1, "Итого").font = Font(bold=True)
        c = ind.cell(total_row, 3, f"=SUM(C4:C{total_row - 1})")
    else:
        total_row = 4
        ind.cell(4, 1, "Итого").font = Font(bold=True)
        c = ind.cell(4, 3, 0)
    money(c)
    c.font = Font(bold=True)
    # Работы ссылаются на C4 как на итог. Перенесём итог в C4, если статьи есть:
    # выше P ссылается на Косвенные!$C$4. Поправим на total_row.
    widths(ind, [36, 50, 16])

    # Ссылка на итог косвенных — фактическая строка, не всегда C4.
    if total_row != 4:
        for i in range(len(works)):
            r = 4 + i
            ww.cell(r, 16, f'=IF(SUM($O$4:$O${wend})=0,0,Косвенные!$C${total_row}*O{r}/SUM($O$4:$O${wend}))')
            ww.cell(r, 16).number_format = "#,##0.00"
            ww.cell(r, 16).border = THIN

    ws = wb.create_sheet("Свод", 0)
    title(ws, f"Бюджет себестоимости — {data.get('project', '')}")
    head(ws, 3, ["Статья", "Значение"])
    rows = [
        (4, "Материалы, с/с", f"=SUM(Материалы!O4:O{mend})", "#,##0.00"),
        (5, "Оплата бригадам", f"=SUM(Работы!L4:L{wend})", "#,##0.00"),
        (6, "Расходники", f"=SUM(Работы!M4:M{wend})", "#,##0.00"),
        (7, "Косвенные", f"=Косвенные!C{total_row}", "#,##0.00"),
        (8, "Полная себестоимость", "=B4+B5+B6+B7", "#,##0.00"),
        (9, "КП работ (с косвенными)", f"=SUM(Работы!Q4:Q{wend})", "#,##0.00"),
        (10, "КП материалов", f"=SUM(Материалы!P4:P{mend})", "#,##0.00"),
        (11, "КП без НДС", "=B9+B10", "#,##0.00"),
        (12, "НДС", "=B11*Параметры!B10", "#,##0.00"),
        (13, "КП с НДС", "=B11+B12", "#,##0.00"),
        (14, "Наценка на с/с", '=IF(B8=0,0,B11/B8-1)', "0.0%"),
        (15, "Маржа", '=IF(B11=0,0,1-B8/B11)', "0.0%"),
        (16, "Диапазон типа тендера", f"{tender.get('mk_from', 0):.0%} – {tender.get('mk_to', 0):.0%}", None),
        (17, "Переторжка −5 % (КП б/НДС)", "=B11*0.95", "#,##0.00"),
        (18, "Маржа при −5 %", '=IF(B17=0,0,1-B8/B17)', "0.0%"),
        (19, "Переторжка −10 % (КП б/НДС)", "=B11*0.9", "#,##0.00"),
        (20, "Маржа при −10 %", '=IF(B19=0,0,1-B8/B19)', "0.0%"),
        (21, "Минимальная цена, наценка 15 %", "=B8*1.15", "#,##0.00"),
        (22, "Трудоёмкость, чел.-ч", f"=SUMPRODUCT(Работы!I4:I{wend},Работы!D4:D{wend})", "#,##0.00"),
    ]
    if nmck:
        rows.append((23, "КП к НМЦК", '=IF(Параметры!B13=0,0,B11/Параметры!B13-1)', "0.0%"))
    for row, name, val, fmt in rows:
        ws.cell(row, 1, name).border = THIN
        c = ws.cell(row, 2, val)
        c.border = THIN
        if fmt:
            c.number_format = fmt
        if name in ("Полная себестоимость", "КП без НДС", "КП с НДС"):
            ws.cell(row, 1).font = Font(bold=True)
            c.font = Font(bold=True)
    ws.conditional_formatting.add(
        "B14",
        CellIsRule(operator="lessThan", formula=["0.15"],
                   fill=PatternFill("solid", fgColor=ORANGE)),
    )
    ws["A25"] = "Заказчику этот файл не передавать. Итоги — формулы, их не заменять числами."
    widths(ws, [42, 22])

    pas = wb.create_sheet("Паспорт тендера")
    title(pas, "Паспорт тендера")
    head(pas, 3, ["Поле", "Значение"])
    fields = [("Заказчик", data.get("customer")), ("Объект", data.get("project")),
              ("Регион", data.get("region")), ("Тип тендера", data.get("tender_type")),
              ("НМЦК б/НДС", "=Параметры!B13"),
              ("Форма КП", "под ключ" if turnkey else "раздельная")]
    fields += [(k, v) for k, v in data.get("passport", [])]
    for i, (k, v) in enumerate(fields):
        pas.cell(4 + i, 1, k).border = THIN
        cell = pas.cell(4 + i, 2, v)
        cell.border = THIN
        cell.alignment = Alignment(wrap_text=True)
        if k == "НМЦК б/НДС":
            cell.number_format = "#,##0.00"
    widths(pas, [28, 80])

    src = wb.create_sheet("Цены материалов (источники)")
    title(src, "Источники цен")
    head(src, 3, ["Материал", "Ед.", "Магазин", "Ссылка", "Цена с НДС", "Цена б/НДС", "Дата", "Примечание"])
    for i in range(len(mats)):
        r = 4 + i
        src.cell(r, 1, f"=Материалы!B{r}")
        src.cell(r, 2, f"=Материалы!D{r}")
        src.cell(r, 3, f"=Материалы!R{r}")
        src.cell(r, 4, f"=Материалы!S{r}")
        src.cell(r, 5, f"=Материалы!F{r}")
        src.cell(r, 6, f"=Материалы!H{r}")
        src.cell(r, 7, f"=Материалы!T{r}")
        src.cell(r, 8, f"=Материалы!U{r}")
        for c in range(1, 9):
            cell = src.cell(r, c)
            cell.border = THIN
            if c in (5, 6):
                cell.number_format = "#,##0.00"
    widths(src, [36, 8, 18, 40, 14, 14, 12, 28])

    rk = wb.create_sheet("Риски и вопросы")
    title(rk, "Риски и вопросы")
    head(rk, 3, ["№", "Описание", "Влияние, руб.", "Что спросить", "Допущение"])
    for i, item in enumerate(data.get("risks", [])):
        for c, v in enumerate([item.get("n", i + 1), item.get("desc"), item.get("impact"),
                               item.get("ask"), item.get("assumption")], 1):
            cell = rk.cell(4 + i, c, v)
            cell.border = THIN
            cell.alignment = Alignment(wrap_text=True)
        rk.row_dimensions[4 + i].height = 32
    widths(rk, [6, 40, 16, 36, 40])

    ad = wb.create_sheet("Добавить в тарификатор")
    title(ad, "Добавить в тарификатор")
    head(ad, 3, ["Работа", "Ед.", "Цена КП", "Оплата бригаде", "Норма, чел.-ч", "Основание"])
    extra = [(4 + i, w) for i, w in enumerate(works) if not spr_has(w)]
    for n, (src_row, w) in enumerate(extra):
        r = 4 + n
        ad.cell(r, 1, f"=Работы!B{src_row}")
        ad.cell(r, 2, f"=Работы!C{src_row}")
        ad.cell(r, 3, f"=Работы!G{src_row}")
        ad.cell(r, 4, f"=Работы!H{src_row}")
        ad.cell(r, 5, f"=Работы!I{src_row}")
        ad.cell(r, 6, w.get("src") or w.get("norm_src"))
        for c in range(1, 7):
            ad.cell(r, c).border = THIN
            if c in (3, 4, 5):
                ad.cell(r, c).number_format = "#,##0.00"
    widths(ad, [46, 10, 14, 18, 16, 36])

    kpws = wb.create_sheet("КП")
    kpws["A1"] = f"Коммерческое предложение — {data.get('project', '')}"
    kpws["A1"].font = Font(bold=True, size=14, color=NAVY)
    kpws["A2"] = "Цены без НДС. Суммы — формулы листа «Работы» и «Материалы»."
    turn_headers = ["№", "Позиция", "Ед.", "Кол-во", "Цена за ед.", "Сумма"]
    sep_headers = ["№", "Позиция", "Ед.", "Кол-во", "Материалы за ед.", "Работа за ед.", "Сумма"]
    headers = turn_headers if turnkey else sep_headers
    head(kpws, 4, headers)
    matched = []
    for i, w in enumerate(works):
        matched.append(4 + i)
    loose = []
    for i, m in enumerate(mats):
        key = str(m.get("work") or "")
        if not key or not any(key in (str(w.get("art") or ""), str(w.get("name") or "")) for w in works):
            loose.append(4 + i)
    kp_rows = []
    for n, src_row in enumerate(matched):
        r = 5 + n
        kp_rows.append(r)
        kpws.cell(r, 1, n + 1)
        kpws.cell(r, 2, f"=Работы!B{src_row}")
        kpws.cell(r, 3, f"=Работы!C{src_row}")
        kpws.cell(r, 4, f"=Работы!D{src_row}")
        mat = (
            f'IF(Работы!A{src_row}="",0,SUMIF(Материалы!$A$4:$A${mend},Работы!A{src_row},Материалы!$P$4:$P${mend}))'
            f'+IF(OR(Работы!B{src_row}="",Работы!B{src_row}=Работы!A{src_row}),0,'
            f'SUMIF(Материалы!$A$4:$A${mend},Работы!B{src_row},Материалы!$P$4:$P${mend}))'
        )
        if turnkey:
            kpws.cell(r, 5, f"=IF(D{r}=0,0,({mat})/D{r}+Работы!R{src_row})")
            kpws.cell(r, 6, f"=D{r}*E{r}")
        else:
            kpws.cell(r, 5, f"=IF(D{r}=0,0,({mat})/D{r})")
            kpws.cell(r, 6, f"=Работы!R{src_row}")
            kpws.cell(r, 7, f"=D{r}*E{r}+D{r}*F{r}")
    base = 5 + len(matched)
    for n, src_row in enumerate(loose):
        r = base + n
        kp_rows.append(r)
        kpws.cell(r, 1, len(matched) + n + 1)
        kpws.cell(r, 2, f"=Материалы!B{src_row}")
        kpws.cell(r, 3, f"=Материалы!D{src_row}")
        kpws.cell(r, 4, f"=Материалы!E{src_row}")
        if turnkey:
            kpws.cell(r, 5, f"=Материалы!N{src_row}")
            kpws.cell(r, 6, f"=D{r}*E{r}")
        else:
            kpws.cell(r, 5, f"=Материалы!N{src_row}")
            kpws.cell(r, 6, 0)
            kpws.cell(r, 7, f"=D{r}*E{r}+D{r}*F{r}")
    last = (kp_rows[-1] + 1) if kp_rows else 5
    sum_col = 6 if turnkey else 7
    if kp_rows:
        total_formula = f"=SUM({get_column_letter(sum_col)}{kp_rows[0]}:{get_column_letter(sum_col)}{kp_rows[-1]})"
    else:
        total_formula = "=0"
    labels = [("Итого без НДС", total_formula),
              ("НДС", f"={get_column_letter(sum_col)}{last}*Параметры!B10"),
              ("Итого с НДС", f"={get_column_letter(sum_col)}{last}+{get_column_letter(sum_col)}{last + 1}")]
    for i, (label, formula) in enumerate(labels):
        kpws.cell(last + i, 2, label).font = Font(bold=True)
        c = kpws.cell(last + i, sum_col, formula)
        c.font = Font(bold=True)
        c.number_format = "#,##0.00"
    for r in kp_rows:
        for c in range(1, sum_col + 1):
            cell = kpws.cell(r, c)
            cell.border = THIN
            if c >= 4:
                cell.number_format = "#,##0.00"
        kpws.row_dimensions[r].height = 28
    widths(kpws, [6, 55, 8, 12, 18, 16, 16])

    for name in ("Свод", "Работы", "Материалы", "КП"):
        wb[name].sheet_properties.tabColor = NAVY
    order = ["Свод", "Параметры", "Работы", "Материалы", "Косвенные", "КП",
             "Паспорт тендера", "Цены материалов (источники)", "Риски и вопросы",
             "Добавить в тарификатор"]
    for i, name in enumerate(order):
        current = wb.sheetnames.index(name)
        wb.move_sheet(name, offset=i - current)
    wb.save(path)


def spr_has(w):
    return bool(w.get("art")) and w.get("sec")


def write_kp(path, data, works, mats, kp, turnkey):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "КП"
    ws["A1"] = f"Коммерческое предложение — {data.get('project', '')}"
    ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    ws["A2"] = 'ООО «НЕОСТРОЙ». Цены без НДС, если не указано иное.'
    headers = ["№", "Позиция", "Ед.", "Кол-во", "Материалы за ед.", "Работа за ед.", "Сумма"]
    if turnkey:
        headers = ["№", "Позиция", "Ед.", "Кол-во", "Цена за ед.", "Сумма"]
    head(ws, 4, headers)
    rows = []
    for w in works:
        mat_sum = sum(m["kp_unit"] * m["qty"] for m in mats if str(m.get("work")) in (str(w.get("art")), w.get("name")))
        rows.append((w.get("name"), w.get("unit"), w["qty"], mat_sum / w["qty"] if w["qty"] else 0,
                     w["kp_unit_full"], mat_sum + w["kp_sum"]))
    loose = [m for m in mats if not any(str(m.get("work")) in (str(w.get("art")), w.get("name")) for w in works)]
    for m in loose:
        rows.append((m.get("name"), m.get("unit"), m["qty"], m["kp_unit"], 0, m["kp_unit"] * m["qty"]))
    ws["I1"] = 0.22
    ws["I1"].number_format = "0%"
    ws["H1"] = "НДС"
    sum_col = 6 if turnkey else 7
    for i, (name, unit, qty, mat_u, work_u, total) in enumerate(rows):
        del total
        r = 5 + i
        price = (mat_u + work_u) if turnkey else None
        vals = [i + 1, name, unit, qty, price if turnkey else mat_u, None if turnkey else work_u]
        for c, v in enumerate(vals, 1):
            if v is None:
                continue
            cell = ws.cell(r, c, v)
            cell.border = THIN
            if isinstance(v, float):
                cell.number_format = '#,##0.00'
        formula = f"=D{r}*E{r}" if turnkey else f"=D{r}*E{r}+D{r}*F{r}"
        cell = ws.cell(r, sum_col, formula)
        cell.border = THIN
        cell.number_format = '#,##0.00'
        ws.row_dimensions[r].height = 28
    last = 5 + len(rows)
    letter = get_column_letter(sum_col)
    if rows:
        total_formula = f"=SUM({letter}5:{letter}{last - 1})"
    else:
        total_formula = "=0"
    labels = [("Итого без НДС", total_formula),
              ("НДС 22 %", f"={letter}{last}*I1"),
              ("Итого с НДС", f"={letter}{last}+{letter}{last + 1}")]
    for i, (label, val) in enumerate(labels):
        ws.cell(last + i, 2, label).font = Font(bold=True)
        c = ws.cell(last + i, sum_col, val)
        c.font = Font(bold=True)
        c.number_format = '#,##0.00'
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{last - 1}"
    widths(ws, [6, 55, 8, 12, 18, 16, 16])
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddFooter.right.text = "ООО «НЕОСТРОЙ»  |  стр. &P из &N"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{4 + len(rows)}"
    wb.save(path)


def write_map(path, data, works, hours, kp):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ресурсы"
    title(ws, f"Карта реализации — {data.get('project', '')}")
    head(ws, 3, ["Код", "Чел.-ч", "Численность при 40 ч/нед", "Как отдавать"])
    by = {}
    for w in works:
        by.setdefault(w.get("spec") or "—", 0)
        by[w["spec"] or "—"] += (w.get("norm") or 0) * w["qty"]
    for i, (spec, h) in enumerate(by.items()):
        r = 4 + i
        ws.cell(r, 1, spec).border = THIN
        c = ws.cell(r, 2, h)
        money(c)
        cell = ws.cell(r, 3, f"=IF(B{r}=0,0,B{r}/40)")
        cell.border = THIN
        cell.number_format = "0.00"
        ws.cell(r, 4, "по тарификатору, лист «Нормы»").border = THIN
    end = 4 + len(by)
    ws.cell(end, 1, "Всего чел.-ч").font = Font(bold=True)
    total_h = ws.cell(end, 2, f"=SUM(B4:B{end - 1})" if by else 0)
    money(total_h)
    ws.cell(end + 1, 1, f"Сумма графика по деньгам равна КП {kp:,.0f} б/НДС. "
            "После правки бюджета карту собирают заново: это отдельный файл.")
    widths(ws, [14, 14, 28, 36])
    g = wb.create_sheet("График")
    title(g, "График")
    sched = data.get("schedule", [])
    if not sched:
        g["A3"] = ("Длительность раздела = сумма (норма × объём) / (численность × 8 ч × смены). "
                   "Даты в Расчёт.json не заданы — график заполнить, когда будут даты начала.")
        g.column_dimensions["A"].width = 100
    else:
        weeks = max(int(s.get("start_week", 1)) + int(s.get("weeks", 1)) - 1 for s in sched)
        head(g, 3, ["Работа", "Людей"] + [f"н{n}" for n in range(1, weeks + 1)])
        for i, s in enumerate(sched):
            g.cell(4 + i, 1, s.get("name")).border = THIN
            g.cell(4 + i, 2, s.get("people")).border = THIN
            start = int(s.get("start_week", 1))
            for n in range(int(s.get("weeks", 1))):
                cell = g.cell(4 + i, 2 + start + n, "■")
                cell.fill = PatternFill("solid", fgColor=PALE)
    wb.save(path)


def write_letter(path, data, kp):
    # Титульник.docx без стилей Word — письмо собирается заново, реквизиты из карточки.
    doc = Document()
    req = requisites()
    head_p = doc.add_paragraph()
    run = head_p.add_run("ООО «НЕОСТРОЙ»")
    run.bold = True
    if req:
        doc.add_paragraph(", ".join(f"{k}: {v}" for k, v in req.items() if v))
    title_p = doc.add_paragraph()
    title_run = title_p.add_run(f"Коммерческое предложение — {data.get('project', '')}")
    title_run.bold = True
    p = doc.add_paragraph()
    p.add_run(f"Предмет: {data.get('subject') or data.get('project', '')}. ")
    p.add_run(f"Стоимость работ и материалов — {kp:,.2f} руб. без НДС, ")
    p.add_run(f"{kp * VAT:,.2f} руб. с НДС 22 %.")
    for line in data.get("letter", []):
        doc.add_paragraph(line)
    doc.add_paragraph("Срок действия предложения — 30 дней, если в документах тендера не указан другой.")
    doc.add_paragraph("Генеральный директор _________________ / Улогов Р. В.")
    doc.add_paragraph("Исполнитель: Каменев Алексей, +7 921 370-86-27, contact@neostroi.com")
    doc.save(path)


def write_registry(out_dir, data, cost, kp, mat_ss, mat_kp, brig, work_kp):
    row = {
        "Дата расчёта": data.get("date"),
        "Заказчик": data.get("customer"),
        "Объект": data.get("project"),
        "Предмет": data.get("subject"),
        "Регион": data.get("region"),
        "Вид работ": data.get("work_kind"),
        "Тип тендера": data.get("tender_type"),
        "Период цены": data.get("period"),
        "НМЦК б/НДС": data.get("nmck"),
        "С/с б/НДС": round(cost, 2),
        "КП б/НДС": round(kp, 2),
        "КП с НДС": round(kp * VAT, 2),
        "Наценка мат.": round(mat_kp / mat_ss - 1, 4) if mat_ss else None,
        "Наценка раб.": round(work_kp / (brig or 1) - 1, 4) if brig else None,
        "Решение по расчёту": data.get("decision") or "рассчитано",
        "Переторжки (история)": data.get("retender") or "",
        "Папка": data.get("folder") or Path(out_dir).name,
        "Источник расчёта (память)": "tools/budget.py + тарификатор v2",
        "Примечания": data.get("notes") or "",
    }
    (Path(out_dir) / "Строка реестра.json").write_text(
        json.dumps(row, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    src = Path(sys.argv[1])
    build(json.loads(src.read_text(encoding="utf-8")), src.parent)
