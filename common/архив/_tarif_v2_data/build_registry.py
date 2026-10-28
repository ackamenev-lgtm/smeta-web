"""Реестр тендеров НЕОСТРОЙ. python -I build_registry.py <scratchpad> <out.xlsx>"""
import sys, os, json, math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

SP, OUT = sys.argv[1], sys.argv[2]
rows = []
import glob
for p in sorted(glob.glob(os.path.join(SP, "reg_out_*.json"))):
    rows += json.load(open(p, encoding="utf-8-sig"))
# дубли между пакетами: оставляем профильную запись
def drop(x):
    mf, obj, subj, cust = x.get("memory_file"), x.get("object") or "", (x.get("subject") or "").lower(), x.get("customer") or ""
    if mf == "project_budgets_sept2026_3folders.md" and "Тосно" in obj: return True
    if mf == "project_tendery_avgust_2026.md" and "МОНЛИД" in cust and "парковк" in subj: return True
    if mf == "project_metro_katalog_smr.md": return True
    return False
rows = [x for x in rows if not drop(x)]
for x in rows:
    if x.get("memory_file") == "project_tendery_29sep2026_c.md" and "METRO" in (x.get("customer") or ""):
        x["notes"] = ((x.get("notes") or "") + " Первый вариант 11.09: с/с 11,23 млн, КП 16,85 млн б/НДС.").strip()
    if x.get("memory_file") == "project_pyaterochka_rc.md":
        x["result"] = "проиграли"; x["place_note"] = "По цене 1-е место в 5 лотах из 9; проиграли на квалификации (оборот 20–50 млн в анкете). Итог 12.08.2026."
rows.sort(key=lambda x: (x.get("date") or "9999", x.get("customer") or ""))

B = Side(style="thin", color="BFBFBF"); BOX = Border(left=B, right=B, top=B, bottom=B)
HDR = PatternFill("solid", fgColor="1F3864"); HDR2 = PatternFill("solid", fgColor="7F6000"); WH = Font(color="FFFFFF", bold=True)
INP = PatternFill("solid", fgColor="FFF2CC"); SEC = PatternFill("solid", fgColor="D9E1F2")
def CF(c): return PatternFill("solid", fgColor=c, bgColor=c)
NUM0 = '#,##0;-#,##0;"–"'; PCT = '0.0%;-0.0%;"–"'

TYPES = ["Крупные сети и девелоперы — конкурентный тендер", "Переторжка / 2 тур / вход к новому заказчику",
         "Fit-out, частный бизнес, малые объёмы", "ТО и абонентское обслуживание", "Госзакупки 44-ФЗ / 223-ФЗ"]
KINDS = ["СМР / ремонт", "Fit-out", "КТО / ТО", "Рамочные расценки", "Поставка и монтаж", "ПИР+СМР", "Другое"]
DEC = ["подано", "рассчитано", "переторжка", "не участвовать"]
RES = ["ждём итогов", "выиграли", "проиграли", "отменён", "не подавали"]
PER = ["разово", "в месяц", "в год", "за срок контракта"]

wb = Workbook()
ws = wb.active; ws.title = "Реестр"
ws["A1"] = "РЕЕСТР ТЕНДЕРОВ НЕОСТРОЙ"; ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = ("Синие колонки — расчёт (заполнено по расчётам с июля 2026). Жёлтые колонки — РЕЗУЛЬТАТ, заполняет отдел тендеров после подведения итогов. "
            "Наценка и маржа считаются формулами из с/с и КП. Цифры — руб. без НДС за период из колонки «Период». Аналитика — на листе «Аналитика».")
ws.merge_cells("A2:R2"); ws["A2"].alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[2].height = 32
COLS = [  # (заголовок, ширина, ключ/формула, тип)
    ("№", 5), ("Дата расчёта", 11), ("Заказчик", 22), ("Объект", 30), ("Предмет", 40), ("Регион", 10), ("Вид работ", 14),
    ("Тип тендера", 28), ("Период цены", 10), ("НМЦК б/НДС", 14), ("С/с б/НДС", 14), ("КП б/НДС", 14), ("КП с НДС", 14),
    ("Наценка\n= КП ÷ с/с − 1", 10), ("Маржа\n= 1 − с/с ÷ КП", 10), ("КП к НМЦК", 9), ("Наценка мат.", 8), ("Наценка раб.", 8),
    ("Решение по расчёту", 13), ("Переторжки (история)", 30),
    ("Подано, дата", 11), ("Наша финальная цена б/НДС", 14), ("РЕЗУЛЬТАТ", 13), ("Место", 7), ("Цена победителя б/НДС", 14),
    ("Наша цена к победителю\n= наша ÷ победитель − 1", 12), ("Участников", 9), ("Причина / комментарий по итогам", 34),
    ("Папка", 22), ("Источник расчёта (память)", 26), ("Примечания", 50),
]
H = 4
for i, (t, w) in enumerate(COLS, 1):
    c = ws.cell(H, i, t); c.font, c.border = WH, BOX
    c.fill = HDR2 if 21 <= i <= 28 else HDR
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.column_dimensions[L(i)].width = w
ws.row_dimensions[H].height = 80
ws.cell(3, 1, "РАСЧЁТ").font = Font(bold=True, color="1F3864"); ws.cell(3, 21, "ИТОГИ ТЕНДЕРА (заполнять)").font = Font(bold=True, color="7F6000")
N_EMPTY = 60
R0 = H + 1
def num(v): return isinstance(v, (int, float)) and not isinstance(v, bool)
for k, x in enumerate(rows + [None] * N_EMPTY):
    r = R0 + k
    x = x or {}
    mk = f'=IF(AND(ISNUMBER(K{r}),ISNUMBER(L{r}),K{r}>0),L{r}/K{r}-1,{x["markup"] if num(x.get("markup")) else chr(34)*2})'
    mg = f'=IF(AND(ISNUMBER(K{r}),ISNUMBER(L{r}),L{r}>0),1-K{r}/L{r},{x["margin"] if num(x.get("margin")) else chr(34)*2})'
    vals = [k + 1, x.get("date"), x.get("customer"), x.get("object"), x.get("subject"), x.get("region"), x.get("kind"),
            x.get("tender_type"), x.get("period"), x.get("nmck_wo_vat"), x.get("cost_wo_vat"), x.get("kp_wo_vat"), x.get("kp_with_vat"),
            mk, mg, f'=IF(AND(ISNUMBER(J{r}),ISNUMBER(L{r}),J{r}>0),L{r}/J{r}-1,"")',
            x.get("mat_markup"), x.get("work_markup"), x.get("decision"), x.get("rebids"),
            None, x.get("kp_wo_vat") if x.get("result") else None,
            x.get("result") or ("не подавали" if x.get("decision") == "не участвовать" else ("ждём итогов" if x.get("decision") in ("подано", "переторжка") else None)), None, None,
            f'=IF(AND(ISNUMBER(V{r}),ISNUMBER(Y{r}),Y{r}>0),V{r}/Y{r}-1,"")', None, x.get("place_note"),
            x.get("folder"), x.get("memory_file"), x.get("notes")]
    if not x:
        vals[0] = k + 1
    for j, v in enumerate(vals, 1):
        if j == 2 and v:
            import datetime
            try: v = datetime.date.fromisoformat(v)
            except Exception: pass
        c = ws.cell(r, j, v); c.border = BOX
        c.alignment = Alignment(vertical="top", wrap_text=j in (4, 5, 20, 28, 31))
        if j == 2: c.number_format = "DD.MM.YYYY"
        if j in (10, 11, 12, 13, 22, 25): c.number_format = NUM0
        if j in (14, 15, 16, 17, 18, 26): c.number_format = PCT
        if 21 <= j <= 28: c.fill = INP
        if j == 21: c.number_format = "DD.MM.YYYY"
RL = R0 + len(rows) + N_EMPTY - 1
for col, lst in (("H", TYPES), ("G", KINDS), ("S", DEC), ("W", RES), ("I", PER)):
    dv = DataValidation(type="list", formula1='"' + ",".join(lst) + '"', allow_blank=True)
    if len('"' + ",".join(lst) + '"') > 255:   # длинный список — через лист
        pass
    ws.add_data_validation(dv); dv.add(f"{col}{R0}:{col}{RL}")
ws.conditional_formatting.add(f"W{R0}:W{RL}", FormulaRule(formula=[f'W{R0}="выиграли"'], fill=CF("C6E0B4")))
ws.conditional_formatting.add(f"W{R0}:W{RL}", FormulaRule(formula=[f'W{R0}="проиграли"'], fill=CF("F8CBAD")))
ws.conditional_formatting.add(f"O{R0}:O{RL}", FormulaRule(formula=[f'AND(ISNUMBER(O{R0}),O{R0}<0)'], fill=CF("F8CBAD")))
ws.conditional_formatting.add(f"O{R0}:O{RL}", FormulaRule(formula=[f'AND(ISNUMBER(O{R0}),O{R0}>=0,O{R0}<0.15)'], fill=CF("FFE699")))
ws.conditional_formatting.add(f"S{R0}:S{RL}", FormulaRule(formula=[f'S{R0}="не участвовать"'], fill=CF("D9D9D9")))
ws.freeze_panes = ws.cell(R0, 4)
ws.auto_filter.ref = f"A{H}:{L(len(COLS))}{RL}"

# ---- списки для проверки данных (тип тендера длинный — через лист)
wl = wb.create_sheet("Списки")
for i, t in enumerate(TYPES, 1): wl.cell(i, 1, t)
wl.sheet_state = "hidden"
ws.data_validations.dataValidation = [d for d in ws.data_validations.dataValidation if not str(d.sqref).startswith("H")]
dvt = DataValidation(type="list", formula1=f"=Списки!$A$1:$A${len(TYPES)}", allow_blank=True); ws.add_data_validation(dvt); dvt.add(f"H{R0}:H{RL}")

# ============ Аналитика ============
wa = wb.create_sheet("Аналитика")
wa["A1"] = "АНАЛИТИКА: ПРИ КАКОЙ НАЦЕНКЕ МЫ ВЫИГРЫВАЕМ"; wa["A1"].font = Font(bold=True, size=14)
wa["A2"] = ("Считается формулами по листу «Реестр». Колонки результата заполнятся, когда в реестре будут проставлены итоги (РЕЗУЛЬТАТ, цена победителя). "
            "Цель — через 15–20 закрытых тендеров откалибровать диапазоны наценки на листе «Тип тендера» файла материалов.")
wa.merge_cells("A2:L2"); wa["A2"].alignment = Alignment(wrap_text=True, vertical="top"); wa.row_dimensions[2].height = 32
RR = f"Реестр!$H${R0}:$H${RL}"; RK = f"Реестр!$G${R0}:$G${RL}"
RN = f"Реестр!$N${R0}:$N${RL}"; RO = f"Реестр!$O${R0}:$O${RL}"; RW = f"Реестр!$W${R0}:$W${RL}"; RZ = f"Реестр!$Z${R0}:$Z${RL}"
RS = f"Реестр!$S${R0}:$S${RL}"

def block(r0, title, key_rng, keys, label):
    wa.cell(r0, 1, title).font = Font(bold=True, size=12, color="1F3864")
    hdr = [label, "Расчётов", "Не участвовали", "Средняя наценка (все)", "Средняя маржа (все)", "Ждём итогов",
           "Выиграли", "Проиграли", "Доля побед", "Средняя наценка — выигранные", "Средняя наценка — проигранные",
           "Наша цена к победителю (проигранные), среднее"]
    for i, t in enumerate(hdr, 1):
        c = wa.cell(r0 + 1, i, t); c.fill, c.font, c.border = HDR, WH, BOX
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    wa.row_dimensions[r0 + 1].height = 62
    for k, key in enumerate(keys):
        r = r0 + 2 + k
        q = f'$A{r}'
        v = [key, f"=COUNTIF({key_rng},{q})", f'=COUNTIFS({key_rng},{q},{RS},"не участвовать")',
             f'=IFERROR(AVERAGEIFS({RN},{key_rng},{q},{RS},"<>не участвовать"),"")',
             f'=IFERROR(AVERAGEIFS({RO},{key_rng},{q},{RS},"<>не участвовать"),"")',
             f'=COUNTIFS({key_rng},{q},{RW},"ждём итогов")',
             f'=COUNTIFS({key_rng},{q},{RW},"выиграли")', f'=COUNTIFS({key_rng},{q},{RW},"проиграли")',
             f'=IF(G{r}+H{r}=0,"",G{r}/(G{r}+H{r}))',
             f'=IFERROR(AVERAGEIFS({RN},{key_rng},{q},{RW},"выиграли"),"")',
             f'=IFERROR(AVERAGEIFS({RN},{key_rng},{q},{RW},"проиграли"),"")',
             f'=IFERROR(AVERAGEIFS({RZ},{key_rng},{q},{RW},"проиграли"),"")']
        for j, val in enumerate(v, 1):
            c = wa.cell(r, j, val); c.border = BOX
            if j in (4, 5, 10, 11, 12): c.number_format = PCT
            if j == 9: c.number_format = "0%"
    rt = r0 + 2 + len(keys)
    wa.cell(rt, 1, "ИТОГО").font = Font(bold=True)
    for j, col in ((2, "B"), (3, "C"), (6, "F"), (7, "G"), (8, "H")):
        c = wa.cell(rt, j, f"=SUM({col}{r0+2}:{col}{rt-1})"); c.font = Font(bold=True); c.border = BOX
    return rt + 3

for i, w in enumerate([46, 10, 11, 12, 12, 12, 10, 10, 10, 13, 13, 14], 1):
    wa.column_dimensions[L(i)].width = w
nr = block(4, "По типу тендера", RR, TYPES, "Тип тендера")
nr = block(nr, "По виду работ", RK, KINDS, "Вид работ")
wa.cell(nr, 1, "Как читать: если «средняя наценка — выигранные» стабильно ниже целевого диапазона — диапазон завышен; "
                "если проигрываем с ценой выше победителя на 5–10 % — снижать наценку на работы в этом типе тендеров.").font = Font(italic=True, color="595959")

# ============ Как вести ============
wh = wb.create_sheet("Как вести")
wh.column_dimensions["A"].width = 140
T = ["КАК ВЕСТИ РЕЕСТР",
     "1. Новый тендер добавляется в реестр после каждого расчёта (шаг 12 промта «Расчёт тендера и КП»): дата, заказчик, объект, тип, с/с, КП, наценка, решение.",
     "2. После подачи — отдел тендеров ставит дату подачи и нашу финальную цену (после переторжки — последнюю).",
     "3. После итогов — РЕЗУЛЬТАТ (выиграли / проиграли / отменён), место, цену победителя и число участников, причину проигрыша (цена, квалификация, сроки, отказ заказчика).",
     "4. Цену победителя спрашивать у организатора или смотреть в протоколе (на ЭТП она обычно публикуется). Без неё нельзя понять, насколько мы дороже рынка.",
     "5. Раз в месяц смотреть лист «Аналитика»: доля побед и наценка выигранных тендеров по типам — по ним корректируем целевые диапазоны наценки.",
     "Синие колонки заполнены по расчётам в памяти Claude (июль–октябрь 2026). Решение «подано» стоит только там, где подача была явно зафиксирована; РЕЗУЛЬТАТ пуст, если неизвестно, подавали ли; "
     "по остальным «рассчитано» — проверьте и поправьте статус.",
     "Наценка/маржа: формула из с/с и КП; если какой-то из сумм нет — подставлено значение из расчёта. Периоды разные (разово / в месяц / в год) — суммы между строками не складывать."]
for i, t in enumerate(T, 1):
    c = wh.cell(i, 1, t); c.alignment = Alignment(wrap_text=True, vertical="top")
    c.font = Font(bold=True, size=14, color="1F3864") if i == 1 else Font()
    if i > 1: wh.row_dimensions[i].height = 15 * max(1, math.ceil(len(t) / 135))

wb._sheets = [wb["Реестр"], wb["Аналитика"], wb["Как вести"], wb["Списки"]]
wb.active = 0
wb.save(OUT)
print("rows", len(rows))
