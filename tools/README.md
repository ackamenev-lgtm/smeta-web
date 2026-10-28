# Скрипты расчёта

Справочники xlsx не читать. Всё нужное из них достают эти команды.

```bash
python tools/tarif.py search перегородка ГКЛ -n 5     # артикул, цена КП, бригада, норма
python tools/tarif.py art 8.1.22                      # карточка + комплект материалов
python tools/tarif.py kit 7.2.25
python tools/tarif.py cat кабель                      # отходы и наценки категории
python tools/tarif.py rash ELM                        # % расходников
python tools/tarif.py rate OTD                        # ставка часа
python tools/tarif.py types
python tools/tarif.py params
python tools/prices.py search "кабель ВВГнг 3х1.5"    # база цен; «устарела» — старше 45 дней
python tools/prices.py top outbox/<лот>/суммы.json    # кому искать 2 источника
python tools/budget.py outbox/<лот>/Расчёт.json       # бюджет, КП, карта, письмо, строка реестра
python tools/render.py бюджет.xlsx --sheets "Свод,Работы,Материалы" --out outbox/<лот>/png
```

`--msk` у поиска по тарификатору — цены Москвы. Регион по умолчанию СПб.

## суммы.json — аргумент `prices.py top`

```json
[{"name": "Кабель ВВГнг 3х2.5", "sum": 480000}, {"name": "Краска ВД", "sum": 12000}]
```

`sum` — оценка суммы позиции в рублях. Скрипт помечает, каким позициям нужны два источника.

## Цены.json — что лот сдаёт в базу

Пишется в `outbox/<лот>/Цены.json`. Базу `common/` лот не трогает: её собирает командный чат командой `python tools/prices.py merge`.

```json
[{"name": "Кабель ВВГнг(А)-LS 3х1.5", "unit": "м", "shop": "ЭТМ", "url": "https://...",
  "price_vat": 85.5, "date": "2026-10-28", "note": "", "vat_included": true}]
```

## Расчёт.json — вход бюджета

Цены работ не писать: их подставит тарификатор по артикулу. `kp` и `brig` нужны только для работ без артикула.

```json
{
  "project": "Кровля склада, Пулково",
  "customer": "ООО «Пример»",
  "subject": "Ремонт кровли",
  "region": "СПб",
  "work_kind": "СМР / ремонт",
  "tender_type": "Крупные сети и девелоперы — конкурентный тендер",
  "period": "разово",
  "date": "2026-10-28",
  "nmck": null,
  "form": "separate",
  "months_to_purchase": 1,
  "advance_covers_materials": true,
  "finance_months": 0,
  "decision": "рассчитано",
  "passport": [["Срок подачи", "01.11.2026"], ["Оплата", "аванс 30 %"]],
  "works": [
    {"name": "Окраска стен", "unit": "м2", "qty": 120, "art": "8.2.34", "match": "точно"}
  ],
  "materials": [
    {"work": "8.2.34", "name": "Краска ВД", "cat": "Краски, грунты, ЛКМ", "unit": "кг",
     "qty": 40, "price_vat": 450, "retail_vat": 520, "shop": "Петрович",
     "url": "https://...", "date": "2026-10-28", "note": "", "vat_included": true}
  ],
  "indirect": [{"name": "Прораб", "calc": "0,5 мес × 150 000", "sum": 75000}],
  "risks": [{"desc": "Марка краски не задана", "impact": 0, "ask": "Уточнить марку", "assumption": "Tikkurila"}],
  "schedule": [{"name": "Окраска", "people": 2, "start_week": 1, "weeks": 2}],
  "letter": ["Срок работ — 15 рабочих дней.", "Вывоз мусора не входит."]
}
```

`form`: `separate` (раздельная) или `turnkey` (под ключ). `region`: СПб, Москва, Ленобласть, Подмосковье или другой — для другого доставка 10 % и цены СПб. `service: true` у материала — проверяемая услуга (вывоз, аренда): наценка 15 %, без отходов. `cat` — название категории из `tarif.py cat`.
