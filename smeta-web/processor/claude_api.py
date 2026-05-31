"""
Трёхэтапный анализ сметы:
1. Извлечение структуры + всех позиций + всех материалов
2. Рыночные цены на материалы (построчно)
3. Рыночные расценки на работы (построчно)
"""
import json
import os
import anthropic

client = anthropic.Anthropic(api_key=os.environ.get('ANTHROPIC_API_KEY', ''))
MODEL = 'claude-sonnet-4-6'
MAX_TOKENS = 8000


# ── Вызов 1: Полная структура сметы ──────────────────────────────────────────

SYS_EXTRACT = """Ты анализируешь строительную смету. Верни ТОЛЬКО JSON, без пояснений.

АЛГОРИТМ:
1. total_no_vat = СУММА итогов всех разделов (НЕ НМЦ!)
2. vat: ищи "НДС 22% | 395 437" (числа могут быть с пробелами: "395 437"=395437). 
   Если нашёл → vat_rate=0.22, vat=число.
   Если не нашёл → vat = nmc - total_no_vat.
3. Структура затрат: ищи строки "материальные ресурсы | 787 069" и т.д.
4. Разделы: все разделы с ПОЛНЫМ списком позиций и материалов.
5. Единицы НАТУРАЛЬНЫЕ: "100 м²"→"м²" qty×100; "1000 шт"→"шт" qty×1000; "100 шт"→"шт" qty×100.
6. Итог позиции: "Всего по позиции = N" или последняя сумма позиции.
7. Если итог раздела=0 но позиции ненулевые → сложи позиции.

КОМПАКТНЫЙ JSON (экономь токены, используй короткие поля):
{
  "proj":{"name":"","addr":"","ikz":"","date":"","period":""},
  "st":{
    "mat":0,"ot":0,"otm":0,"em":0,"tr":0,"nr":0,"sp":0,
    "tot":0,"vat_r":0.22,"vat":0,"nmc":0
  },
  "sec":[{
    "n":"Раздел X. Название",
    "tot":0,
    "pos":[{"i":"1","c":"ГЭСН57-01-002-01","nm":"Разборка линолеума","u":"м²","q":58.4,"ot":3242,"sum":7817}],
    "mat":[{"c":"прайс","nm":"Виниловая плитка Decoria","u":"м²","q":59.568,"p":1057.38,"sum":62986}]
  }]
}

Только JSON."""


def extract_structure(text: str) -> dict:
    msg = client.messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=SYS_EXTRACT,
        messages=[{'role': 'user', 'content': f'Смета:\n\n{text}'}]
    )
    raw = msg.content[0].text.strip().replace('```json','').replace('```','').strip()
    data = json.loads(raw)
    # Нормализуем к стандартным именам полей
    return _normalize(data)


def _normalize(d: dict) -> dict:
    """Конвертирует компактные поля в стандартные."""
    proj = d.get('proj', d.get('project', {}))
    st   = d.get('st',   d.get('smeta_structure', {}))
    secs = d.get('sec',  d.get('sections', []))

    def _p(x, *keys):
        for k in keys:
            if k in x: return x[k]
        return 0

    result = {
        'project': {
            'name':       proj.get('name',''),
            'address':    proj.get('addr', proj.get('address','')),
            'ikz':        proj.get('ikz',''),
            'defect_date':proj.get('date', proj.get('defect_date','')),
            'price_period':proj.get('period', proj.get('price_period','')),
        },
        'smeta_structure': {
            'materials':   _p(st,'mat','materials'),
            'ot_workers':  _p(st,'ot','ot_workers'),
            'otm_machines':_p(st,'otm','otm_machines'),
            'machine_op':  _p(st,'em','machine_op'),
            'transport':   _p(st,'tr','transport'),
            'overhead':    _p(st,'nr','overhead'),
            'profit':      _p(st,'sp','profit'),
            'total_no_vat':_p(st,'tot','total_no_vat'),
            'vat_rate':    st.get('vat_r', st.get('vat_rate', 0.22)),
            'vat':         _p(st,'vat'),
            'nmc':         _p(st,'nmc'),
        },
        'sections': []
    }

    for sec in secs:
        positions = []
        for p in sec.get('pos', sec.get('positions', [])):
            positions.append({
                'num':      p.get('i',  p.get('num','')),
                'code':     p.get('c',  p.get('code','')),
                'name':     p.get('nm', p.get('name','')),
                'unit':     p.get('u',  p.get('unit','')),
                'qty':      p.get('q',  p.get('qty',0)),
                'ot_total': p.get('ot', p.get('ot_total',0)),
                'total':    p.get('sum',p.get('total',0)),
            })
        materials = []
        for m in sec.get('mat', sec.get('materials', [])):
            materials.append({
                'code':  m.get('c',  m.get('code','')),
                'name':  m.get('nm', m.get('name','')),
                'unit':  m.get('u',  m.get('unit','')),
                'qty':   m.get('q',  m.get('qty',0)),
                'price': m.get('p',  m.get('price',0)),
                'total': m.get('sum',m.get('total',0)),
            })
        result['sections'].append({
            'name':      sec.get('n', sec.get('name','')),
            'total':     sec.get('tot', sec.get('total',0)),
            'positions': positions,
            'materials': materials,
        })

    # Если materials=0 — сложи из разделов
    if result['smeta_structure']['materials'] == 0:
        result['smeta_structure']['materials'] = sum(
            m['total'] for s in result['sections'] for m in s['materials']
        )
    return result


# ── Вызов 2: Рыночные цены на материалы (построчно) ──────────────────────────

SYS_MAT_MARKET = """Ты эксперт рынка стройматериалов СПб 2026.
Получи список материалов из сметы. Для каждого материала укажи рыночную цену.

Цена сайта (с НДС 22%) → цена б/НДС = цена_сайта / 1.22
Источники: petrovich.ru, lemanapro.ru

Ценовые ориентиры (б/НДС, СПб 2026):
Клей CM11 25кг: 24 руб/кг; CM17: 42 руб/кг
Затирка CE40: 115 руб/кг
Грунтовка Тифенгрунд: 57 руб/л; акриловая ВД-АК-02: 100 руб/кг
Шпатлёвка Ветонит ТТ: 36 руб/кг; Кнауф Финишная: 29 руб/кг
Штукатурка Ротбанд: 38 руб/кг; Unterputz: 18 руб/кг
Краска ВД-АК Tikkurila: 450 руб/кг; Sniezka: 200 руб/кг
Самовыравнивающиеся смеси: 22 руб/кг
Раствор М150 сухой: 19 руб/кг
Мастика КН-2: 180 руб/кг
Гравий керамзит М800: 1800 руб/м³
Кварцвинил SPC: 700 руб/м²; линолеум ПВХ: 400 руб/м²
Керамогранит 30×60, 60×60: 750-900 руб/м²
Плитка настенная: 600-800 руб/м²
Плинтус ПВХ 48мм: 25 руб/м
Дюбели 6мм: 1.5 руб/шт; шурупы 3.5мм: 0.8 руб/шт
Заглушки/соединители/уголки плинтуса: 8-12 руб/шт
Панели потолочные Armstrong 600×600: 450 руб/м²
Рейка алюм. потолочная 100мм: 98 руб/м; гребёнка: 72 руб/м; уголок алюм.: 60 руб/м
Подвесы стальные: 40 руб/шт
Сетка стальная сварная 4мм: 120 руб/м²
Мешки строительные: 15 руб/шт
ГКЛВ 12.5мм: 157 руб/м²; ГКЛ 12.5мм: 123 руб/м²
Профиль ПН 28×27: 30 руб/м; ПП 60×27: 46 руб/м; стоечный 100×50: 120 руб/м
Двери ПВХ м²: 8000 руб/м²
Подоконник ПВХ: 800 руб/м
Стеклопакет однокамерный: 2500 руб/м²
Минвата базальт.: 5000-6000 руб/м³
Кабель ВВГнг 3×1.5: 53 руб/м; 3×2.5: 90 руб/м; 3×4: 140 руб/м
Труба гофр. ПВХ d20: 12 руб/м
Труба PP-RCT 20мм: 82 руб/м
Водонагреватель 50л: 6000 руб/шт
Машины/механизмы (маш.-ч): оценка 90% от сметной цены
Вода, электроэнергия, прочие мелкие: 90% от сметной цены
Неизвестный материал: 85% от сметной цены

Верни JSON-массив (по одному объекту на каждый входной материал, в том же порядке):
[{"p_vat": цена_с_НДС, "p": цена_б_НДС, "src": "источник/аналог"}]

Только JSON-массив."""


def get_material_market_prices(all_materials: list) -> list:
    """Получает рыночные цены для всех материалов. Возвращает список цен."""
    if not all_materials:
        return []

    # Компактный список для API
    mat_list = []
    for m in all_materials:
        mat_list.append(f"{m['name']}|{m['unit']}|{m['qty']}|{m['price']}")

    msg = client.messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=SYS_MAT_MARKET,
        messages=[{'role': 'user', 'content': 'Материалы (название|ед|кол|смет.цена):\n' + '\n'.join(mat_list)}]
    )
    raw = msg.content[0].text.strip().replace('```json','').replace('```','').strip()
    try:
        prices = json.loads(raw)
        if len(prices) == len(all_materials):
            return prices
        # Дополняем если меньше
        while len(prices) < len(all_materials):
            prices.append({"p_vat": 0, "p": 0, "src": "нет данных"})
        return prices[:len(all_materials)]
    except Exception:
        return [{"p_vat": round(m['price']*1.1), "p": m['price'], "src": "оценка"} for m in all_materials]


# ── Вызов 3: Рыночные расценки на работы (построчно) ─────────────────────────

SYS_WORK_MARKET = """Ты эксперт рынка строительных работ СПб 2026.
Получи список работ из сметы. Для каждой работы укажи рыночную расценку бригады.
Расценка = только работа без материалов (otdelka-spb.ru, demontazh-centr.ru).

Ориентиры расценок (руб/ед):
Разборка линолеума/ПВХ: 100/м²; паркета/фанеры: 200/м²; плитки пол: 450/м²
Разборка стяжки: 650/м²; оснований (лаги): 150/м²; плинтусов: 50/м
Затаривание мусора: 500/т
Теплоизоляция засыпная керамзит: 400/м³
Железнение/проливка: 80/м²
Стяжка ц/и 20мм: 550/м²; доп.слой 5мм: 100/м²; самовыравн.: 400/м²
Армирование стяжки: 150/м²
Укладка кварцвинила/ПВХ плитки: 500/м²; плинтус ПВХ: 100/м
Укладка керамогранита пол: 1600/м²
Демонтаж потолка Армстронг: 300/м²; реечного: 400/м²
Монтаж потолка Армстронг: 800/м²; реечного алюм.: 900/м²
Очистка стен от краски: 150/м²
Смена жалюзийных решёток: 300/шт
Перегородка ГКЛ 1 слой: 1200/м²; 2 слоя: 1800/м²
Штукатурка оконных/дверных откосов: 500/м²
Штукатурка внутр. ц/и улучш.: 750/м²
Облицовка стен плиткой: 1450/м²
Грунтовка: 50/м²; шпатлёвка стен: 400/м²; окраска ВД: 600/м²
Установка люков сантехн.: 800/шт
Разборка ГКЛ откосов: 200/м²; снятие подоконников: 200/шт
Установка подоконника ПВХ: 1000/м
Устройство каркаса откосов: 500/м²; облицовка ГКЛ откосов: 800/м²
Замена стеклопакетов: 2000/м²
Демонтаж дверей: 800/шт; установка ПВХ до 3м²: 3500/шт
Демонтаж светильников: 200/шт; выключателей: 100/шт
Трубы гофр. ПВХ монтаж: 80/м; затягивание кабеля: 50/м
Светильники накладные: 500/шт; встраиваемые: 600/шт
Штробление: 360/м; кабель в штробе: 180/м
Выключатели/розетки: 450/шт
Трубы водоснабжения d20: 180/м; монтаж водонагревателя: 2000/шт
Прочие работы / без аналога: 1.8 × (ОТ на ед.)

Верни JSON-массив (по одному объекту на каждую входную позицию, в том же порядке):
[{"r": расценка_руб_на_ед, "src": "источник"}]

Только JSON-массив."""


def get_work_market_rates(all_positions: list) -> list:
    """Получает рыночные расценки для всех позиций работ."""
    if not all_positions:
        return []

    pos_list = []
    for p in all_positions:
        ot_per = round(p['ot_total'] / p['qty'], 2) if p.get('qty') else 0
        pos_list.append(f"{p['name']}|{p['unit']}|{p['qty']}|ОТ_ед:{ot_per}")

    msg = client.messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=SYS_WORK_MARKET,
        messages=[{'role': 'user', 'content': 'Позиции (название|ед|объём|ОТ_на_ед):\n' + '\n'.join(pos_list)}]
    )
    raw = msg.content[0].text.strip().replace('```json','').replace('```','').strip()
    try:
        rates = json.loads(raw)
        if len(rates) == len(all_positions):
            return rates
        while len(rates) < len(all_positions):
            rates.append({"r": 0, "src": "нет данных"})
        return rates[:len(all_positions)]
    except Exception:
        # Запасной вариант: 1.8 × ОТ на ед.
        result = []
        for p in all_positions:
            ot_per = round(p['ot_total'] / p['qty'], 2) if p.get('qty') else 0
            result.append({"r": round(ot_per * 1.8), "src": "оценка"})
        return result


# ── Сборка рыночного анализа из детальных данных ─────────────────────────────

def get_market_analysis(smeta_data: dict, mat_prices: list, work_rates: list) -> dict:
    """
    Собирает полный рыночный анализ из детальных данных по каждой строке.
    Не требует API-вызова — только математика.
    """
    st = smeta_data['smeta_structure']
    secs = smeta_data['sections']

    # Назначаем рыночные цены материалам
    mat_idx = 0
    for sec in secs:
        for mat in sec.get('materials', []):
            if mat_idx < len(mat_prices):
                mp = mat_prices[mat_idx]
                mat['market_price_vat'] = mp.get('p_vat', 0)
                mat['market_price']     = mp.get('p', 0)
                mat['market_total']     = round(mat['qty'] * mp.get('p', 0))
                mat['economy']          = mat['total'] - mat['market_total']
                mat['market_source']    = mp.get('src', '')
            mat_idx += 1

    # Назначаем расценки работам
    pos_idx = 0
    for sec in secs:
        for pos in sec.get('positions', []):
            if pos_idx < len(work_rates):
                wr = work_rates[pos_idx]
                pos['market_rate']  = wr.get('r', 0)
                pos['market_total'] = round(pos['qty'] * wr.get('r', 0)) if pos.get('qty') else 0
                pos['work_source']  = wr.get('src', '')
            pos_idx += 1

    # Считаем итоги по разделам
    market_sections = []
    for sec in secs:
        mat_smeta  = sum(m['total']        for m in sec.get('materials', []))
        mat_market = sum(m.get('market_total', 0) for m in sec.get('materials', []))
        wrk_market = sum(p.get('market_total', 0) for p in sec.get('positions',  []))
        market_sections.append({
            'name':              sec['name'],
            'mat_smeta':         mat_smeta,
            'mat_market_no_vat': mat_market,
            'work_market':       wrk_market,
            'economy_mat':       mat_smeta - mat_market,
            'economy_work':      sum(p.get('ot_total',0) for p in sec.get('positions',[])) - wrk_market,
        })

    # Глобальные итоги
    total_no_vat   = st['total_no_vat']
    mat_market_tot = sum(s['mat_market_no_vat'] for s in market_sections)
    wrk_market_tot = sum(s['work_market']       for s in market_sections)
    otm   = st.get('otm_machines', 0)
    em    = st.get('machine_op',   0)
    tr    = st.get('transport',    0)
    nr    = st.get('overhead',     0)
    sp    = st.get('profit',       0)

    net_eco     = total_no_vat - mat_market_tot - wrk_market_tot - otm - em - tr
    net_eco_pct = round(net_eco / total_no_vat * 100, 1) if total_no_vat else 0

    smeta_data['market'] = {
        'note':            'petrovich.ru/lemanapro.ru ÷1.22 (НДС 22%), otdelka-spb.ru, mai 2026',
        'vat_divisor':     1.22,
        'materials_no_vat':mat_market_tot,
        'works_total':     wrk_market_tot,
        'sections':        market_sections,
        'overhead_saved':  nr,
        'profit_saved':    sp,
        'extra_work_cost': wrk_market_tot - st.get('ot_workers', 0),
        'net_economy':     round(net_eco),
        'net_economy_pct': net_eco_pct,
    }
    return smeta_data
