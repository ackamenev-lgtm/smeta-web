"""
Трёхэтапный анализ сметы.
"""
import json
from processor.json_repair import repair_json
import os
import anthropic

MODEL = 'claude-sonnet-4-6'
MAX_TOKENS = 8000

def _client():
    """Ленивая инициализация клиента."""
    return anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])


SYS_EXTRACT = """Ты анализируешь строительную смету. Верни ТОЛЬКО JSON, без пояснений.

АЛГОРИТМ:
1. total_no_vat = СУММА итогов всех разделов (НЕ НМЦ!)
2. vat: ищи "НДС 22% | 395 437" (числа с пробелами: "395 437"=395437).
   Если нашёл → vat_rate=0.22, vat=число. Если не нашёл → vat = nmc - total_no_vat.
3. Структура: ищи "материальные ресурсы | 787 069" и т.д.
4. Все разделы с ПОЛНЫМ списком позиций и материалов.
5. Единицы НАТУРАЛЬНЫЕ: "100 м²"→"м²" qty×100; "1000 шт"→"шт" qty×1000; "100 шт"→"шт" qty×100.
6. Если итог раздела=0 но позиции ненулевые → сложи позиции.

КОМПАКТНЫЙ JSON:
{
  "proj":{"name":"","addr":"","ikz":"","date":"","period":""},
  "st":{"mat":0,"ot":0,"otm":0,"em":0,"tr":0,"nr":0,"sp":0,"tot":0,"vat_r":0.22,"vat":0,"nmc":0},
  "sec":[{
    "n":"Раздел X. Название","tot":0,
    "pos":[{"i":"1","c":"ГЭСН57-01-002-01","nm":"Разборка линолеума","u":"м²","q":58.4,"ot":3242,"sum":7817}],
    "mat":[{"c":"прайс","nm":"Виниловая плитка Decoria","u":"м²","q":59.568,"p":1057.38,"sum":62986}]
  }]
}
Только JSON."""


SYS_MAT_MARKET = """Ты эксперт рынка стройматериалов СПб 2026.
Для каждого материала дай рыночную цену. Цена б/НДС = цена_сайта / 1.22.

Ориентиры (б/НДС):
Клей CM11: 24 руб/кг; CM17: 42; Затирка CE40: 115; Грунтовка Тифенгрунд: 57/л; акрил ВД-АК: 100/кг
Шпатлёвка цем Ветонит ТТ: 36/кг; гипс Кнауф: 29; Штукатурка Ротбанд: 38; цем Кнауф: 18
Краска ВД-АК Tikkurila: 450/кг; Sniezka: 200; Самовыравниваюшиеся: 22/кг; М150 сухой: 19
Мастика КН-2: 180/кг; Керамзит М800: 1800/м³; Кварцвинил SPC: 700/м²; Лин.ПВХ: 400/м²
Керамогранит 30×60: 800/м²; Плитка настенная: 650/м²; Плинтус ПВХ: 25/м
Дюбели 6мм: 1.5/шт; Шурупы 3.5мм: 0.8; Заглушки/соединители/уголки плинтуса: 10/шт
Панели Armstrong 600×600: 450/м²; Рейка алюм 100мм: 98/м; Гребёнка: 72/м; Уголок алюм: 60/м
Подвесы стальные: 40/шт; Сетка сварная 4мм: 120/м²; Мешки строит.: 15/шт
ГКЛВ 12.5мм: 157/м²; ГКЛ 12.5мм: 123/м²; Профиль ПН28: 30/м; ПП60: 46/м; стоечный100: 120/м
Двери ПВХ: 8000/м²; Подоконник ПВХ: 800/м; Стеклопакет: 2500/м²; Минвата: 5500/м³
Кабель ВВГнг3×1.5: 53/м; 3×2.5: 90; Труба гофр ПВХ d20: 12/м; PP-RCT 20мм: 82/м
Водонагреватель 50л: 6000/шт; Механизмы/маш.-ч: 90% от смет.цены; Прочее: 88% от смет.цены

Верни JSON-массив (один объект на каждый входной материал, тот же порядок):
[{"p_vat": цена_с_НДС, "p": цена_б_НДС, "src": "аналог/источник"}]
Только JSON-массив."""


SYS_WORK_MARKET = """Ты эксперт рынка строительных работ СПб 2026.
Для каждой работы дай рыночную расценку бригады (только работа, без материалов).
Источники: otdelka-spb.ru, demontazh-centr.ru.

Ориентиры (руб/ед):
Разборка: линолеума 100/м²; паркета/фанеры 200/м²; плитки пол 450/м²; стяжки 650/м²
Разборка оснований 150/м²; плинтусов 50/м; затаривание мусора 500/т
Теплоизоляция засыпная 400/м³; железнение 80/м²
Стяжка ц/и 20мм 550/м²; доп.слой 100/м²; самовыравн 400/м²; армирование 150/м²
Укладка кварцвинила 500/м²; плинтус ПВХ 100/м; керамогранит пол 1600/м²
Демонтаж потолка Армстронг 300/м²; реечного 400/м²
Монтаж потолка Армстронг 800/м²; реечного алюм 900/м²
Очистка стен от краски 150/м²; смена жалюзи 300/шт
Перегородки ГКЛ 1сл 1200/м²; 2сл 1800/м²; отбивка штукатурки 280/м²
Штукатурка: откосов 500/м²; ц/и улучш 750/м²; грунтовка 50/м²; шпатлёвка 400/м²
Плитка стены 1450/м²; окраска ВД 600/м²; люки сантехн 800/шт
Разборка ГКЛ 200/м²; снятие подоконников 200/шт; установка подоконника ПВХ 1000/м
Каркас откосов 500/м²; ГКЛ откосы 800/м²; замена стеклопакетов 2000/м²
Демонтаж дверей 800/шт; установка ПВХ до 3м² 3500/шт
Демонтаж светильников 200/шт; выключателей 100/шт
Трубы гофр монтаж 80/м; кабель в трубе 50/м; накладные светильники 500/шт; встраиваемые 600/шт
Штробление 360/м; кабель в штробе 180/м; выключатели/розетки 450/шт
Трубы ХВС/ГВС d20 180/м; монтаж водонагревателя 2000/шт; вывоз мусора: по смете
Неизвестная работа: 1.8 × (ОТ на ед.)

Верни JSON-массив (один объект на каждую входную позицию, тот же порядок):
[{"r": расценка_руб_на_ед, "src": "источник"}]
Только JSON-массив."""


def _normalize(d: dict) -> dict:
    proj = d.get('proj', d.get('project', {}))
    st   = d.get('st',   d.get('smeta_structure', {}))
    secs = d.get('sec',  d.get('sections', []))

    def _p(x, *keys):
        for k in keys:
            if k in x: return x[k]
        return 0

    result = {
        'project': {
            'name':        proj.get('name',''),
            'address':     proj.get('addr', proj.get('address','')),
            'ikz':         proj.get('ikz',''),
            'defect_date': proj.get('date', proj.get('defect_date','')),
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
    if result['smeta_structure']['materials'] == 0:
        result['smeta_structure']['materials'] = sum(
            m['total'] for s in result['sections'] for m in s['materials']
        )
    return result


def extract_structure(text: str) -> dict:
    msg = _client().messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=SYS_EXTRACT,
        messages=[{'role':'user','content':f'Смета:\n\n{text}'}]
    )
    raw = msg.content[0].text.strip().replace('```json','').replace('```','').strip()
    return _normalize(repair_json(raw))


def get_material_market_prices(all_materials: list) -> list:
    if not all_materials:
        return []
    mat_list = '\n'.join(f"{m['name']}|{m['unit']}|{m['qty']}|{m['price']}" for m in all_materials)
    msg = _client().messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=SYS_MAT_MARKET,
        messages=[{'role':'user','content':'Материалы (название|ед|кол|смет.цена):\n'+mat_list}]
    )
    raw = msg.content[0].text.strip().replace('```json','').replace('```','').strip()
    try:
        prices = repair_json(raw)
        if not isinstance(prices, list):
            prices = list(prices.values()) if isinstance(prices, dict) else []
        while len(prices) < len(all_materials):
            prices.append({"p_vat":0,"p":0,"src":"нет данных"})
        return prices[:len(all_materials)]
    except Exception:
        return [{"p_vat":round(m['price']*1.1),"p":m['price'],"src":"оценка"} for m in all_materials]


def get_work_market_rates(all_positions: list) -> list:
    if not all_positions:
        return []
    pos_list = []
    for p in all_positions:
        qty = p.get('qty') or 1
        ot_per = round(p.get('ot_total',0)/qty, 2) if qty else 0
        pos_list.append(f"{p['name']}|{p.get('unit','')}|{qty}|ОТ:{ot_per}")
    msg = _client().messages.create(
        model=MODEL, max_tokens=MAX_TOKENS, system=SYS_WORK_MARKET,
        messages=[{'role':'user','content':'Позиции (название|ед|объём|ОТ_на_ед):\n'+'\n'.join(pos_list)}]
    )
    raw = msg.content[0].text.strip().replace('```json','').replace('```','').strip()
    try:
        rates = repair_json(raw)
        if not isinstance(rates, list):
            rates = []
        while len(rates) < len(all_positions):
            rates.append({"r":0,"src":"нет данных"})
        return rates[:len(all_positions)]
    except Exception:
        result = []
        for p in all_positions:
            qty = p.get('qty') or 1
            ot_per = round(p.get('ot_total',0)/qty, 2) if qty else 0
            result.append({"r":round(ot_per*1.8),"src":"оценка"})
        return result


def get_market_analysis(smeta_data: dict, mat_prices: list, work_rates: list) -> dict:
    st   = smeta_data['smeta_structure']
    secs = smeta_data['sections']
    mat_idx = 0
    for sec in secs:
        for mat in sec.get('materials',[]):
            if mat_idx < len(mat_prices):
                mp = mat_prices[mat_idx]
                mat['market_price_vat'] = mp.get('p_vat',0)
                mat['market_price']     = mp.get('p',0)
                mat['market_total']     = round(mat['qty'] * mp.get('p',0))
                mat['economy']          = mat['total'] - mat['market_total']
                mat['market_source']    = mp.get('src','')
            mat_idx += 1
    pos_idx = 0
    for sec in secs:
        for pos in sec.get('positions',[]):
            if pos_idx < len(work_rates):
                wr = work_rates[pos_idx]
                pos['market_rate']  = wr.get('r',0)
                pos['market_total'] = round((pos.get('qty') or 0) * wr.get('r',0))
                pos['work_source']  = wr.get('src','')
            pos_idx += 1
    market_sections = []
    for sec in secs:
        mat_smeta  = sum(m['total']              for m in sec.get('materials',[]))
        mat_market = sum(m.get('market_total',0) for m in sec.get('materials',[]))
        wrk_market = sum(p.get('market_total',0) for p in sec.get('positions', []))
        market_sections.append({
            'name':              sec['name'],
            'mat_smeta':         mat_smeta,
            'mat_market_no_vat': mat_market,
            'work_market':       wrk_market,
            'economy_mat':       mat_smeta - mat_market,
            'economy_work':      sum(p.get('ot_total',0) for p in sec.get('positions',[])) - wrk_market,
        })
    TOTAL      = st.get('total_no_vat',0)
    mat_tot    = sum(s['mat_market_no_vat'] for s in market_sections)
    wrk_tot    = sum(s['work_market']       for s in market_sections)
    otm = st.get('otm_machines',0); em = st.get('machine_op',0); tr = st.get('transport',0)
    net_eco    = TOTAL - mat_tot - wrk_tot - otm - em - tr
    net_eco_p  = round(net_eco/TOTAL*100, 1) if TOTAL else 0
    smeta_data['market'] = {
        'note':            'petrovich/lemanapro ÷1.22 (НДС 22%), otdelka-spb.ru, май 2026',
        'vat_divisor':     1.22,
        'materials_no_vat':mat_tot,
        'works_total':     wrk_tot,
        'sections':        market_sections,
        'overhead_saved':  st.get('overhead',0),
        'profit_saved':    st.get('profit',0),
        'extra_work_cost': wrk_tot - st.get('ot_workers',0),
        'net_economy':     round(net_eco),
        'net_economy_pct': net_eco_p,
    }
    return smeta_data
