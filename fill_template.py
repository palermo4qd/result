"""
Заполнение шаблона Глава 2.docx исходными данными из бухгалтерской отчётности
ООО "РЕГИОНАЛЬНАЯ СПЕЦИАЛИЗИРОВАННАЯ КОМПАНИЯ" (РЕГИОНСПЕЦКОМ).

Все числовые значения в исходных файлах фактически выражены в РУБЛЯХ
(несмотря на формальную пометку "Единица измерения: тыс. руб." на формах).
Значения, переносимые в таблицы шаблона (где явно указано "тыс. руб."),
конвертированы делением на 1000 с округлением до целых тыс. руб.

Расчётные показатели (рентабельность, фондоотдача, структура %, средняя ЗП и т.п.)
НЕ заполняются — только исходные данные. Не найденные значения помечаются «н/д».
"""
from copy import deepcopy
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

SRC = '/home/ubuntu/work/source.docx'
DST_DOC = '/home/ubuntu/repos/result/ГЛАВА 2 (заполненный).docx'
DST_NOTES = '/home/ubuntu/repos/result/Замечания.md'
DST_REPORT = '/home/ubuntu/repos/result/Отчёт_по_заполнению.md'

# ---------- ИСТОЧНИКИ ----------
# Все значения ниже — в РУБЛЯХ из исходных файлов.
# При вставке в таблицы шаблона значения в "тыс. руб." делятся на 1000.

# Балансовые показатели по коду (на 31.12 года)
BAL = {
    1100: {2021: 0,        2022: 0,         2023: 271313.49,    2024: 12140486.76,  2025: 18653808.24},
    1150: {2021: 0,        2022: 0,         2023: 271313.49,    2024: 12140486.76,  2025: 18586638.24},
    1170: {2023: 0,        2024: 0,         2025: 67170},
    1200: {2021: 5704000,  2022: 5491000,   2023: 65137413.69,  2024: 406874005.84, 2025: 507935645.36},
    1210: {2021: 0,        2022: 0,         2023: 34215322.30,  2024: 69513031.37,  2025: 229299298.28},
    1220: {2023: 404482.35, 2024: 175797.45, 2025: 2872.27},
    1230: {2021: None,     2022: 3124000,   2023: 21430763.27,  2024: 278455567.66, 2025: 264196381.07},
    1240: {2023: 0,        2024: 0,         2025: 0},
    1250: {2021: 1150000,  2022: 582000,    2023: 6895136.09,   2024: 56531640.69,  2025: 13959425.89},
    1260: {2021: None,     2022: 1785000,   2023: 2191709.68,   2024: 2197968.67,   2025: 477667.85},
    1300: {2021: 2745000,  2022: 2811000,   2023: 5911328.59,   2024: 6734309.69,   2025: 17481404.29},
    1310: {2022: 10000,    2023: 10000,     2024: 10000,        2025: 20000},
    1370: {2022: 2801000,  2023: 5901328.59, 2024: 6724309.69,  2025: 17461404.29},
    1400: {2023: 0,        2024: 6585216.29, 2025: 995500},
    1410: {2023: 0,        2024: 721000,    2025: 995500},
    1450: {2024: 5864216.29},
    1500: {2021: 2959000,  2022: 2680000,   2023: 60387398.59,  2024: 405694966.62, 2025: 504384703.55},
    1520: {2021: 2959000,  2022: 2680000,   2023: 59497398.59,  2024: 405694966.62, 2025: 504373689.29},
    1600: {2021: 5704000,  2022: 5491000,   2023: 65408727.18,  2024: 419014492.60, 2025: 526589453.60},
}

# Отчёт о фин. результатах (за период)
OFR = {
    2110: {2021: None, 2022: 36856000,    2023: 227035993.98, 2024: 269631858.62, 2025: 399379068.92},
    2120: {2021: None, 2022: -35478000,   2023: -223348132.85, 2024: -266611472.08, 2025: -383610253.52},
    2100: {2022: None, 2023: 3688000,     2024: 3021000,      2025: 15769000},
    2210: {2023: 0,    2024: 0,           2025: 0},
    2220: {2023: 0,    2024: 0,           2025: -1136742.27},
    2200: {2023: 3688000, 2024: 3021000,  2025: 14632000},
    2310: {2023: 0,    2024: 0,           2025: 0},
    2320: {2023: 0,    2024: 0,           2025: 466.64},
    2330: {2023: 0,    2024: -106931.57,  2025: -658331.89},
    2340: {2022: 0,    2023: 248227.48,   2024: 292700.80,    2025: 2327.78},
    2350: {2022: -451000, 2023: -65862.57, 2024: -1697000.99, 2025: -2400131.07},
    2300: {2023: 3870000, 2024: 1510000,  2025: 11576000},
    2410: {2022: 3014000, 2023: -585982,  2024: -686174,      2025: -829309.99},
    2400: {2021: None, 2022: 3941000,     2023: 3100000,      2024: 824000,       2025: 10747000},
}

# Затраты по элементам (Пояснения 6 — 2024 файл, Пояснения 10 — 2025 файл)
COSTS = {
    'amort':   {2023: 45286.51,   2024: 259019.78,  2025: 2290099.71},
    'mat':     {2023: 9516039.68, 2024: 8057368.39, 2025: 517997000},
    'zp':      {2023: 16375664.47, 2024: 16311415.91, 2025: 18346841.22},
    'svf':     {2023: 4334274.01, 2024: 3271528.45, 2025: 4232969.18},
    'other':   {2023: 228182190.48, 2024: 272845222.62, 2025: 0},
    'total':   {2023: 258453000,  2024: 300744555.15, 2025: 542866737.62},  # Итого по элементам
    'sebes':   {2023: 223348000,  2024: 266611472.08, 2025: 384746995.79},  # Расходы по обыч. видам
}

# ОФ по группам (первоначальная стоимость, на конец года) — Пояснения 4 (2025) / Пояснения 2 (2024)
# офисное оборудование - "Машины и оборудование", транспортные средства - "Транспортные средства"
OS_GROUP = {
    'office':    {2023: 316600,   2024: 316600,   2025: 316600},        # Офисное оборудование (Машины и оборудование)
    'transport': {2023: 0,        2024: 12128193.05, 2025: 12128193.05}, # Транспортные средства
    'total':     {2023: 316600,   2024: 12444793.05, 2025: 12444793.05}, # ОФ всего, первон. ст-сть
}

# Численность сотрудников — данных в источниках нет
HEADCOUNT = {2021: None, 2022: None, 2023: None, 2024: None, 2025: None}

YEARS = [2021, 2022, 2023, 2024, 2025]
NA = 'н/д'

def fmt_thousand(v):
    """Конвертация рублей в тыс. руб. (округление до 1 тыс.)."""
    if v is None:
        return NA
    if isinstance(v, str):
        return v
    val = round(v / 1000)
    s = f"{int(val):,}".replace(',', ' ')
    return s

def fmt_thousand_signed(v):
    if v is None:
        return NA
    return fmt_thousand(v)

def fmt_neg_to_pos(v):
    """Себестоимость и затраты: показываем модулем."""
    if v is None:
        return NA
    return fmt_thousand(abs(v))

# ---------- Запись значений в ячейки ----------
def set_cell(table, r, c, value):
    """Записывает текст в ячейку, сохраняя стиль первого run и параграфа."""
    cell = table.rows[r].cells[c]
    # Очистка содержимого, кроме первого параграфа
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)
    p = cell.paragraphs[0]
    # удаляем все runs, оставляя один
    for run in p.runs:
        run._element.getparent().remove(run._element)
    new_run = p.add_run(str(value))
    # лёгкое выравнивание по центру для числовых
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # стандартный размер шрифта (как в шаблоне, чаще всего 12pt)
    new_run.font.size = Pt(12)

def is_empty_cell(cell):
    return cell.text.strip() == ''

# ---------- Основная логика заполнения ----------
def fill_doc():
    doc = docx.Document(SRC)
    tables = doc.tables

    fill_log = []  # для отчёта

    # ===== Таблица 0: Стоимость и состав ОФ (тыс. руб.) =====
    # Структура: R0 шапка | R1 Здания | R2 Машины | R3 Транспорт | R4 И т.п. | R5 Прочие | R6 Всего
    # Колонки: C0 название | C1=2021 | C2=2022 | C3=2023 | C4=2024 | C5=2025
    t = tables[0]
    # Здания, сооружения — данных по группе нет → 0 (отсутствуют)
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, '0')  # Здания, сооружения — отсутствуют
        fill_log.append((f'Т0.R1.C{cidx}', f'Здания/сооружения {y}', 'Пояснения 4 (2025) / Пояснения 2 (2024)', '0'))
        # Машины и оборудование — Офисное оборудование
        v = OS_GROUP['office'].get(y)
        set_cell(t, 2, cidx, fmt_thousand(v) if v is not None else (NA if y in (2021,2022) else fmt_thousand(0)))
        fill_log.append((f'Т0.R2.C{cidx}', f'Машины и обор. {y}', 'Пояснения 4.1/2.1 (Офисн. обор.)', fmt_thousand(v) if v is not None else NA))
        # Транспорт
        v = OS_GROUP['transport'].get(y)
        set_cell(t, 3, cidx, fmt_thousand(v) if v is not None else NA)
        fill_log.append((f'Т0.R3.C{cidx}', f'Транспорт {y}', 'Пояснения 4.1/2.1 (Транспортные ср-ва)', fmt_thousand(v) if v is not None else NA))
        # И т.п. → 0
        set_cell(t, 4, cidx, '0')
        # Прочие → 0
        set_cell(t, 5, cidx, '0')
        # Всего
        v = OS_GROUP['total'].get(y)
        set_cell(t, 6, cidx, fmt_thousand(v) if v is not None else (NA if y in (2021,2022) else fmt_thousand(0)))
        fill_log.append((f'Т0.R6.C{cidx}', f'ОФ всего {y}', 'Пояснения 4.1/2.1 (первон. ст-ть)', fmt_thousand(v) if v is not None else NA))
    # Для 2021/2022 — ОФ не было (1150 = 0). Группы отсутствуют → 0.
    for cidx, y in [(1,2021),(2,2022)]:
        set_cell(t, 1, cidx, '0')
        set_cell(t, 2, cidx, '0')
        set_cell(t, 3, cidx, '0')
        set_cell(t, 4, cidx, '0')
        set_cell(t, 5, cidx, '0')
        set_cell(t, 6, cidx, '0')

    # ===== Таблица 1: Структура ОФ (%) — РАСЧЁТНАЯ → не заполняем =====
    # Оставляем как есть.

    # ===== Таблица 2: ОФ — движение =====
    # R1 Стоимость на начало | R2 Поступило | R3 Выбыло | R4 Стоимость на конец | R5 Среднегод | R6 Кобн | R7 Квыб
    t = tables[2]
    # Источник: первон. ст-сть из Пояснений; для 2024-2025 поступило/выбыло из Пояснений
    # 2023: нач=0, поступ=316600, выб=0, конец=316600
    # 2024: нач=316600, поступ=12128193.05 (транспорт), выб=0, конец=12444793.05
    # 2025: нач=12444793.05, поступ=0 (по 2025 файлу 4.1), выб=0, конец=12444793.05
    moves = {
        2021: {'beg': 0, 'in': 0, 'out': 0, 'end': 0},
        2022: {'beg': 0, 'in': 0, 'out': 0, 'end': 0},
        2023: {'beg': 0, 'in': 316600, 'out': 0, 'end': 316600},
        2024: {'beg': 316600, 'in': 12128193.05, 'out': 0, 'end': 12444793.05},
        2025: {'beg': 12444793.05, 'in': 0, 'out': 0, 'end': 12444793.05},
    }
    for cidx, y in enumerate(YEARS, start=1):
        m = moves[y]
        set_cell(t, 1, cidx, fmt_thousand(m['beg']))
        set_cell(t, 2, cidx, fmt_thousand(m['in']))
        set_cell(t, 3, cidx, fmt_thousand(m['out']))
        set_cell(t, 4, cidx, fmt_thousand(m['end']))
        # Среднегод. ст-сть, Кобн, Квыб — расчётные, оставляем
        for src_label, val in [('начало',m['beg']),('поступило',m['in']),('выбыло',m['out']),('конец',m['end'])]:
            fill_log.append((f'Т2.{src_label}.{y}', f'ОФ {src_label} {y}', 'Пояснения 4.1/2.1', fmt_thousand(val)))

    # ===== Таблица 7: Показатели (Выручка, Прибыль от продаж, ОФ, Численность, …) =====
    # R1 Выручка | R2 Прибыль от продаж | R3 ОФ | R4 Численность | R5-R8 — расчётные
    t = tables[7]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, fmt_thousand(OFR[2110].get(y)))
        set_cell(t, 2, cidx, fmt_thousand(OFR[2200].get(y)))
        # Основные фонды — берём строку 1150 (балансовая ОФ + ППА)
        set_cell(t, 3, cidx, fmt_thousand(BAL[1150].get(y)))
        set_cell(t, 4, cidx, NA)  # Численность нет в источниках
        fill_log.append((f'Т7.Выручка.{y}', f'Выручка 2110 {y}', 'ОФР стр. 2110', fmt_thousand(OFR[2110].get(y))))
        fill_log.append((f'Т7.ПрибПрод.{y}', f'Прибыль от продаж 2200 {y}', 'ОФР стр. 2200', fmt_thousand(OFR[2200].get(y))))
        fill_log.append((f'Т7.ОФ.{y}', f'ОФ 1150 {y}', 'Баланс стр. 1150', fmt_thousand(BAL[1150].get(y))))

    # ===== Таблица 8: Состав оборотных средств (тыс. руб.) =====
    # R1 Запасы | R2 Деб задолженность | R3 Денежные | R4 Всего
    t = tables[8]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, fmt_thousand(BAL[1210].get(y)))
        set_cell(t, 2, cidx, fmt_thousand(BAL[1230].get(y)))
        set_cell(t, 3, cidx, fmt_thousand(BAL[1250].get(y)))
        set_cell(t, 4, cidx, fmt_thousand(BAL[1200].get(y)))
        for code, label in [(1210,'Запасы'),(1230,'Дебиторская'),(1250,'Денежные'),(1200,'Всего ОА')]:
            fill_log.append((f'Т8.{label}.{y}', f'{label} {y}', f'Баланс стр. {code}', fmt_thousand(BAL[code].get(y))))

    # ===== Таблица 9: Структура оборотных средств (тыс. руб., полная) =====
    # R1 Запасы | R2 Деб | R3 Денежн | R4 НДС | R5 Финвложения | R6 Прочие | R7 Всего=100%
    t = tables[9]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, fmt_thousand(BAL[1210].get(y)))
        set_cell(t, 2, cidx, fmt_thousand(BAL[1230].get(y)))
        set_cell(t, 3, cidx, fmt_thousand(BAL[1250].get(y)))
        set_cell(t, 4, cidx, fmt_thousand(BAL[1220].get(y)))
        set_cell(t, 5, cidx, fmt_thousand(BAL[1240].get(y)))
        set_cell(t, 6, cidx, fmt_thousand(BAL[1260].get(y)))
        # R7 итого=100% оставляем как в шаблоне

    # ===== Таблица 11: Сравнение Выручки/ОА/Чистой прибыли =====
    # R2 Выручка | R3 Стоимость оборотных активов | R4 Чистая прибыль
    # Колонки: C0 наименование | C1..C5 годы 2021..2025 | C6 изменение тыс.руб (расчёт) | C7 проц. (расчёт)
    t = tables[11]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 2, cidx, fmt_thousand(OFR[2110].get(y)))
        set_cell(t, 3, cidx, fmt_thousand(BAL[1200].get(y)))
        set_cell(t, 4, cidx, fmt_thousand(OFR[2400].get(y)))

    # ===== Таблица 13: ФОТ, численность, средняя ЗП =====
    # R1 ФОТ | R2 Численность | R3 ср. ЗП (расч.) | R4 Абс. прирост (расч.) | R5 Прирост % (расч.)
    t = tables[13]
    fot = {2021: None, 2022: None, 2023: COSTS['zp'][2023], 2024: COSTS['zp'][2024], 2025: COSTS['zp'][2025]}
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, fmt_thousand(fot[y]))
        set_cell(t, 2, cidx, NA)  # численность отсутствует
        fill_log.append((f'Т13.ФОТ.{y}', f'ФОТ {y}', 'Пояснения 6/10 — Расходы на оплату труда', fmt_thousand(fot[y])))

    # ===== Таблица 15: Выручка/Численность/Выработка =====
    t = tables[15]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, fmt_thousand(OFR[2110].get(y)))
        set_cell(t, 2, cidx, NA)
        # Выработка — расчётная

    # ===== Таблица 16: Затраты по элементам (тыс. руб.) =====
    # R1 Амортизация | R2 Материальные затраты | R3 Расходы на оплату труда | R4 СВФ | R5 И т.д. | R6,R7 пусто | R8 Прочие издержки | R9 Итого
    t = tables[16]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, fmt_thousand(COSTS['amort'].get(y)))
        set_cell(t, 2, cidx, fmt_thousand(COSTS['mat'].get(y)))
        set_cell(t, 3, cidx, fmt_thousand(COSTS['zp'].get(y)))
        set_cell(t, 4, cidx, fmt_thousand(COSTS['svf'].get(y)))
        set_cell(t, 8, cidx, fmt_thousand(COSTS['other'].get(y)))
        set_cell(t, 9, cidx, fmt_thousand(COSTS['total'].get(y)))
        for k, lbl in [('amort','Амортизация'),('mat','Материальные'),('zp','ОТ'),('svf','СВФ'),('other','Прочие'),('total','Итого')]:
            fill_log.append((f'Т16.{lbl}.{y}', f'{lbl} {y}', 'Пояснения 6 (2024) / Пояснения 10 (2025)', fmt_thousand(COSTS[k].get(y))))

    # ===== Таблица 17: Структура затрат (%) — РАСЧЁТНАЯ =====
    # Не заполняем

    # ===== Таблица 18: Выручка/ОФ/Численность + расчёты =====
    # R2 Выручка | R3 Среднегод ОФ | R4 Среднегод численность | R5 заголовок | R6 на работника (расч.) | R7 на 1 руб ОФ (расч.)
    t = tables[18]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 2, cidx, fmt_thousand(OFR[2110].get(y)))
        # Среднегод ОФ — для 2023..2025 можем посчитать как (нач+конец)/2, но это расчёт.
        # В качестве источника возьмём остаток на конец года (т.к. среднегод. — расчётная).
        # Помечаем н/д для среднегод., чтобы не считать.
        set_cell(t, 3, cidx, NA)
        set_cell(t, 4, cidx, NA)

    # ===== Таблица 24: Выручка/Себест/Валовая/Уровень/Числен/Выр на раб. =====
    t = tables[24]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 0, cidx, fmt_thousand(OFR[2110].get(y)))
        # Себестоимость со знаком минус в источнике — в таблице обычно положительное число
        sebes = OFR[2120].get(y)
        set_cell(t, 1, cidx, fmt_thousand(abs(sebes)) if sebes is not None else NA)
        set_cell(t, 2, cidx, fmt_thousand(OFR[2100].get(y)))
        # уровень валовой прибыли — расчёт; численность — н/д; выручка/работника — расчёт
        set_cell(t, 4, cidx, NA)

    # ===== Таблица 25: Подробный ОФР =====
    # R2 Выручка | R3 Себестоимость | R4 Валовая прибыль | R5 Управленческие расходы | R6 Прибыль от продаж
    # R7 Проценты к уплате | R8 Проценты к получению | R9 Прочие доходы | R10 Прочие расходы
    # R11 Прибыль до налогообложения | R12 ЕН при УСН | R13 Налог на прибыль | R14 Чистая прибыль
    t = tables[25]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 2, cidx, fmt_thousand(OFR[2110].get(y)))
        sebes = OFR[2120].get(y)
        set_cell(t, 3, cidx, fmt_thousand(abs(sebes)) if sebes is not None else NA)
        set_cell(t, 4, cidx, fmt_thousand(OFR[2100].get(y)))
        upr = OFR[2220].get(y)
        set_cell(t, 5, cidx, fmt_thousand(abs(upr)) if upr is not None else NA)
        set_cell(t, 6, cidx, fmt_thousand(OFR[2200].get(y)))
        proc_upl = OFR[2330].get(y)
        set_cell(t, 7, cidx, fmt_thousand(abs(proc_upl)) if proc_upl is not None else NA)
        set_cell(t, 8, cidx, fmt_thousand(OFR[2320].get(y)))
        set_cell(t, 9, cidx, fmt_thousand(OFR[2340].get(y)))
        proc_rash = OFR[2350].get(y)
        set_cell(t, 10, cidx, fmt_thousand(abs(proc_rash)) if proc_rash is not None else NA)
        set_cell(t, 11, cidx, fmt_thousand(OFR[2300].get(y)))
        set_cell(t, 12, cidx, NA)  # ЕН при УСН — компания на ОСН
        nal = OFR[2410].get(y)
        # Знак налога: в источниках 2023-2025 минус (расход), 2022 плюс (рестейтмент). Покажем модулем.
        set_cell(t, 13, cidx, fmt_thousand(abs(nal)) if nal is not None else NA)
        set_cell(t, 14, cidx, fmt_thousand(OFR[2400].get(y)))

    # ===== Таблица 26: Чистая прибыль динамика =====
    t = tables[26]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 1, cidx, fmt_thousand(OFR[2400].get(y)))

    # ===== Таблица 29: Выручка =====
    t = tables[29]
    for cidx, y in enumerate(YEARS, start=1):
        set_cell(t, 2, cidx, fmt_thousand(OFR[2110].get(y)))

    # ===== Таблица 30: Совокупные затраты, Прибыль от продаж =====
    # R0 Совокупные затраты | R1 Прибыль от продаж | R2/R3 рентабельность (расч.)
    t = tables[30]
    for cidx, y in enumerate(YEARS, start=1):
        # совокупные затраты (себестоимость + управленческие)
        sebes = OFR[2120].get(y)
        upr = OFR[2220].get(y)
        if sebes is not None or upr is not None:
            tot = (abs(sebes) if sebes else 0) + (abs(upr) if upr else 0)
        else:
            tot = None
        set_cell(t, 0, cidx, fmt_thousand(tot))
        set_cell(t, 1, cidx, fmt_thousand(OFR[2200].get(y)))

    # Сохраняем
    doc.save(DST_DOC)
    print(f'Saved: {DST_DOC}')
    return fill_log


if __name__ == '__main__':
    log = fill_doc()
    print(f'Total filled: {len(log)} cells (logged)')
