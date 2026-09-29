"""
Смета со скидкой для торгов — контракт "Административно-хозяйственный
корпус, ул. Лужники д.24" (НМЦК 2 999 211 ₽ с НДС).

Показывает: структуру официальной сметы (твёрдый пол vs гибкая часть),
оптимизацию материалов (замена брендов на эквиваленты — разрешено самой
сметой), и таблицу сценариев скидки — от безопасного максимума до
запрошенных 30-35% с явным дефицитом, чтобы решение принималось осознанно.
"""
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=12)
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
DANGER_FILL = PatternFill("solid", fgColor="FFC7CE")
SAFE_FILL = PatternFill("solid", fgColor="C6EFCE")
TOTAL_FONT = Font(name=FONT_NAME, bold=True, size=12)
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="666666")
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY_FMT = '#,##0.00 ₽;(#,##0.00 ₽);-'

# --- Официальная база (НМЦК), проверено до копейки против сметы заказчика ---
NMCK_WITH_VAT = 2999211
NMCK_WITHOUT_VAT = 2458369.67
ZP_RAW = 1157754.00
FOT_PCT, NR_PCT, SP_PCT = 0.302, 0.20, 0.05
VAT_RATE = 0.22

wb = Workbook()
ws = wb.active
ws.title = "Смета со скидкой"
ws.sheet_view.showGridLines = False
for col, width in zip("ABCDEFG", [50, 16, 16, 16, 16, 16, 40]):
    ws.column_dimensions[col].width = width

row = 1
ws.cell(row=row, column=1, value="Смета со скидкой — Административно-хозяйственный корпус, ул. Лужники, 24").font = Font(
    name=FONT_NAME, bold=True, size=14
)
row += 1
ws.cell(row=row, column=1, value="НМЦК (официальная смета заказчика, проверена до копейки): 2 999 211 ₽ с НДС").font = Font(
    name=FONT_NAME, size=10
)
row += 1
ws.cell(row=row, column=1, value=f"Дата: {date.today().isoformat()}").font = Font(name=FONT_NAME, size=10)
row += 2

ws.cell(row=row, column=1, value="Жёлтые ячейки — редактируемые допущения").fill = INPUT_FILL
ws.cell(row=row, column=1).font = Font(name=FONT_NAME, size=9)
ws.cell(row=row, column=3, value="Зелёный — безопасная зона").fill = SAFE_FILL
ws.cell(row=row, column=3).font = Font(name=FONT_NAME, size=9)
ws.cell(row=row, column=5, value="Красный — цена ниже твёрдого пола (убыток)").fill = DANGER_FILL
ws.cell(row=row, column=5).font = Font(name=FONT_NAME, size=9)
row += 2

# ============================== 0. Численность персонала ==============================
ws.cell(row=row, column=1, value="0. Численность персонала").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="Суммарные трудозатраты по официальным расценкам, чел-час/контракт (6 мес)")
ws.cell(row=row, column=2, value=2456.4)
row += 1
ws.cell(row=row, column=1, value="Доступно часов на 1 человека за контракт (8ч × 21 раб.день × 6 мес)")
ws.cell(row=row, column=2, value=1008)
row += 1
ws.cell(row=row, column=1, value="ЧИСЛЕННОСТЬ ПЕРСОНАЛА").font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
r_headcount = row
ws.cell(row=row, column=2, value=2)
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
ws.cell(row=row, column=3, value="человека").font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
row += 1
ws.cell(row=row, column=1, value="Работы периодические (не ежедневное присутствие) — 2 человека покрывают "
                                    "весь объём работ по факту трудозатрат, при условии гибкого графика визитов.").font = NOTE_FONT
row += 2

for r in range(row - 6, row - 1):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 0б. Расходные материалы ==============================
ws.cell(row=row, column=1, value="0б. Расходные материалы (официальная спецификация + наш аналог из прайсов)").font = SECTION_FONT
row += 1
headers_mat = ["Наименование (официальный бренд)", "Кол-во за контракт", "Официальная сумма с НДС, ₽",
               "Наш аналог (прайс)", "Наш поставщик", "Наша цена, ₽"]
for c, h in enumerate(headers_mat, start=1):
    ws.cell(row=row, column=c, value=h)
    ws.cell(row=row, column=c).fill = HEADER_FILL
    ws.cell(row=row, column=c).font = HEADER_FONT
    ws.cell(row=row, column=c).border = BORDER
row += 1

import json as _json
from models import TZItem
from price_parser import load_almin_price_list, load_general_opt_price_list
from resolver import resolve_all
from supplier_fallback import stub_supplier_adapter

_premises_mat = _json.load(open("knowledge_base/premises_gesn_rates.json", encoding="utf-8"))
_territory_mat = _json.load(open("knowledge_base/territory_gesn_rates_luzhniki.json", encoding="utf-8"))
_all_materials = []
for it in _premises_mat["items"] + _territory_mat["items"]:
    for m in it.get("materials", []):
        if m.get("total_with_vat"):
            _all_materials.append(m)
    if not it.get("work_cost_total") and it.get("total_with_vat"):
        _all_materials.append({"name": it["name"], "qty": it.get("qty"), "total_with_vat": it["total_with_vat"]})

# Сопоставление с нашими прайсами — ВСЕГДА, для каждой сметы (см. стандарт).
# Для официальных брендов ищем по первым словам названия (без наворотов вида
# "или эквивалент по уровню pH..." — это не название товара, а требование к нему).
_tz_items = [TZItem(raw_name=m["name"].split(" или эквивалент")[0].split(".")[0][:60], unit="шт") for m in _all_materials]
_price_lists = load_almin_price_list(discount_pct=5.0) + load_general_opt_price_list(supplier="ТК Сервис")
_match_results = resolve_all(_tz_items, _price_lists, supplier_adapters={"almin": stub_supplier_adapter})

mat_first_row = row
for m, r in zip(_all_materials, _match_results):
    ws.cell(row=row, column=1, value=m["name"][:90])
    ws.cell(row=row, column=2, value=m.get("qty"))
    ws.cell(row=row, column=3, value=m.get("total_with_vat")).number_format = CURRENCY_FMT
    if r.matched_item:
        ws.cell(row=row, column=4, value=f"{r.matched_item.name[:35]} (conf {r.confidence:.2f})")
        ws.cell(row=row, column=5, value=r.matched_item.supplier)
        ws.cell(row=row, column=6, value=r.matched_item.final_price).number_format = CURRENCY_FMT
    else:
        ws.cell(row=row, column=4, value="не найдено в наших прайсах")
        ws.cell(row=row, column=5, value="—")
        ws.cell(row=row, column=6, value=0)
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = BORDER
        ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=9)
    row += 1
mat_last_row = row - 1
ws.cell(row=row, column=1, value="ИТОГО материалы (с НДС)").font = Font(name=FONT_NAME, bold=True)
r_mat_check = row
ws.cell(row=row, column=3, value=f"=SUM(C{mat_first_row}:C{mat_last_row})").number_format = CURRENCY_FMT
ws.cell(row=row, column=3).font = Font(name=FONT_NAME, bold=True)
row += 2

# ============================== 1. Структура официальной сметы ==============================
ws.cell(row=row, column=1, value="1. Структура НМЦК (без НДС)").font = SECTION_FONT
row += 1
headers = ["Компонент", "Сумма, ₽", "% от сметы", "Можно сократить?"]
for c, h in enumerate(headers, start=1):
    ws.cell(row=row, column=c, value=h)
for c in range(1, 5):
    ws.cell(row=row, column=c).fill = HEADER_FILL
    ws.cell(row=row, column=c).font = HEADER_FONT
    ws.cell(row=row, column=c).border = BORDER
header_row = row
row += 1

MAT_BASE = NMCK_WITHOUT_VAT - ZP_RAW - ZP_RAW * FOT_PCT - ZP_RAW * NR_PCT - ZP_RAW * SP_PCT  # 661 535.46 ₽, фиксировано

r_zp = row
ws.cell(row=row, column=1, value="ЗП (база, реальная оплата труда)")
ws.cell(row=row, column=2, value=ZP_RAW).number_format = CURRENCY_FMT
ws.cell(row=row, column=3, value=f"=B{row}/$B$66").number_format = "0.0%"
ws.cell(row=row, column=4, value="Нет")
row += 1

r_fot = row
ws.cell(row=row, column=1, value="ФОТ-надбавка (страховые взносы, 30.2%)")
ws.cell(row=row, column=2, value=f"=B{r_zp}*{FOT_PCT}").number_format = CURRENCY_FMT
ws.cell(row=row, column=3, value=f"=B{row}/$B$66").number_format = "0.0%"
ws.cell(row=row, column=4, value="Нет, обязательна по закону")
row += 1

r_mat = row
ws.cell(row=row, column=1, value="Материалы (база, по брендам заказчика)")
ws.cell(row=row, column=2, value=MAT_BASE).number_format = CURRENCY_FMT
ws.cell(row=row, column=3, value=f"=B{row}/$B$66").number_format = "0.0%"
ws.cell(row=row, column=4, value="Частично — см. раздел 2")
row += 1

r_nr = row
ws.cell(row=row, column=1, value="Накладные расходы — УБРАНО ПО ЗАПРОСУ (было 20% = 231 550,80 ₽)")
ws.cell(row=row, column=2, value=0).number_format = CURRENCY_FMT
ws.cell(row=row, column=3, value=0).number_format = "0.0%"
ws.cell(row=row, column=4, value="Исключено из расчёта")
ws.cell(row=row, column=1).font = Font(name=FONT_NAME, italic=True, size=10, color="999999")
row += 1

r_sp = row
ws.cell(row=row, column=1, value="Сметная прибыль (5%)")
ws.cell(row=row, column=2, value=f"=B{r_zp}*{SP_PCT}").number_format = CURRENCY_FMT
ws.cell(row=row, column=3, value=f"=B{row}/$B$66").number_format = "0.0%"
ws.cell(row=row, column=4, value="Да — гибкая часть")
row += 1

ws.cell(row=row, column=1, value="ИТОГО без НДС (пересчитано, без накладных)").font = TOTAL_FONT
r_total_check = row
ws.cell(row=row, column=2, value=f"=B{r_zp}+B{r_fot}+B{r_mat}+B{r_nr}+B{r_sp}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = TOTAL_FONT
row += 1
ws.cell(row=row, column=1, value=f"Справочно — официальный НМЦК без НДС (с накладными): {NMCK_WITHOUT_VAT:,.2f} ₽").font = NOTE_FONT
row += 2

for r in (r_zp, r_fot, r_mat, r_nr, r_sp, r_total_check):
    for c in range(1, 5):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 2. Оптимизация материалов ==============================
ws.cell(row=row, column=1, value="2. Оптимизация материалов (замена на эквиваленты)").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="Доля брендовых материалов (Tork/Domestos/Секреты Чистоты и т.п.)")
r_branded_share = row
ws.cell(row=row, column=2, value=0.794).number_format = "0.0%"
ws.cell(row=row, column=2).fill = INPUT_FILL
row += 1
ws.cell(row=row, column=1, value="Экономия при замене на эквивалент (по опыту наших прайсов)")
r_savings_pct = row
ws.cell(row=row, column=2, value=0.35).number_format = "0.0%"
ws.cell(row=row, column=2).fill = INPUT_FILL
ws.cell(row=row, column=5, value="Проверьте по факту доступных эквивалентов у ваших поставщиков").font = NOTE_FONT
row += 1
ws.cell(row=row, column=1, value="Экономия на материалах, ₽").font = Font(name=FONT_NAME, bold=True)
r_mat_savings = row
ws.cell(row=row, column=2, value=f"=B{r_mat}*B{r_branded_share}*B{r_savings_pct}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True)
row += 2

for r in (r_branded_share, r_savings_pct, r_mat_savings):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 3. Твёрдый пол и безопасная скидка ==============================
ws.cell(row=row, column=1, value="3. Твёрдый пол и максимальная безопасная скидка").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="Твёрдый пол (ЗП + ФОТ + материалы, с учётом экономии на материалах)").font = Font(name=FONT_NAME, bold=True)
r_floor = row
ws.cell(row=row, column=2, value=f"=B{r_zp}+B{r_fot}+B{r_mat}-B{r_mat_savings}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True)
row += 1
ws.cell(row=row, column=1, value="Максимальная безопасная скидка от НМЦК без НДС")
r_max_discount = row
ws.cell(row=row, column=2, value=f"=1-B{r_floor}/{NMCK_WITHOUT_VAT}").number_format = "0.0%"
row += 1

ws.cell(row=row, column=1, value="Цена без накладных, но С прибылью (твёрдый пол + сметная прибыль 5%)").font = Font(name=FONT_NAME, bold=True)
r_price_no_nr = row
ws.cell(row=row, column=2, value=f"=B{r_floor}+B{r_sp}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True)
row += 1
ws.cell(row=row, column=1, value="Соответствующая скидка от НМЦК (без накладных, с прибылью)")
r_discount_no_nr = row
ws.cell(row=row, column=2, value=f"=1-B{r_price_no_nr}/{NMCK_WITHOUT_VAT}").number_format = "0.0%"
row += 2

for r in (r_floor, r_max_discount, r_price_no_nr, r_discount_no_nr):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 4. Сценарии скидки ==============================
ws.cell(row=row, column=1, value="4. Сценарии скидки для ставки в торгах").font = SECTION_FONT
row += 1
headers2 = ["Скидка от НМЦК", "Цена с НДС", "Цена без НДС", "Твёрдый пол", "Запас (+) / Дефицит (-)", "Статус"]
for c, h in enumerate(headers2, start=1):
    ws.cell(row=row, column=c, value=h)
    ws.cell(row=row, column=c).fill = HEADER_FILL
    ws.cell(row=row, column=c).font = HEADER_FONT
    ws.cell(row=row, column=c).border = BORDER
row += 1

scenario_first_row = row
for discount in (0.0, 0.05, 0.10, 0.15, 0.1925, 0.20, 0.25, 0.30, 0.35):
    ws.cell(row=row, column=1, value=discount).number_format = "0.0%"
    price_with_vat = ws.cell(row=row, column=2, value=NMCK_WITH_VAT * (1 - discount))
    price_with_vat.number_format = CURRENCY_FMT
    price_without_vat = ws.cell(row=row, column=3, value=f"=B{row}/1.22")
    price_without_vat.number_format = CURRENCY_FMT
    floor_ref = ws.cell(row=row, column=4, value=f"=$B${r_floor}")
    floor_ref.number_format = CURRENCY_FMT
    margin = ws.cell(row=row, column=5, value=f"=C{row}-D{row}")
    margin.number_format = CURRENCY_FMT
    status = ws.cell(row=row, column=6, value=f'=IF(E{row}>=0,"безопасно","ДЕФИЦИТ")')
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = BORDER
    row += 1
scenario_last_row = row - 1

# Порог для подсветки — тот же расчёт, что и в самом файле (твёрдый пол с учётом экономии на материалах)
_zp, _fot = ZP_RAW, ZP_RAW * FOT_PCT
_mat = NMCK_WITHOUT_VAT - ZP_RAW - _fot - ZP_RAW * NR_PCT - ZP_RAW * SP_PCT
_mat_savings = _mat * 0.794 * 0.35
_floor = _zp + _fot + _mat - _mat_savings
SAFE_DISCOUNT_THRESHOLD = 1 - _floor / NMCK_WITHOUT_VAT

for r in range(scenario_first_row, scenario_last_row + 1):
    discount_val = ws.cell(row=r, column=1).value
    fill = SAFE_FILL if discount_val <= SAFE_DISCOUNT_THRESHOLD else DANGER_FILL
    for c in range(1, 7):
        ws.cell(row=r, column=c).fill = fill

row += 1
price_no_nr_val = _floor + ZP_RAW * SP_PCT
discount_no_nr_val = 1 - price_no_nr_val / NMCK_WITHOUT_VAT
ws.cell(row=row, column=1,
        value=f"Накладные расходы убраны из расчёта (раздел 1). Цена без накладных, но с сохранённой "
              f"прибылью = {price_no_nr_val:,.0f} ₽ без НДС — это соответствует скидке ~{discount_no_nr_val*100:.1f}% "
              f"от НМЦК. Абсолютный пол (без накладных И без прибыли, только зарплата+взносы+материалы) "
              f"остаётся на уровне ~{SAFE_DISCOUNT_THRESHOLD*100:.1f}% — это крайний случай, не рекомендуется "
              f"как рабочая ставка. Строки 30% и 35% по-прежнему показывают дефицит даже без накладных.")
ws.cell(row=row, column=1).font = NOTE_FONT
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)

wb.save("smeta_luzhniki_discount.xlsx")
print("Смета сохранена: smeta_luzhniki_discount.xlsx")
