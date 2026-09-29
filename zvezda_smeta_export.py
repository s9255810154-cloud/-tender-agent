"""
Независимая смета — Договор ООО «Киностудия ЗВЕЗДА» / ООО «СпецЭкоСервис»,
КТК «Главкино», Красногорск. Срок 11.11.2025-10.11.2026 (12 мес).

Реальные данные ИЗ ДОГОВОРА: площадь по этажам (табл.2), штатное расписание
по сменам (табл.3), список оборудования (табл.5), реальная цена контракта
(23 299 999,92 ₽ с НДС 5% за год). Ставки/нормы — из нашей базы знаний.
"""
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

from knowledge_base import ObjectParams, estimate_consumable_quantity, load_consumable_norms, load_salary_rules
from models import TZItem
from price_parser import load_almin_price_list, load_general_opt_price_list
from resolver import resolve_all
from supplier_fallback import stub_supplier_adapter
from unit_price import estimate_purchase_quantity

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=12)
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
TOTAL_FONT = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="666666")
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY_FMT = '#,##0.00 ₽;(#,##0.00 ₽);-'

CONTRACT_MONTHS = 12
REAL_PRICE_TOTAL = 23299999.92
AREA_COMPLEX_SUPPORT = 5456.6
AREA_ON_DEMAND = 15216.0
AREA_TERRITORY = 20299.0
AREA_LAWN = 17780.0
FOT_PCT = 0.302

STAFF = {
    "Уборщик производственных и служебных помещений (многоуровневая ставка)": {"count": 10, "note": "3 (7-8ч) + 5 (8-19ч) + 2 (19-20ч)"},
    "Дворник / Уборщик территории": {"count": 5, "note": "круглогодично"},
}
STAFF_WINTER_ONLY = {
    "Тракторист-механизатор": {"count": 1, "rate": 95000, "note": "01.11-31.03, ОЦЕНКА ставки — нет в нашей базе"},
}

sr = load_salary_rules()
rate_by_key = {}
for r in sr["rules"]:
    key = r.get("role") or r.get("service_type")
    if key:
        rate_by_key[key] = (r["monthly_salary_range"]["min"] + r["monthly_salary_range"]["max"]) / 2

_price_lists = load_almin_price_list(discount_pct=5.0) + load_general_opt_price_list(supplier="ТК Сервис")

wb = Workbook()
ws = wb.active
ws.title = "Смета"
ws.sheet_view.showGridLines = False
for col, width in zip("ABCDEFG", [58, 30, 16, 14, 16, 14, 14]):
    ws.column_dimensions[col].width = width

row = 1
ws.cell(row=row, column=1, value="Независимая смета — КТК «Главкино» (Киностудия ЗВЕЗДА / СпецЭкоСервис)").font = Font(
    name=FONT_NAME, bold=True, size=14
)
row += 1
ws.cell(row=row, column=1, value=f"Дата: {date.today().isoformat()} | Площадь/штат — реальные данные из договора; ставки/нормы — наша база знаний").font = NOTE_FONT
row += 2
ws.cell(row=row, column=1, value="Жёлтые ячейки — допущения").fill = INPUT_FILL
ws.cell(row=row, column=1).font = Font(name=FONT_NAME, size=9)
row += 2

ws.cell(row=row, column=1, value="1. Срок контракта и площадь").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="Срок оказания услуг, мес")
r_months = row
ws.cell(row=row, column=2, value=CONTRACT_MONTHS)
row += 1
ws.cell(row=row, column=1, value="Площадь комплексной+поддерживающей уборки (этажи 1-4), м²")
r_area_main = row
ws.cell(row=row, column=2, value=AREA_COMPLEX_SUPPORT)
row += 1
ws.cell(row=row, column=1, value="Площадь по заявкам (вспом.помещения, тех.этаж, павильоны, склады), м²")
r_area_demand = row
ws.cell(row=row, column=2, value=AREA_ON_DEMAND)
row += 1
ws.cell(row=row, column=1, value="Территория (асфальт/брусчатка), м²")
r_area_territory = row
ws.cell(row=row, column=2, value=AREA_TERRITORY)
row += 1
ws.cell(row=row, column=1, value="Газоны, м²")
r_area_lawn = row
ws.cell(row=row, column=2, value=AREA_LAWN)
row += 2
for r in range(row - 6, row - 1):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

ws.cell(row=row, column=1, value="2. Численность и ФОТ (реальное штатное расписание)").font = SECTION_FONT
row += 1
headers_staff = ["Должность", "Чел.", "Ставка, ₽/мес (наша база)", "Сумма, ₽"]
for c, h in enumerate(headers_staff, start=1):
    ws.cell(row=row, column=c, value=h)
    ws.cell(row=row, column=c).fill = HEADER_FILL
    ws.cell(row=row, column=c).font = HEADER_FONT
    ws.cell(row=row, column=c).border = BORDER
row += 1
staff_first_row = row
for role, info in STAFF.items():
    rate = rate_by_key.get(role, 0)
    display_role = "Уборщик помещений (наша средняя ставка)" if "многоуровневая" in role else role
    ws.cell(row=row, column=1, value=f"{display_role} ({info['note']})")
    ws.cell(row=row, column=2, value=info["count"])
    rate_cell = ws.cell(row=row, column=3, value=rate)
    rate_cell.number_format = CURRENCY_FMT
    rate_cell.fill = INPUT_FILL
    ws.cell(row=row, column=4, value=f"=B{row}*C{row}").number_format = CURRENCY_FMT
    for c in range(1, 5):
        ws.cell(row=row, column=c).border = BORDER
        ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=9)
    row += 1
staff_last_row = row - 1

ws.cell(row=row, column=1, value="ФОТ/мес (круглогодичный штат)").font = Font(name=FONT_NAME, bold=True)
r_fot = row
ws.cell(row=row, column=4, value=f"=SUM(D{staff_first_row}:D{staff_last_row})").number_format = CURRENCY_FMT
ws.cell(row=row, column=4).font = Font(name=FONT_NAME, bold=True)
row += 1
ws.cell(row=row, column=1, value="ФОТ-надбавка 30,2% (обязательна)")
r_fot_pct = row
ws.cell(row=row, column=4, value=f"=D{r_fot}*{FOT_PCT}").number_format = CURRENCY_FMT
row += 1
role, info = next(iter(STAFF_WINTER_ONLY.items()))
ws.cell(row=row, column=1, value=f"{role} ({info['note']})")
ws.cell(row=row, column=2, value=info["count"])
rate_cell = ws.cell(row=row, column=3, value=info["rate"])
rate_cell.number_format = CURRENCY_FMT
rate_cell.fill = INPUT_FILL
r_tractorist = row
ws.cell(row=row, column=4, value=f"=B{row}*C{row}").number_format = CURRENCY_FMT
for c in range(1, 5):
    ws.cell(row=row, column=c).border = BORDER
    ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=9)
row += 1
ws.cell(row=row, column=1, value="ФОТ тракториста за контракт (5 зимних месяцев)")
r_tractorist_total = row
ws.cell(row=row, column=4, value=f"=D{r_tractorist}*5").number_format = CURRENCY_FMT
row += 2
for r in range(staff_last_row, row - 1):
    for c in (1, 2, 3, 4):
        ws.cell(row=r, column=c).border = BORDER

ws.cell(row=row, column=1, value="3. Расходные материалы (резолвер + механизированная норма + коррекция ×1/7)").font = SECTION_FONT
row += 1
TOTAL_AREA_FOR_MATERIALS = AREA_COMPLEX_SUPPORT + AREA_ON_DEMAND
params = ObjectParams(area_sqm=TOTAL_AREA_FOR_MATERIALS, cleaning_days=26, general_days=1,
                       staff_count=sum(v["count"] for v in STAFF.values()), contract_months=1)
norms = load_consumable_norms()
LARGE_OBJECT_CORRECTION = 1 / 7
COUNT_BASED_NORMS = {"мусорный_мешок_крупный", "салфетки_универсальные"}
DURABLE_NOT_MONTHLY = {"комплект_уборки"}

_tz_items, _norm_qtys = [], []
for norm_id in norms:
    if norm_id == "моющее_средство_полы_стены" or norm_id in DURABLE_NOT_MONTHLY:
        continue
    qty, unit, _ = estimate_consumable_quantity(norm_id, params)
    if norm_id not in COUNT_BASED_NORMS and norm_id != "моющее_средство_полы_механизированная_уборка":
        qty = round(qty * LARGE_OBJECT_CORRECTION, 2)
    _tz_items.append(TZItem(raw_name=norms[norm_id]["display_name"], qty=qty, unit=unit))
    _norm_qtys.append((qty, unit))
_match_results = resolve_all(_tz_items, _price_lists, supplier_adapters={"almin": stub_supplier_adapter})

headers_mat = ["Материал", "Найдено в прайсе", "Поставщик", "Упаковок", "Цена/уп, ₽", "Сумма/мес, ₽"]
for c, h in enumerate(headers_mat, start=1):
    ws.cell(row=row, column=c, value=h)
    ws.cell(row=row, column=c).fill = HEADER_FILL
    ws.cell(row=row, column=c).font = HEADER_FONT
    ws.cell(row=row, column=c).border = BORDER
row += 1
mat_first_row = row
for r, (qty, unit) in zip(_match_results, _norm_qtys):
    ws.cell(row=row, column=1, value=f"{r.tz_item.raw_name[:38]} ({qty} {unit})")
    matched_name = r.matched_item.name[:26] if r.matched_item else "НЕ НАЙДЕНО"
    ws.cell(row=row, column=2, value=f"{matched_name} (c{r.confidence:.2f})" if r.matched_item else matched_name)
    ws.cell(row=row, column=3, value=r.matched_item.supplier if r.matched_item else "—")
    packs = estimate_purchase_quantity(qty, unit, r.matched_item) if r.matched_item and unit in ("л", "кг", "шт") else None
    packs_cell = ws.cell(row=row, column=4, value=packs if packs is not None else qty)
    if packs is None:
        packs_cell.fill = INPUT_FILL
    price = r.matched_item.final_price if r.matched_item else 0
    price_cell = ws.cell(row=row, column=5, value=price)
    price_cell.number_format = CURRENCY_FMT
    if not r.matched_item:
        price_cell.fill = INPUT_FILL
    ws.cell(row=row, column=6, value=f"=D{row}*E{row}").number_format = CURRENCY_FMT
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = BORDER
        if c != 1:
            ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=9)
    row += 1
mat_last_row = row - 1
ws.cell(row=row, column=1, value="ИТОГО материалы/мес").font = Font(name=FONT_NAME, bold=True)
r_mat = row
ws.cell(row=row, column=6, value=f"=SUM(F{mat_first_row}:F{mat_last_row})").number_format = CURRENCY_FMT
ws.cell(row=row, column=6).font = Font(name=FONT_NAME, bold=True)
row += 1
ws.cell(row=row, column=1, value="Считано по внутренней площади (комплексная+по заявке = 20 672,6 м²), без территории/газонов.").font = NOTE_FONT
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
row += 2

ws.cell(row=row, column=1, value="4. Оборудование (список из договора — разовая закупка, не в себестоимости/мес)").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1,
        value="Тележки×6, швабры×6, пылесос×1, поломоечная машина аккум.×1 (до 4000 м²/заряд), тачки садовые×3, "
              "лопаты/движки/ледорубы (комплект), ручной инвентарь (комплект), газонокосилки×2, триммеры×3, "
              "снегоуборщик самоходный×1, трактор с отвалом+щёткой×1 (ноя-мар). Не пересчитано детально в этой смете.")
ws.cell(row=row, column=1).font = NOTE_FONT
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
row += 2

ws.cell(row=row, column=1, value="5. Территория — справочно (не включена в смету)").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="20 299 м² асфальт/брусчатка + 17 780 м² газонов — периодичность и нормы для этого "
                                    "масштаба не откалиброваны так же подробно, как для помещений. Дворники (5 чел.) "
                                    "уже учтены в ФОТ выше — повторно не добавляются.")
ws.cell(row=row, column=1).font = NOTE_FONT
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
row += 2

ws.cell(row=row, column=1, value="6. ИТОГО СЕБЕСТОИМОСТЬ В МЕСЯЦ (без НДС, без оборудования/территории отдельно)").font = TOTAL_FONT
r_total_month = row
ws.cell(row=row, column=2, value=f"=D{r_fot}+D{r_fot_pct}+F{r_mat}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = TOTAL_FONT
row += 1
ws.cell(row=row, column=1, value="ИТОГО за контракт (12 мес, круглогодичная часть + зимний тракторист)").font = TOTAL_FONT
r_total_year = row
ws.cell(row=row, column=2, value=f"=B{r_total_month}*12+D{r_tractorist_total}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = TOTAL_FONT
row += 2

ws.cell(row=row, column=1, value="Справочно (не источник для расчёта — только сверка):").font = NOTE_FONT
row += 1
ws.cell(row=row, column=1, value="Реальная цена по договору за весь срок (с НДС 5%)")
ws.cell(row=row, column=2, value=REAL_PRICE_TOTAL).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="Реальная цена в месяц (для сравнения)")
ws.cell(row=row, column=2, value=REAL_PRICE_TOTAL / 12).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="Разница нашей оценки (за контракт) и реальной цены")
ws.cell(row=row, column=2, value=f"=B{r_total_year}-{REAL_PRICE_TOTAL}").number_format = CURRENCY_FMT

wb.save("smeta_zvezda_independent.xlsx")
print("Смета сохранена: smeta_zvezda_independent.xlsx")
