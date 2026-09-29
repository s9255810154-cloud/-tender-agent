"""
Смета — Поликлиника №5, уборка помещений 2027 (НМЦК 29 000 000 ₽ с НДС).

СТРУКТУРА (стандарт, зафиксированный пользователем для ВСЕХ смет):
  срок контракта -> цена за месяц -> цена за год -> численность персонала
  -> расходные материалы (сопоставление с прайсами ВСЕГДА, количество —
  формулой от норматива расхода) -> остальные детали.
"""
from datetime import date
import json

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

from models import TZItem
from price_parser import load_almin_price_list, load_general_opt_price_list
from resolver import resolve_all
from supplier_fallback import stub_supplier_adapter

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=12)
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
SAFE_FILL = PatternFill("solid", fgColor="C6EFCE")
DANGER_FILL = PatternFill("solid", fgColor="FFC7CE")
TOTAL_FONT = Font(name=FONT_NAME, bold=True, size=12)
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="666666")
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY_FMT = '#,##0.00 ₽;(#,##0.00 ₽);-'

with open("medical_periodic_cost.json", encoding="utf-8") as f:
    periodic = json.load(f)
ZP_PERIODIC = periodic["zp_year"]
MR_PERIODIC = periodic["mr_year"]

AREA_SQM = 16026.9
CONTRACT_MONTHS = 12
NMCK_WITH_VAT = 29_000_000.0
VAT_RATE = 0.22
NMCK_WITHOUT_VAT = NMCK_WITH_VAT / (1 + VAT_RATE)
NR_PCT, SP_PCT = 0.70, 0.10
FOT_PCT_DAILY, NR_PCT_D, SP_PCT_D = 0.302, 0.20, 0.05
EMISS_SALARY = 79181.10

periodic_total_no_vat = ZP_PERIODIC * (1 + NR_PCT + SP_PCT) + MR_PERIODIC
remaining_for_daily = NMCK_WITHOUT_VAT - periodic_total_no_vat
daily_multiplier = 1 + FOT_PCT_DAILY + NR_PCT_D + SP_PCT_D
zp_daily_base = remaining_for_daily / daily_multiplier
staff_implied = round(zp_daily_base / (EMISS_SALARY * 12))
zp_daily_fot = zp_daily_base * FOT_PCT_DAILY
floor = ZP_PERIODIC + MR_PERIODIC + zp_daily_base + zp_daily_fot
max_safe_discount = 1 - floor / NMCK_WITHOUT_VAT

wb = Workbook()
ws = wb.active
ws.title = "Смета"
ws.sheet_view.showGridLines = False
for col, width in zip("ABCDEFGH", [55, 30, 16, 14, 14, 16, 14, 14]):
    ws.column_dimensions[col].width = width

row = 1
ws.cell(row=row, column=1, value="Смета — Поликлиника №5, уборка помещений, 2027 год").font = Font(
    name=FONT_NAME, bold=True, size=14
)
row += 1
ws.cell(row=row, column=1, value=f"Дата: {date.today().isoformat()}").font = Font(name=FONT_NAME, size=10)
row += 2
ws.cell(row=row, column=1, value="Жёлтые ячейки — допущения/нормативы, эксперт может скорректировать").fill = INPUT_FILL
ws.cell(row=row, column=1).font = Font(name=FONT_NAME, size=9)
row += 2

# ============================== 1. Срок контракта ==============================
ws.cell(row=row, column=1, value="1. Срок контракта").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="Срок оказания услуг, мес")
r_months = row
ws.cell(row=row, column=2, value=CONTRACT_MONTHS)
ws.cell(row=row, column=2).fill = INPUT_FILL
row += 1
ws.cell(row=row, column=1, value="Площадь объекта, м² (3 корпуса, см. Приложение 3)")
r_area = row
ws.cell(row=row, column=2, value=AREA_SQM)
row += 2
for r in range(row - 3, row - 1):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 2. Цена за месяц / за год ==============================
ws.cell(row=row, column=1, value="2. Себестоимость").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="Твёрдый пол за весь контракт (без НДС)")
r_floor = row
ws.cell(row=row, column=2, value=floor).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="СЕБЕСТОИМОСТЬ В МЕСЯЦ").font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
r_per_month = row
ws.cell(row=row, column=2, value=f"=B{r_floor}/B{r_months}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
row += 1
ws.cell(row=row, column=1, value="СЕБЕСТОИМОСТЬ В ГОД").font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
ws.cell(row=row, column=2, value=f"=B{r_per_month}*12").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
row += 2
for r in range(row - 4, row - 1):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 3. Численность персонала ==============================
ws.cell(row=row, column=1, value="3. Численность персонала").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="Подразумеваемая базовая ЗП ежедневного штата, ₽/год")
ws.cell(row=row, column=2, value=round(zp_daily_base, 2)).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="Оклад, ₽/мес (ЕМИСС/Росстат)")
r_salary = row
ws.cell(row=row, column=2, value=EMISS_SALARY).number_format = CURRENCY_FMT
ws.cell(row=row, column=2).fill = INPUT_FILL
row += 1
ws.cell(row=row, column=1, value="ЧИСЛЕННОСТЬ ПЕРСОНАЛА").font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
r_staff = row
ws.cell(row=row, column=2, value=staff_implied)
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
ws.cell(row=row, column=3, value="человек").font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
row += 1
ws.cell(row=row, column=1,
        value="Плюс периодические работы (мытьё окон/стен и т.п., см. раздел 5) — выполняются этим же "
              "штатом по мере необходимости.").font = NOTE_FONT
row += 2
for r in range(row - 5, row - 1):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 4. Расходные материалы ==============================
ws.cell(row=row, column=1, value="4. Расходные материалы").font = SECTION_FONT
row += 1

_MATERIALS_NORMS = [
    ("Хозяйственное мыло твердое 72%", "person", 0.5, "шт/чел/мес"),
    ("Хозяйственное мыло жидкое 72%", "area", 0.01, "л/100м²/мес"),
    ("Пакеты для мусора", "person", 20, "шт/чел/мес"),
    ("Перчатки резиновые", "person", 2, "пар/чел/мес"),
    ("Средство для чистки хромированного металла", "area", 0.02, "л/100м²/мес"),
    ("Средство для удаления отложений", "area", 0.01, "л/100м²/мес"),
    ("Жидкое мыло для тела и рук", "area", 0.05, "л/100м²/мес"),
    ("Мыло туалетное (без запаха)", "person", 0.5, "шт/чел/мес"),
    ("Бумага туалетная с тиснением", "area", 2.0, "рул/100м²/мес"),
    ("Салфетки бумажные", "area", 1.0, "уп/100м²/мес"),
    ("Жидкое санитарно-гигиеническое средство для унитазов, раковин", "area", 0.03, "л/100м²/мес"),
    ("Средство для удаления неприятных запахов", "area", 0.01, "л/100м²/мес"),
    ("Чистящий порошок антибактериальный", "area", 0.02, "кг/100м²/мес"),
    ("Средство для дезинфекции", "area", 0.05, "л/100м²/мес"),
    ("Очиститель труб и жироуловителей", "area", 0.005, "л/100м²/мес"),
    ("Средство для чистки ванн/раковин/унитазов", "area", 0.03, "л/100м²/мес"),
    ("Чистящее средство для стеклянных поверхностей", "area", 0.02, "л/100м²/мес"),
    ("Моющее средство для офисной мебели", "area", 0.01, "л/100м²/мес"),
    ("Средство для удаления жевательной резинки", "area", 0.002, "л/100м²/мес"),
    ("Моющее средство для полов, не требующее смывания", "area", 0.05, "л/100м²/мес"),
    ("Нейтральное моющее средство для керамических полов", "area", 0.05, "л/100м²/мес"),
    ("Чистящее средство для оргтехники", "area", 0.005, "л/100м²/мес"),
    ("Средство для очистки тканевых поверхностей", "area", 0.005, "л/100м²/мес"),
    ("Порошковый шампунь для ковролина", "area", 0.005, "кг/100м²/мес"),
    ("Отбеливающее средство", "area", 0.01, "л/100м²/мес"),
    ("Бумажные полотенца", "area", 2.0, "уп/100м²/мес"),
    ("Тряпки для мытья полов, моющие насадки", "person", 1.0, "шт/чел/мес"),
    ("Щётки, совки, скребки, вёдра и т.п.", "person", 0.2, "шт/чел/мес"),
]

_tz_items = [TZItem(raw_name=n, unit="шт") for n, *_ in _MATERIALS_NORMS]
_price_lists = load_almin_price_list(discount_pct=5.0) + load_general_opt_price_list(supplier="ТК Сервис")
_match_results = resolve_all(_tz_items, _price_lists, supplier_adapters={"almin": stub_supplier_adapter})
_matches_by_name = {r.tz_item.raw_name: r for r in _match_results}

headers_mat = ["Наименование материала (по ТЗ)", "Найдено в прайсе", "Поставщик", "Норматив (ред.)", "База", "Кол-во/мес", "Цена, ₽", "Сумма/мес, ₽"]
for c, h in enumerate(headers_mat, start=1):
    ws.cell(row=row, column=c, value=h)
    ws.cell(row=row, column=c).fill = HEADER_FILL
    ws.cell(row=row, column=c).font = HEADER_FONT
    ws.cell(row=row, column=c).border = BORDER
row += 1

mat_first_row = row
for name, basis, norm_value, norm_unit in _MATERIALS_NORMS:
    r = _matches_by_name[name]
    ws.cell(row=row, column=1, value=name)
    matched_name = r.matched_item.name[:33] if r.matched_item else "НЕ НАЙДЕНО — искать поставщика"
    ws.cell(row=row, column=2, value=f"{matched_name} (conf {r.confidence:.2f})" if r.matched_item else matched_name)
    ws.cell(row=row, column=3, value=r.matched_item.supplier if r.matched_item else "—")

    norm_cell = ws.cell(row=row, column=4, value=norm_value)
    norm_cell.fill = INPUT_FILL
    ws.cell(row=row, column=5, value=norm_unit)

    basis_val = f"B{r_staff}" if basis == "person" else f"B{r_area}/100"
    qty_cell = ws.cell(row=row, column=6, value=f"=D{row}*{basis_val}")
    qty_cell.number_format = "0.00"

    price = r.matched_item.final_price if r.matched_item else 0
    price_cell = ws.cell(row=row, column=7, value=price)
    price_cell.number_format = CURRENCY_FMT
    if not r.matched_item:
        price_cell.fill = INPUT_FILL

    sum_cell = ws.cell(row=row, column=8, value=f"=F{row}*G{row}")
    sum_cell.number_format = CURRENCY_FMT

    if r.matched_item and r.confidence < 0.75:
        ws.cell(row=row, column=1).font = Font(name=FONT_NAME, size=9, italic=True, color="C00000")
    for c in range(1, 9):
        ws.cell(row=row, column=c).border = BORDER
        if c != 1:
            ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=9)
    row += 1
mat_last_row = row - 1

ws.cell(row=row, column=1, value="ИТОГО материалы в месяц").font = Font(name=FONT_NAME, bold=True)
r_mat_month = row
ws.cell(row=row, column=8, value=f"=SUM(H{mat_first_row}:H{mat_last_row})").number_format = CURRENCY_FMT
ws.cell(row=row, column=8).font = Font(name=FONT_NAME, bold=True)
row += 1
ws.cell(row=row, column=1, value="ИТОГО материалы в год")
ws.cell(row=row, column=8, value=f"=H{r_mat_month}*12").number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1,
        value=f"Сопоставление с прайсами — автоматически, всегда ({sum(1 for r in _match_results if r.matched_item)}/"
              f"{len(_MATERIALS_NORMS)} найдено в Альмин/ТК Сервис). Количество — ФОРМУЛА: норматив (колонка "
              f"'Норматив', жёлтая) × база (площадь/100 или численность) — меняете норматив, количество и сумма "
              f"пересчитываются сами. Нормативы ориентировочные — нет утверждённого стандарта под медицинские "
              f"категории, порядок величин взят по аналогии с consumable_norms.json для обычных помещений.").font = NOTE_FONT
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
row += 2

# ============================== 5. Периодические операции ==============================
ws.cell(row=row, column=1, value="5. Периодические операции (официальные расценки СН-2012)").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="ЗП (база, из расценок × реальные площади/количества × допущенная частота)")
r_zp_p = row
ws.cell(row=row, column=2, value=ZP_PERIODIC).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="Материалы (из расценок)")
r_mr_p = row
ws.cell(row=row, column=2, value=MR_PERIODIC).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="НР (70%) + СП (10%) от ЗП — конвенция расценок")
r_markup_p = row
ws.cell(row=row, column=2, value=f"=B{r_zp_p}*{NR_PCT + SP_PCT}").number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="Итого периодическая часть").font = Font(name=FONT_NAME, bold=True)
r_periodic_total = row
ws.cell(row=row, column=2, value=f"=B{r_zp_p}+B{r_mr_p}+B{r_markup_p}").number_format = CURRENCY_FMT
ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True)
row += 2
for r in (r_zp_p, r_mr_p, r_markup_p, r_periodic_total):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

# ============================== 6. НМЦК и сценарии скидки ==============================
ws.cell(row=row, column=1, value="6. НМЦК и сценарии скидки").font = SECTION_FONT
row += 1
ws.cell(row=row, column=1, value="НМЦК с НДС")
r_nmck_vat = row
ws.cell(row=row, column=2, value=NMCK_WITH_VAT).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="НМЦК без НДС")
r_nmck = row
ws.cell(row=row, column=2, value=NMCK_WITHOUT_VAT).number_format = CURRENCY_FMT
row += 1
ws.cell(row=row, column=1, value="Максимальная безопасная скидка от НМЦК")
ws.cell(row=row, column=2, value=f"=1-B{r_floor}/B{r_nmck}").number_format = "0.0%"
row += 2
for r in (r_nmck_vat, r_nmck):
    for c in (1, 2):
        ws.cell(row=r, column=c).border = BORDER

headers2 = ["Скидка от НМЦК", "Цена с НДС", "Цена без НДС", "Твёрдый пол", "Запас/Дефицит", "Статус"]
for c, h in enumerate(headers2, start=1):
    ws.cell(row=row, column=c, value=h)
    ws.cell(row=row, column=c).fill = HEADER_FILL
    ws.cell(row=row, column=c).font = HEADER_FONT
    ws.cell(row=row, column=c).border = BORDER
row += 1
scenario_first = row
for discount in (0.0, 0.05, 0.10, 0.15, round(max_safe_discount, 3), 0.25, 0.30, 0.35):
    ws.cell(row=row, column=1, value=discount).number_format = "0.0%"
    ws.cell(row=row, column=2, value=NMCK_WITH_VAT * (1 - discount)).number_format = CURRENCY_FMT
    ws.cell(row=row, column=3, value=f"=B{row}/1.22").number_format = CURRENCY_FMT
    ws.cell(row=row, column=4, value=f"=$B${r_floor}").number_format = CURRENCY_FMT
    ws.cell(row=row, column=5, value=f"=C{row}-D{row}").number_format = CURRENCY_FMT
    ws.cell(row=row, column=6, value=f'=IF(E{row}>=0,"безопасно","ДЕФИЦИТ")')
    for c in range(1, 7):
        ws.cell(row=row, column=c).border = BORDER
    row += 1
scenario_last = row - 1
for r in range(scenario_first, scenario_last + 1):
    d = ws.cell(row=r, column=1).value
    fill = SAFE_FILL if d <= max_safe_discount + 0.001 else DANGER_FILL
    for c in range(1, 7):
        ws.cell(row=r, column=c).fill = fill

row += 1
ws.cell(row=row, column=1,
        value=f"Подразумеваемый норматив площади на человека: {AREA_SQM/staff_implied:,.1f} м²/чел — сравните с "
              f"вашим диапазоном 800-1200 м²/чел. Ежедневный штат — ПОДРАЗУМЕВАЕМЫЙ из НМЦК, не из отдельного "
              f"расчёта заказчика (в отличие от Лужников).").font = NOTE_FONT
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)

wb.save("smeta_medical_discount.xlsx")
print("Смета сохранена: smeta_medical_discount.xlsx")
print(f"Численность: {staff_implied} чел | Твёрдый пол: {floor:,.2f} руб | Безопасная скидка: {max_safe_discount*100:.1f}%")
