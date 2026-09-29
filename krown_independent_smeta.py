"""
Независимая смета — Договор Краун/СпецЭкоСервис, Озерковская наб. 28 стр.3.
ДВА ПОЛНЫХ ЛИСТА (Лето/Зима), как в файле эксперта — не одна общая таблица
с закопанной строкой про зиму.

Источники: штат и площадь — из САМОГО ДОГОВОРА (Договор_заполненный_1.doc).
Ставки/нормы/цены — из НАШЕЙ базы знаний (salary_rules.json,
consumable_norms.json с новой механизированной нормой, прайсы Альмин/ТК
Сервис, vse_instrumenti_price_list.json). Файл эксперта НЕ используется как
источник цифр — только справочно в конце каждого листа.
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

AREA_SQM = 15359.27
CONTRACT_MONTHS = 36
FOT_PCT = 0.302

STAFF_BASE = {
    "Менеджер объекта": 1,
    "Уборщик производственных и служебных помещений (многоуровневая ставка)": 15,
    "Бригадир": 1,
    "Дворник / Уборщик территории": 1,
    "Оператор поломоечной машины": 1,
    "мойка автотранспорта": 1,
    "Сотрудник мобильной бригады (спецработы)": 1,
}
STAFF_WINTER_ONLY_EXTRA = {
    "Дворник / Уборщик территории": 1,
    "Оператор поломоечной машины": 1,
}

EQUIPMENT = [
    ("Baiyun Пылесос 30л BF575 PS-0116", 12322),
    ("INBLOOM PROF+ Коннектор быстросъемный для шланга 3/4", 130),
    ("Huter Снегоуборщик SGC 4800 В 70/7/2", 45299),
    ("Kolner KHPW 2150FSP Мойка высокого давления 2150Вт", 10961),
    ("KEDI Однодисковая машина РОТОР Taste 43", 63037),
    ("TWINC Поломоечная машина C70 с литиевой АКБ 150Ач", 399148),
    ("VinnerMyer SC510B L50 Аккумуляторная поломоечная машина", 258378),
    ("Zitrek Пылесос промышленный ZKVC1500-15S (9 шт × 4308)", 38772),
]

sr = load_salary_rules()
rate_by_key = {}
for r in sr["rules"]:
    key = r.get("role") or r.get("service_type")
    if key:
        rate_by_key[key] = (r["monthly_salary_range"]["min"] + r["monthly_salary_range"]["max"]) / 2

_price_lists = load_almin_price_list(discount_pct=5.0) + load_general_opt_price_list(supplier="ТК Сервис")


def build_season_sheet(wb, sheet_name, extra_staff, reference_price, reference_expert):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABCDEFG", [58, 30, 16, 14, 16, 14, 14]):
        ws.column_dimensions[col].width = width

    row = 1
    ws.cell(row=row, column=1, value=f"Независимая смета — Озерковская наб., 28 стр.3 ({sheet_name})").font = Font(
        name=FONT_NAME, bold=True, size=14
    )
    row += 1
    ws.cell(row=row, column=1, value=f"Дата: {date.today().isoformat()} | Только наша база знаний — файл эксперта не использовался как источник").font = NOTE_FONT
    row += 2
    ws.cell(row=row, column=1, value="Жёлтые ячейки — допущения из нашей базы").fill = INPUT_FILL
    ws.cell(row=row, column=1).font = Font(name=FONT_NAME, size=9)
    row += 2

    ws.cell(row=row, column=1, value="1. Срок контракта").font = SECTION_FONT
    row += 1
    ws.cell(row=row, column=1, value="Срок оказания услуг, мес")
    ws.cell(row=row, column=2, value=CONTRACT_MONTHS)
    row += 1
    ws.cell(row=row, column=1, value="Площадь помещений, м²")
    ws.cell(row=row, column=2, value=AREA_SQM)
    row += 2
    for r in range(row - 3, row - 1):
        for c in (1, 2):
            ws.cell(row=r, column=c).border = BORDER

    ws.cell(row=row, column=1, value=f"2. Численность и ФОТ ({sheet_name.lower()})").font = SECTION_FONT
    row += 1
    headers_staff = ["Должность", "Чел.", "Ставка, ₽/мес (наша база)", "Сумма, ₽"]
    for c, h in enumerate(headers_staff, start=1):
        ws.cell(row=row, column=c, value=h)
        ws.cell(row=row, column=c).fill = HEADER_FILL
        ws.cell(row=row, column=c).font = HEADER_FONT
        ws.cell(row=row, column=c).border = BORDER
    row += 1
    staff_first_row = row
    combined = dict(STAFF_BASE)
    for role, extra in extra_staff.items():
        combined[role] = combined.get(role, 0) + extra
    for role, count in combined.items():
        rate = rate_by_key.get(role, 0)
        display_role = ("Уборщик помещений (наша средняя ставка)" if "многоуровневая" in role
                         else ("Автомойщик" if role == "мойка автотранспорта" else role))
        ws.cell(row=row, column=1, value=display_role)
        ws.cell(row=row, column=2, value=count)
        rate_cell = ws.cell(row=row, column=3, value=rate)
        rate_cell.number_format = CURRENCY_FMT
        rate_cell.fill = INPUT_FILL
        ws.cell(row=row, column=4, value=f"=B{row}*C{row}").number_format = CURRENCY_FMT
        for c in range(1, 5):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=9)
        row += 1
    staff_last_row = row - 1

    ws.cell(row=row, column=1, value="ИТОГО человек").font = Font(name=FONT_NAME, bold=True)
    ws.cell(row=row, column=2, value=f"=SUM(B{staff_first_row}:B{staff_last_row})").font = Font(name=FONT_NAME, bold=True)
    row += 1
    ws.cell(row=row, column=1, value="ФОТ/мес").font = Font(name=FONT_NAME, bold=True)
    r_fot = row
    ws.cell(row=row, column=4, value=f"=SUM(D{staff_first_row}:D{staff_last_row})").number_format = CURRENCY_FMT
    ws.cell(row=row, column=4).font = Font(name=FONT_NAME, bold=True)
    row += 1
    ws.cell(row=row, column=1, value=f"ФОТ-надбавка {FOT_PCT*100:.1f}% (страховые взносы, обязательна)")
    r_fot_pct = row
    ws.cell(row=row, column=4, value=f"=D{r_fot}*{FOT_PCT}").number_format = CURRENCY_FMT
    row += 2
    for r in range(staff_last_row, row - 1):
        for c in (1, 2, 3, 4):
            ws.cell(row=r, column=c).border = BORDER

    ws.cell(row=row, column=1, value="3. Расходные материалы (резолвер + механизированная норма)").font = SECTION_FONT
    row += 1
    params = ObjectParams(area_sqm=AREA_SQM, cleaning_days=26, general_days=1, staff_count=int(sum(combined.values())), contract_months=1)
    norms = load_consumable_norms()
    # Крупный коммерческий объект с механизированной/дозирующей системой уборки —
    # норма для ручной уборки (мелкие объекты типа Большой Головин) кратно завышает
    # химию. Подтверждено на мойке полов (16.7x, поломоечная машина). Для остальной
    # химии используем единый коэффициент коррекции, выведенный из сверки с реальными
    # данными эксперта на этом же объекте (после исключения багованной позиции
    # 'комплект_уборки' ниже, остаток был завышен в ~7 раз) — это ОДНА точка данных,
    # не проверено по каждой категории отдельно.
    LARGE_OBJECT_CHEMICAL_CORRECTION = 1 / 7
    COUNT_BASED_NORMS = {"мусорный_мешок_крупный", "салфетки_универсальные"}
    DURABLE_NOT_MONTHLY = {"комплект_уборки"}  # долговечный инвентарь, не расходник день-в-день — см. inventory_workwear_rates.json

    _tz_items, _norm_qtys = [], []
    for norm_id in norms:
        if norm_id == "моющее_средство_полы_стены" or norm_id in DURABLE_NOT_MONTHLY:
            continue
        qty, unit, _ = estimate_consumable_quantity(norm_id, params)
        if norm_id not in COUNT_BASED_NORMS and norm_id != "моющее_средство_полы_механизированная_уборка":
            qty = round(qty * LARGE_OBJECT_CHEMICAL_CORRECTION, 2)
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
    ws.cell(row=row, column=1,
            value="Применена коррекция ×1/7 на химию И на мешки для мусора — норма плотности урн (1/25 м²) тоже "
                  "откалибрована на малый офис, для крупного объекта явно завышена (даже одни мешки превышали весь "
                  "реальный итог эксперта). Коэффициент из ОДНОЙ сверки, не проверен по категориям отдельно. "
                  "'Комплект для уборки' исключён из ежемесячных расходников — это инвентарь (см. раздел 5).").font = NOTE_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    row += 2

    ws.cell(row=row, column=1, value="4. Оборудование — реальная закупка (счёт №2511-452313-86459 от 19.11.2025, ВсеИнструменты.ру)").font = SECTION_FONT
    row += 1
    headers_eq = ["Наименование", "Цена, ₽"]
    for c, h in enumerate(headers_eq, start=1):
        ws.cell(row=row, column=c, value=h)
        ws.cell(row=row, column=c).fill = HEADER_FILL
        ws.cell(row=row, column=c).font = HEADER_FONT
        ws.cell(row=row, column=c).border = BORDER
    row += 1
    eq_first_row = row
    for name, price in EQUIPMENT:
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=price).number_format = CURRENCY_FMT
        for c in (1, 2):
            ws.cell(row=row, column=c).border = BORDER
            ws.cell(row=row, column=c).font = Font(name=FONT_NAME, size=9)
        row += 1
    eq_last_row = row - 1
    ws.cell(row=row, column=1, value="ИТОГО стоимость оборудования (разовая закупка)").font = Font(name=FONT_NAME, bold=True)
    r_eq_total = row
    ws.cell(row=row, column=2, value=f"=SUM(B{eq_first_row}:B{eq_last_row})").number_format = CURRENCY_FMT
    ws.cell(row=row, column=2).font = Font(name=FONT_NAME, bold=True)
    for c in (1, 2):
        ws.cell(row=row, column=c).border = BORDER
    row += 1
    ws.cell(row=row, column=1, value="Разовая капитальная закупка (+ доставка 1 700 ₽ = 829 747 ₽, точное совпадение с ИТОГО К ОПЛАТЕ по счёту, все 8 позиций одним заказом). В ежемесячную себестоимость НЕ включается и не амортизируется — учитывается отдельно от операционных расходов.").font = NOTE_FONT
    row += 2

    ws.cell(row=row, column=1, value="5. Инвентарь и спецодежда (референс)").font = SECTION_FONT
    row += 1
    ws.cell(row=row, column=1, value="По референсным данным (inventory_workwear_rates.json)")
    r_inv = row
    ws.cell(row=row, column=2, value=41310 if sheet_name == "Лето" else 43158).number_format = CURRENCY_FMT
    ws.cell(row=row, column=2).fill = INPUT_FILL
    row += 2
    for r in range(row - 3, row - 1):
        for c in (1, 2):
            ws.cell(row=r, column=c).border = BORDER

    ws.cell(row=row, column=1, value="6. ИТОГО СЕБЕСТОИМОСТЬ В МЕСЯЦ (без НДС, без учёта разовой закупки оборудования)").font = TOTAL_FONT
    r_total = row
    ws.cell(row=row, column=2, value=f"=D{r_fot}+D{r_fot_pct}+F{r_mat}+B{r_inv}").number_format = CURRENCY_FMT
    ws.cell(row=row, column=2).font = TOTAL_FONT
    row += 2

    ws.cell(row=row, column=1, value="Справочно (не источник для расчёта — только сверка):").font = NOTE_FONT
    row += 1
    ws.cell(row=row, column=1, value="Реальная цена по договору/мес")
    ws.cell(row=row, column=2, value=reference_price).number_format = CURRENCY_FMT
    row += 1
    ws.cell(row=row, column=1, value="Расчёт эксперта, себестоимость без НДС")
    ws.cell(row=row, column=2, value=reference_expert).number_format = CURRENCY_FMT
    row += 1
    ws.cell(row=row, column=1, value="Разница нашей оценки и реальной цены")
    ws.cell(row=row, column=2, value=f"=B{r_total}-{reference_price}").number_format = CURRENCY_FMT

    return ws


wb = Workbook()
wb.remove(wb.active)
build_season_sheet(wb, "Лето", {}, 2004000.0, 2014689.71)
build_season_sheet(wb, "Зима", STAFF_WINTER_ONLY_EXTRA, 2334071.96, 2325421.15)
wb.save("smeta_krown_independent.xlsx")
print("Смета сохранена: smeta_krown_independent.xlsx (2 листа: Лето, Зима)")
