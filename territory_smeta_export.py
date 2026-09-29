"""
Смета по ТЗ "уборка территории" (второй тип, ГБУ ГЦПиКР) — модель "аренда
техники по вызову", как решил пользователь.

Честно исключено из сметы: "Санитарное содержание (уборка)", "Полив",
"Мойка", "Очистка урны" — для них осталась нерешённая неоднозначность
интерпретации площади (см. чат: проверка на 91 дворника провалилась для
"Санитарное содержание", а мелкие категории похожи на неполное извлечение).
Включены только "Уборка снега" и "Посыпка ПГМ" — по ним есть официальный
триггер (762-ПП) и климатическая база (Росгидромет), достаточные для
осмысленного первого прохода.
"""
from datetime import date

from knowledge_base import load_real_salary_benchmarks
from smeta_export import BORDER, CURRENCY_FMT, FONT_NAME, HEADER_FILL, HEADER_FONT, INPUT_FILL, NOTE_FONT, TOTAL_FONT, _style_header_row
from openpyxl import Workbook
from openpyxl.styles import Font

# --- Данные, извлечённые из реального ТЗ (2027 год — чистый, полный год) ---
AREA_SNOW_SQM = 153640.2       # "Уборка снега", 2027
AREA_PGM_SQM = 33608.0          # "Посыпка противогололедными материалами", 2027
CONTRACT_MONTHS = 23            # 01.01.2027 - 30.11.2028
CONTRACT_YEARS = CONTRACT_MONTHS / 12

# --- Климатическая база (Росгидромет, метеостанция Москва ВДНХ) ---
SNOW_EVENTS_PER_YEAR = 116      # дней в году со снегом — см. territory_norms.json

# --- Норматив (762-ПП + отраслевые источники) ---
PGM_NORM_G_PER_SQM = 60

# --- Аренда техники "по вызову" ---
SHIFT_COST_RUB = 5400            # трактор с оператором, 8ч смена, всё включено
# Теоретическая производительность щёточного оборудования — 27800 м²/час
# (заводская характеристика). В городских условиях (повороты, препятствия,
# не открытое поле) реальная производительность ниже — берём консервативную
# оценку, редактируемую в самой смете.
REAL_PRODUCTIVITY_SQM_PER_SHIFT_DEFAULT = 40000

# --- Ручная бригада (дворники, детали — урны, кромки, недоступные технике места) ---
DVORNIK_STAFF_COUNT = 2
# Из real_salary_benchmarks.json — берём ближайший по графику ориентир
benchmarks = load_real_salary_benchmarks()
dvornik_5_2_10h = next(
    (b for b in benchmarks if b["category"] == "дворник" and b["workdays_per_week"] == 5 and b["daily_hours"] == 10),
    None,
)
DVORNIK_SALARY_DEFAULT = dvornik_5_2_10h["avg_salary"] if dvornik_5_2_10h else 60000


def export_territory_smeta(output_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Смета (территория)"
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABCDE", [55, 20, 14, 14, 40]):
        ws.column_dimensions[col].width = width

    row = 1
    ws.cell(row=row, column=1, value="Смета — Уборка территории ГБУ г. Москвы «ГЦПиКР»").font = Font(
        name=FONT_NAME, bold=True, size=14
    )
    row += 1
    ws.cell(row=row, column=1, value="Модель: аренда техники по вызову (не свой тракторист)").font = Font(
        name=FONT_NAME, size=10, italic=True
    )
    row += 1
    ws.cell(row=row, column=1, value=f"Дата формирования: {date.today().isoformat()}").font = Font(name=FONT_NAME, size=10)
    row += 2

    ws.cell(row=row, column=1, value="Жёлтые ячейки — ключевые допущения, отредактируйте под реальные условия").fill = INPUT_FILL
    ws.cell(row=row, column=1).font = Font(name=FONT_NAME, size=9)
    row += 2

    def data_row(label, value, editable=False, fmt=None, note=None):
        nonlocal row
        ws.cell(row=row, column=1, value=label).font = Font(name=FONT_NAME, size=10)
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = Font(name=FONT_NAME, size=10)
        cell.border = BORDER
        if editable:
            cell.fill = INPUT_FILL
        if fmt:
            cell.number_format = fmt
        if note:
            ws.cell(row=row, column=5, value=note).font = NOTE_FONT
        row += 1
        return row - 1

    # ============================== 1. Аренда техники (снег) ==============================
    ws.cell(row=row, column=1, value="1. Аренда техники — уборка снега (по вызову)").font = Font(name=FONT_NAME, bold=True, size=12)
    row += 1

    r_area_snow = data_row("Площадь, подлежащая уборке от снега, м² (из ТЗ, 2027 — чистый год)", AREA_SNOW_SQM)
    r_shift_cost = data_row("Стоимость смены техники, ₽ (трактор+оператор, 8ч, аренда по вызову)", SHIFT_COST_RUB, editable=True,
                              fmt=CURRENCY_FMT, note="Источник: рыночная цена аренды трактора с оператором, см. чат")
    r_productivity = data_row("Реальная производительность за смену, м²/смену", REAL_PRODUCTIVITY_SQM_PER_SHIFT_DEFAULT, editable=True,
                                note="Теоретическая заводская — 27800 м²/час (222400 м²/смена); здесь — консервативная оценка с поправкой на городские условия. ПРОВЕРИТЬ с реальным подрядчиком.")
    row += 1
    ws.cell(row=row, column=1, value="Смен техники нужно на одно событие (снегопад)").font = Font(name=FONT_NAME, size=10)
    r_shifts_per_event = row
    shifts_cell = ws.cell(row=row, column=2, value=f"=ROUNDUP(B{r_area_snow}/B{r_productivity},0)")
    shifts_cell.border = BORDER
    row += 1

    r_events = data_row("Событий (дней со снегопадом) за контракт", None, editable=True,
                          note="Климатическая норма: 116 дней/год со снегом (Росгидромет) — это ВЕРХНЯЯ ГРАНИЦА; часть лёгких снегопадов может не требовать полного вызова техники. ПРОВЕРИТЬ/скорректировать по факту.")
    events_default = round(SNOW_EVENTS_PER_YEAR * CONTRACT_YEARS)
    ws.cell(row=r_events, column=2, value=events_default)
    ws.cell(row=r_events, column=2).fill = INPUT_FILL

    row += 1
    ws.cell(row=row, column=1, value="Итого смен техники за контракт").font = Font(name=FONT_NAME, size=10)
    r_total_shifts = row
    total_shifts_cell = ws.cell(row=row, column=2, value=f"=B{r_shifts_per_event}*B{r_events}")
    total_shifts_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="ИТОГО аренда техники за контракт").font = TOTAL_FONT
    r_equipment_total = row
    eq_total_cell = ws.cell(row=row, column=2, value=f"=B{r_total_shifts}*B{r_shift_cost}")
    eq_total_cell.font = TOTAL_FONT
    eq_total_cell.number_format = CURRENCY_FMT
    eq_total_cell.border = BORDER
    row += 2

    # ============================== 2. Расходные материалы (ПГМ) ==============================
    ws.cell(row=row, column=1, value="2. Противогололедные материалы").font = Font(name=FONT_NAME, bold=True, size=12)
    row += 1

    r_area_pgm = data_row("Площадь для посыпки, м² (из ТЗ, 2027)", AREA_PGM_SQM)
    r_pgm_norm = data_row("Норма расхода, г/м² за обработку", PGM_NORM_G_PER_SQM, editable=True,
                            note="Отраслевой ориентир (50-70 г/м²), не таблица 762-ПП (не раскрыта в доступном фрагменте)")
    r_pgm_events = data_row("Обработок за контракт", events_default, editable=True,
                              note="Используем ту же оценку, что и для смен техники — ПРОВЕРИТЬ отдельно, обработка ПГМ не всегда совпадает с каждым вызовом техники")
    row += 1
    ws.cell(row=row, column=1, value="Итого материала, кг").font = Font(name=FONT_NAME, size=10)
    r_pgm_kg = row
    pgm_kg_cell = ws.cell(row=row, column=2, value=f"=B{r_area_pgm}*B{r_pgm_norm}*B{r_pgm_events}/1000")
    pgm_kg_cell.number_format = "#,##0.0"
    pgm_kg_cell.border = BORDER
    row += 1

    r_pgm_price = data_row("Цена за кг, ₽", 20, editable=True,
                             note="ЗАГЛУШКА — не найдено в наших прайсах (они под клининг помещений, не стройматериалы/ПГМ). Нужен отдельный поставщик, уточнить цену.")
    row += 1

    ws.cell(row=row, column=1, value="ИТОГО материалы за контракт").font = TOTAL_FONT
    r_materials_total = row
    mat_total_cell = ws.cell(row=row, column=2, value=f"=B{r_pgm_kg}*B{r_pgm_price}")
    mat_total_cell.font = TOTAL_FONT
    mat_total_cell.number_format = CURRENCY_FMT
    mat_total_cell.border = BORDER
    row += 2

    # ============================== 3. Ручная бригада ==============================
    ws.cell(row=row, column=1, value="3. Ручная бригада (урны, кромки, места, недоступные технике)").font = Font(
        name=FONT_NAME, bold=True, size=12
    )
    row += 1
    r_staff = data_row("Численность (дворники)", DVORNIK_STAFF_COUNT, editable=True)
    r_salary = data_row("Оклад, ₽/мес", DVORNIK_SALARY_DEFAULT, editable=True, fmt=CURRENCY_FMT,
                          note=f"Ближайший реальный ориентир из Табель.xlsx: 5/2, 10ч -> {DVORNIK_SALARY_DEFAULT:,} ₽ (выборка n=2, слабая — ПРОВЕРИТЬ)")
    r_months = data_row("Срок контракта, мес", CONTRACT_MONTHS, editable=True)
    row += 1
    ws.cell(row=row, column=1, value="ИТОГО ФОТ ручной бригады за контракт").font = TOTAL_FONT
    r_labor_total = row
    labor_total_cell = ws.cell(row=row, column=2, value=f"=B{r_staff}*B{r_salary}*B{r_months}")
    labor_total_cell.font = TOTAL_FONT
    labor_total_cell.number_format = CURRENCY_FMT
    labor_total_cell.border = BORDER
    row += 2

    # ============================== Итог ==============================
    ws.cell(row=row, column=1, value="ИТОГО СЕБЕСТОИМОСТЬ (техника + материалы + ручная бригада)").font = Font(
        name=FONT_NAME, bold=True, size=13, color="C00000"
    )
    final_cell = ws.cell(row=row, column=2, value=f"=B{r_equipment_total}+B{r_materials_total}+B{r_labor_total}")
    final_cell.number_format = CURRENCY_FMT
    final_cell.font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
    row += 2

    ws.cell(row=row, column=1,
            value="НЕ ВКЛЮЧЕНО в смету (не найдено надёжной интерпретации площади — см. чат): "
                  "«Санитарное содержание (уборка)» (304 992,6 м² — не прошло проверку на здравый смысл), "
                  "«Полив», «Мойка», «Очистка урны» (данные по ним из ТЗ извлечены не полностью). "
                  "Наценка и НДС не добавлены — эта смета только себестоимость техники/материалов/бригады по снегу.")
    ws.cell(row=row, column=1).font = NOTE_FONT
    row += 1
    ws.cell(row=row, column=1,
            value="Все жёлтые ячейки — допущения агента, требуют проверки эксперта/подрядчика прежде чем "
                  "смета уйдёт заказчику.")
    ws.cell(row=row, column=1).font = NOTE_FONT

    wb.save(output_path)
