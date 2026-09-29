"""
Экспорт итоговой сметы в Excel (.xlsx).

Структура книги:
  - Лист "Смета" — сводная таблица: трудозатраты, расходные материалы
    (из resolver.resolve_all), накладные/маржа, итог. Все суммы — формулами,
    не захардкоженными числами, чтобы смета пересчитывалась при правке
    жёлтых ячеек.
  - Лист "Альтернативы" — для позиций, где выбор между поставщиками был
    неочевиден (score близкие, или сравнение цены за единицу ненадёжно) —
    полный список кандидатов, чтобы эксперт мог быстро проверить решение
    агента, не переоткрывая прайсы.

Цветовая легенда (соответствует принятой в xlsx-скилле схеме):
  жёлтый  — ячейки для ввода/проверки экспертом (оклад, численность, наценка,
            и любая позиция с requires_expert_review=True)
  красный текст — cумма по накладным ещё не подтверждена
  обычный чёрный — расчётные формулы, эксперт не трогает
"""
from dataclasses import dataclass
from datetime import date
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from models import MatchResult
from unit_price import estimate_purchase_quantity

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=12)
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")   # эксперт вводит/проверяет
REVIEW_FILL = PatternFill("solid", fgColor="FFC7CE")  # требует подтверждения (низкая уверенность)
TOTAL_FONT = Font(name=FONT_NAME, bold=True, size=12)
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="666666")
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY_FMT = '#,##0.00 ₽;(#,##0.00 ₽);-'


@dataclass
class LaborCostInput:
    """
    Модель "комбинированная смена": один сотрудник за одну смену закрывает
    и ежедневную-основную, и ежедневную-поддерживающую уборку (данные о
    часах — из графика в самом ТЗ, это факт, не оценка). Базовая ставка за
    смену и решение об официальном оформлении — на усмотрение эксперта,
    поэтому обе ячейки редактируемые; разбивка по видам уборки считается
    формулой по факту часов, а не вводится вручную.
    """
    hours_daily_main: float          # ежедневная-основная, ч/день — из графика ТЗ
    hours_daily_supporting: float    # ежедневная-поддерживающая, ч/день — из графика ТЗ
    base_shift_salary: float         # ставка за смену, руб/мес (эксперт: диапазон 75-85 тыс, не выше 90 тыс)
    official_employment: bool        # применять ли +43.5% (НДФЛ+взносы) — по усмотрению эксперта
    official_overhead_pct: float     # 43.5 по умолчанию, тоже редактируемо в смете
    staff_count_recommended: int
    contract_duration_months: int
    staff_count_by_area: int = 0
    staff_count_by_hours: int = 0


@dataclass
class SmetaMeta:
    object_name: str
    region: str
    tender_url: Optional[str] = None
    prepared_date: date = None

    def __post_init__(self):
        if self.prepared_date is None:
            self.prepared_date = date.today()


def _style_header_row(ws: Worksheet, row: int, n_cols: int) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _apply_body_style(ws: Worksheet, row: int, n_cols: int) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name=FONT_NAME, size=10)
        cell.border = BORDER


def export_smeta_to_xlsx(
    consumables: list[MatchResult],
    labor: LaborCostInput,
    meta: SmetaMeta,
    output_path: str,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Смета"
    ws.sheet_view.showGridLines = False

    for col, width in zip("ABCDEFGHIJ", [40, 14, 8, 40, 22, 14, 12, 14, 10, 32]):
        ws.column_dimensions[col].width = width

    row = 1
    ws.cell(row=row, column=1, value="Смета — " + meta.object_name).font = Font(
        name=FONT_NAME, bold=True, size=14
    )
    row += 1
    ws.cell(row=row, column=1, value=f"Регион: {meta.region}").font = Font(name=FONT_NAME, size=10)
    row += 1
    ws.cell(row=row, column=1, value=f"Дата формирования: {meta.prepared_date.isoformat()}").font = \
        Font(name=FONT_NAME, size=10)
    if meta.tender_url:
        row += 1
        ws.cell(row=row, column=1, value=f"Тендер: {meta.tender_url}").font = Font(name=FONT_NAME, size=10)
    row += 2

    # --- Легенда ---
    ws.cell(row=row, column=1, value="Жёлтые ячейки — ввод/проверка экспертом").fill = INPUT_FILL
    ws.cell(row=row, column=1).font = Font(name=FONT_NAME, size=9)
    ws.cell(row=row, column=3, value="Розовые ячейки — низкая уверенность сопоставления, нужна проверка").fill = REVIEW_FILL
    ws.cell(row=row, column=3).font = Font(name=FONT_NAME, size=9)
    row += 2

    # ============================== 1. Трудозатраты ==============================
    ws.cell(row=row, column=1, value="1. Трудозатраты (комбинированная смена: ежедневная-основная + поддерживающая)").font = SECTION_FONT
    row += 1

    def labor_row(label: str, value, editable: bool = False, fmt: Optional[str] = None):
        nonlocal row
        ws.cell(row=row, column=1, value=label).font = Font(name=FONT_NAME, size=10)
        cell = ws.cell(row=row, column=2, value=value)
        cell.font = Font(name=FONT_NAME, size=10)
        cell.border = BORDER
        if editable:
            cell.fill = INPUT_FILL
        if fmt:
            cell.number_format = fmt
        row += 1
        return row - 1

    r_hours_main = labor_row("Часы в смене: ежедневная-основная (из графика ТЗ)", labor.hours_daily_main)
    r_hours_supp = labor_row("Часы в смене: ежедневная-поддерживающая (из графика ТЗ)", labor.hours_daily_supporting)
    row += 1
    ws.cell(row=row, column=1, value="Итого часов в смене").font = Font(name=FONT_NAME, size=10)
    r_hours_total = row
    hrs_total_cell = ws.cell(row=row, column=2, value=f"=B{r_hours_main}+B{r_hours_supp}")
    hrs_total_cell.border = BORDER
    row += 2

    r_base_shift = labor_row(
        "Базовая ставка за смену, руб/мес (эксперт: диапазон 75-85 тыс, не выше 90 тыс)",
        labor.base_shift_salary, editable=True, fmt=CURRENCY_FMT,
    )
    r_official = labor_row(
        "Официальное оформление в штат? (1 = да, 0 = нет — решает эксперт)",
        1 if labor.official_employment else 0, editable=True,
    )
    r_overhead_pct = labor_row(
        "Надбавка при официальном оформлении, % (НДФЛ+страховые взносы+травматизм)",
        labor.official_overhead_pct / 100, editable=True, fmt="0.0%",
    )

    row += 1
    ws.cell(row=row, column=1, value="Итоговая ставка за смену (с учётом оформления)").font = TOTAL_FONT
    r_shift_total = row
    # Потолок 90 000 применяется к БАЗОВОЙ ставке до надбавки за оформление;
    # надбавка (если включена) добавляется поверх уже ограниченной базы.
    shift_total_cell = ws.cell(
        row=row, column=2,
        value=f"=MIN(B{r_base_shift},90000)*(1+B{r_official}*B{r_overhead_pct})",
    )
    shift_total_cell.font = TOTAL_FONT
    shift_total_cell.number_format = CURRENCY_FMT
    shift_total_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  из них — ФОТ на ежедневную-основную (доля по часам)")
    r_fot_main = row
    fot_main_cell = ws.cell(row=row, column=2, value=f"=B{r_shift_total}*B{r_hours_main}/B{r_hours_total}")
    fot_main_cell.number_format = CURRENCY_FMT
    fot_main_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  из них — ФОТ на поддерживающую (доля по часам)")
    r_fot_supp = row
    fot_supp_cell = ws.cell(row=row, column=2, value=f"=B{r_shift_total}*B{r_hours_supp}/B{r_hours_total}")
    fot_supp_cell.number_format = CURRENCY_FMT
    fot_supp_cell.border = BORDER
    row += 2

    labor_row("Численность — оценка по площади объекта (не заполнено, см. knowledge_base)", labor.staff_count_by_area)
    labor_row("Численность — оценка по графику/часам (не заполнено, см. knowledge_base)", labor.staff_count_by_hours)
    r_staff = labor_row(
        "Численность — принято к расчёту (подтвердить!)", labor.staff_count_recommended, editable=True
    )
    r_months = labor_row("Срок контракта, мес", labor.contract_duration_months, editable=True)

    row += 1
    ws.cell(row=row, column=1, value="Итого ФОТ в месяц (все сотрудники)").font = TOTAL_FONT
    r_fot_month = row
    fot_month_cell = ws.cell(row=row, column=2, value=f"=B{r_shift_total}*B{r_staff}")
    fot_month_cell.font = TOTAL_FONT
    fot_month_cell.number_format = CURRENCY_FMT
    fot_month_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="Итого ФОТ в год").font = Font(name=FONT_NAME, size=10)
    r_fot_year = row
    fot_year_cell = ws.cell(row=row, column=2, value=f"=B{r_fot_month}*12")
    fot_year_cell.number_format = CURRENCY_FMT
    fot_year_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="Итого ФОТ за весь срок контракта").font = TOTAL_FONT
    r_labor_total = row
    total_cell = ws.cell(row=row, column=2, value=f"=B{r_fot_month}*B{r_months}")
    total_cell.font = TOTAL_FONT
    total_cell.number_format = CURRENCY_FMT
    total_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  из них за контракт — ежедневная-основная")
    fot_main_total_cell = ws.cell(row=row, column=2, value=f"=B{r_fot_main}*B{r_staff}*B{r_months}")
    fot_main_total_cell.number_format = CURRENCY_FMT
    fot_main_total_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  из них за контракт — поддерживающая")
    fot_supp_total_cell = ws.cell(row=row, column=2, value=f"=B{r_fot_supp}*B{r_staff}*B{r_months}")
    fot_supp_total_cell.number_format = CURRENCY_FMT
    fot_supp_total_cell.border = BORDER
    row += 2

    # ============================== 2. Расходные материалы ==============================
    ws.cell(row=row, column=1, value="2. Расходные материалы").font = SECTION_FONT
    row += 1

    headers = ["Позиция ТЗ", "Потребность (натур.)", "Ед.", "Найдено в прайсе", "Поставщик",
               "Цена за упаковку, ₽", "Упаковок к заказу", "Сумма, ₽", "Confidence", "Проверка"]
    header_row = row
    for col, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col, value=h)
    _style_header_row(ws, header_row, len(headers))
    row += 1

    materials_first_row = row
    for r in consumables:
        natural_qty = r.tz_item.qty
        natural_unit = r.tz_item.unit or ""
        purchase_qty = None
        if natural_qty is not None and r.matched_item is not None:
            purchase_qty = estimate_purchase_quantity(natural_qty, natural_unit, r.matched_item)

        ws.cell(row=row, column=1, value=r.tz_item.raw_name)
        ws.cell(row=row, column=2, value=round(natural_qty, 1) if natural_qty is not None else None)
        ws.cell(row=row, column=3, value=natural_unit)
        ws.cell(row=row, column=4, value=r.matched_item.name if r.matched_item else "НЕ НАЙДЕНО")
        ws.cell(row=row, column=5, value=r.matched_item.supplier if r.matched_item else "—")
        price_cell = ws.cell(
            row=row, column=6,
            value=r.matched_item.final_price if r.matched_item else 0,
        )
        price_cell.number_format = CURRENCY_FMT

        pack_qty_cell = ws.cell(row=row, column=7, value=purchase_qty if purchase_qty is not None else 0)
        if purchase_qty is None:
            pack_qty_cell.fill = INPUT_FILL  # не удалось пересчитать автоматически — эксперт вводит сам

        sum_cell = ws.cell(row=row, column=8, value=f"=G{row}*F{row}")
        sum_cell.number_format = CURRENCY_FMT
        ws.cell(row=row, column=9, value=round(r.confidence, 2))

        review_note = []
        if natural_qty is None:
            review_note.append("не удалось оценить потребность")
        elif purchase_qty is None:
            review_note.append("не удалось перевести в упаковки — ввести вручную")
        if r.requires_expert_review:
            review_note.append("проверить совпадение")
        if not r.price_comparable:
            review_note.append("цена не нормализована (разные ед.изм.)")
        if r.alternatives:
            review_note.append(f"альтернатив: {len(r.alternatives)}")
        ws.cell(row=row, column=10, value="; ".join(review_note) if review_note else "ок")

        _apply_body_style(ws, row, len(headers))
        if r.requires_expert_review or not r.matched_item or purchase_qty is None:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row, column=col).fill = REVIEW_FILL
        if purchase_qty is None:
            pack_qty_cell.fill = INPUT_FILL

        row += 1

    materials_last_row = row - 1
    row += 1
    ws.cell(row=row, column=1, value="Итого материалы").font = TOTAL_FONT
    r_materials_total = row
    mat_total_cell = ws.cell(
        row=row, column=8,
        value=f"=SUM(H{materials_first_row}:H{materials_last_row})" if materials_last_row >= materials_first_row else 0,
    )
    mat_total_cell.font = TOTAL_FONT
    mat_total_cell.number_format = CURRENCY_FMT
    mat_total_cell.border = BORDER
    row += 2

    # ============================== 3. Накладные/маржа ==============================
    ws.cell(row=row, column=1, value="3. Накладные расходы и маржа").font = SECTION_FONT
    row += 1
    ws.cell(row=row, column=1, value="Наценка, % (экспертная оценка — агент не считает сам)")
    margin_cell = ws.cell(row=row, column=2, value=0.20)  # плейсхолдер 20% — эксперт правит
    margin_cell.number_format = "0.0%"
    margin_cell.fill = INPUT_FILL
    margin_cell.border = BORDER
    r_margin = row
    row += 2

    # ============================== Итог ==============================
    ws.cell(row=row, column=1, value="ИТОГО СЕБЕСТОИМОСТЬ (ФОТ + материалы, за весь контракт)").font = TOTAL_FONT
    r_cost = row
    cost_cell = ws.cell(row=row, column=2, value=f"=B{r_labor_total}+H{r_materials_total}")
    cost_cell.number_format = CURRENCY_FMT
    cost_cell.font = TOTAL_FONT
    row += 2

    # ============================== НДС ==============================
    ws.cell(row=row, column=1, value="Ставка НДС, % (эксперт — проверить применимость к вашей системе налогообложения)")
    vat_rate_cell = ws.cell(row=row, column=2, value=0.07)  # плейсхолдер 7% — эксперт правит
    vat_rate_cell.number_format = "0.0%"
    vat_rate_cell.fill = INPUT_FILL
    vat_rate_cell.border = BORDER
    r_vat_rate = row
    row += 2

    # ============================== Себестоимость в месяц / в год ==============================
    ws.cell(row=row, column=1, value="Себестоимость в месяц (ФОТ в месяц + материалы в среднем за месяц)").font = TOTAL_FONT
    r_cost_month = row
    cost_month_cell = ws.cell(row=row, column=2, value=f"=B{r_fot_month}+H{r_materials_total}/B{r_months}")
    cost_month_cell.number_format = CURRENCY_FMT
    cost_month_cell.font = TOTAL_FONT
    cost_month_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  НДС в месяц")
    vat_month_cell = ws.cell(row=row, column=2, value=f"=B{r_cost_month}*B{r_vat_rate}")
    vat_month_cell.number_format = CURRENCY_FMT
    vat_month_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  Себестоимость в месяц с НДС").font = Font(name=FONT_NAME, bold=True, size=10)
    cost_month_vat_cell = ws.cell(row=row, column=2, value=f"=B{r_cost_month}*(1+B{r_vat_rate})")
    cost_month_vat_cell.number_format = CURRENCY_FMT
    cost_month_vat_cell.font = Font(name=FONT_NAME, bold=True, size=10)
    cost_month_vat_cell.border = BORDER
    row += 2

    ws.cell(row=row, column=1, value="Себестоимость в год").font = TOTAL_FONT
    r_cost_year = row
    cost_year_cell = ws.cell(row=row, column=2, value=f"=B{r_cost_month}*12")
    cost_year_cell.number_format = CURRENCY_FMT
    cost_year_cell.font = TOTAL_FONT
    cost_year_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  НДС в год")
    vat_year_cell = ws.cell(row=row, column=2, value=f"=B{r_cost_year}*B{r_vat_rate}")
    vat_year_cell.number_format = CURRENCY_FMT
    vat_year_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  Себестоимость в год с НДС").font = Font(name=FONT_NAME, bold=True, size=10)
    cost_year_vat_cell = ws.cell(row=row, column=2, value=f"=B{r_cost_year}*(1+B{r_vat_rate})")
    cost_year_vat_cell.number_format = CURRENCY_FMT
    cost_year_vat_cell.font = Font(name=FONT_NAME, bold=True, size=10)
    cost_year_vat_cell.border = BORDER
    row += 2

    ws.cell(row=row, column=1, value="  НДС за весь срок контракта")
    vat_contract_cell = ws.cell(row=row, column=2, value=f"=B{r_cost}*B{r_vat_rate}")
    vat_contract_cell.number_format = CURRENCY_FMT
    vat_contract_cell.border = BORDER
    row += 1

    ws.cell(row=row, column=1, value="  Себестоимость за весь срок контракта с НДС").font = Font(
        name=FONT_NAME, bold=True, size=10
    )
    cost_contract_vat_cell = ws.cell(row=row, column=2, value=f"=B{r_cost}*(1+B{r_vat_rate})")
    cost_contract_vat_cell.number_format = CURRENCY_FMT
    cost_contract_vat_cell.font = Font(name=FONT_NAME, bold=True, size=10)
    cost_contract_vat_cell.border = BORDER
    row += 2

    ws.cell(row=row, column=1, value="ИТОГО СМЕТА (с наценкой, без учёта НДС выше — наценка считается от себестоимости без НДС)").font = Font(
        name=FONT_NAME, bold=True, size=13, color="C00000"
    )
    final_cell = ws.cell(row=row, column=2, value=f"=B{r_cost}*(1+B{r_margin})")
    final_cell.number_format = CURRENCY_FMT
    final_cell.font = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
    row += 2

    ws.cell(row=row, column=1,
            value="Источники сумм: B — расчёт кода (детерминированная арифметика), "
                  "желтые ячейки — ввод/проверка эксперта. Материалы посчитаны на весь срок "
                  "контракта, \"в месяц\"/\"в год\" — среднее (реальный расход по месяцам может "
                  "быть неравномерным). НДС посчитан от себестоимости (без наценки) отдельно для "
                  "месяца/года/контракта — проверьте, относится ли ставка 7% к вашей системе "
                  "налогообложения. Смета не отправляется заказчику без подтверждения эксперта "
                  "(базовая ставка, оформление, наценка, ставка НДС, позиции без прайса).")
    ws.cell(row=row, column=1).font = NOTE_FONT

    # ============================== Лист "Альтернативы" ==============================
    needs_alt_sheet = [r for r in consumables if r.alternatives]
    if needs_alt_sheet:
        ws2 = wb.create_sheet("Альтернативы")
        ws2.sheet_view.showGridLines = False
        for col, width in zip("ABCD", [40, 45, 22, 14]):
            ws2.column_dimensions[col].width = width
        headers2 = ["Позиция ТЗ", "Кандидат", "Поставщик", "Цена, ₽"]
        for col, h in enumerate(headers2, start=1):
            ws2.cell(row=1, column=col, value=h)
        _style_header_row(ws2, 1, len(headers2))
        arow = 2
        for r in needs_alt_sheet:
            ws2.cell(row=arow, column=1, value=r.tz_item.raw_name).font = Font(name=FONT_NAME, bold=True, size=10)
            arow += 1
            chosen_label = f"✓ ВЫБРАНО: {r.matched_item.name}" if r.matched_item else "✓ ВЫБРАНО: —"
            ws2.cell(row=arow, column=2, value=chosen_label)
            ws2.cell(row=arow, column=3, value=r.matched_item.supplier if r.matched_item else "—")
            pc = ws2.cell(row=arow, column=4, value=r.matched_item.final_price if r.matched_item else 0)
            pc.number_format = CURRENCY_FMT
            _apply_body_style(ws2, arow, len(headers2))
            arow += 1
            for alt in r.alternatives:
                ws2.cell(row=arow, column=2, value=alt.name)
                ws2.cell(row=arow, column=3, value=alt.supplier)
                pc = ws2.cell(row=arow, column=4, value=alt.final_price)
                pc.number_format = CURRENCY_FMT
                _apply_body_style(ws2, arow, len(headers2))
                arow += 1
            arow += 1

    wb.save(output_path)
