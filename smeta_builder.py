"""
smeta_builder.py — общий переиспользуемый слой для генерации Excel-смет.

Зачем: за время работы накопилось 10+ скриптов *_smeta_export.py
(smeta_export.py, mlrz_smeta_export.py, krown_independent_smeta.py,
territory_smeta_export.py, discount_smeta_export.py, medical_*_export.py,
zvezda_smeta_export.py, ...). Каждый раз повторялись один-в-один: стилевые
константы, стиль заголовка/строки, и, главное, ~60-строчный блок таблицы
расходных материалов (резолвер → таблица → ИТОГО SUM). Это дублирование
уже приводило к багам (см. смету Краун: борта таблицы после вставки новой
секции ссылались на старые номера строк) — типичный риск copy-paste кода.

Этот модуль НЕ пытается свести всё многообразие смет к одному шаблону —
у каждого объекта своя специфика (аренда техники, твёрдый пол/скидка,
лето/зима, поединичные расценки...). Вместо этого он даёт:

  1. Единые стилевые константы (замена HEADER_FILL/INPUT_FILL/... —
     раньше copy-paste в каждом файле, теперь один источник правды).
  2. Класс SmetaSheet — обёртка над листом с курсором строки, чтобы не
     передавать `row` руками и не путать номера строк вручную.
  3. build_materials_table() — стандартизованный блок "3. Расходные
     материалы", обязательный по стандарту сметы (сопоставление через
     resolver ВСЕГДА, наименование поставщика ВСЕГДА, количество —
     ФОРМУЛОЙ норматив×база, а не захардкоженным числом — раньше это
     было computed-значением в Python, теперь настоящая формула Excel
     для большинства типов норм, см. _FORMULA_BUILDERS ниже).
  4. build_contract_term_section() / build_price_summary_section() —
     стандартные блоки "срок контракта" и "цена/мес → цена/год".

Специфичные для объекта секции (аренда техники, твёрдый пол, лето/зима)
по-прежнему пишутся в скрипте объекта, но используют SmetaSheet для
базовых операций — так номера строк не приходится отслеживать вручную.
"""
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.worksheet import Worksheet

from knowledge_base import ObjectParams, load_consumable_norms, load_object_defaults
from models import TZItem
from resolver import resolve_all
from unit_price import estimate_purchase_quantity

# ============================== Стилевые константы (единый источник) ==============================
FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=12)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14)
TOTAL_FONT = Font(name=FONT_NAME, bold=True, size=12)
FINAL_FONT = Font(name=FONT_NAME, bold=True, size=13, color="C00000")
BODY_FONT = Font(name=FONT_NAME, size=10)
SMALL_FONT = Font(name=FONT_NAME, size=9)
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="666666")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")   # эксперт вводит/проверяет
REVIEW_FILL = PatternFill("solid", fgColor="FFC7CE")  # низкая уверенность сопоставления
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY_FMT = '#,##0.00 ₽;(#,##0.00 ₽);-'
PERCENT_FMT = "0.0%"

STANDARD_LEGEND = "Жёлтые ячейки — ввод/проверка экспертом. Розовые — низкая уверенность сопоставления, нужна проверка."


def _style_header_row(ws: Worksheet, row: int, n_cols: int) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _style_body_row(ws: Worksheet, row: int, n_cols: int, font=BODY_FONT) -> None:
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = font
        cell.border = BORDER


# ============================== SmetaSheet: курсор строки + базовые операции ==============================

@dataclass
class SmetaSheet:
    """
    Обёртка над Worksheet с курсором строки — избавляет от ручного
    отслеживания номеров строк (источник как минимум 3 багов за проект:
    захардкоженные $B$14, съехавшие границы после вставки секции,
    ссылка на неправильную колонку после смены раскладки).
    """
    ws: Worksheet
    row: int = 1

    @classmethod
    def new(cls, wb: Workbook, title: str, col_widths: dict[str, int]) -> "SmetaSheet":
        ws = wb.active if wb.active.title == "Sheet" and not wb.worksheets[1:] else wb.create_sheet(title)
        ws.title = title
        ws.sheet_view.showGridLines = False
        for col, width in col_widths.items():
            ws.column_dimensions[col].width = width
        return cls(ws=ws, row=1)

    def title_block(self, object_name: str, subtitle: Optional[str] = None, region: Optional[str] = None,
                     prepared_date: Optional[date] = None, legend: bool = True) -> None:
        self.ws.cell(row=self.row, column=1, value=f"Смета — {object_name}").font = TITLE_FONT
        self.row += 1
        if subtitle:
            self.ws.cell(row=self.row, column=1, value=subtitle).font = NOTE_FONT
            self.row += 1
        if region:
            self.ws.cell(row=self.row, column=1, value=f"Регион: {region}").font = SMALL_FONT
            self.row += 1
        self.ws.cell(row=self.row, column=1,
                      value=f"Дата формирования: {(prepared_date or date.today()).isoformat()}").font = SMALL_FONT
        self.row += 2
        if legend:
            self.ws.cell(row=self.row, column=1, value=STANDARD_LEGEND).font = SMALL_FONT
            self.row += 2

    def section(self, title: str) -> int:
        self.ws.cell(row=self.row, column=1, value=title).font = SECTION_FONT
        self.row += 1
        return self.row

    def value_row(self, label: str, value, *, editable: bool = False, fmt: Optional[str] = None,
                  col: int = 2, bold: bool = False, border_cols: tuple[int, ...] = (1, 2)) -> int:
        """Пишет строку 'label -> value' в текущей строке, двигает курсор, возвращает номер строки."""
        r = self.row
        self.ws.cell(row=r, column=1, value=label).font = TOTAL_FONT if bold else BODY_FONT
        cell = self.ws.cell(row=r, column=col, value=value)
        cell.font = TOTAL_FONT if bold else BODY_FONT
        if fmt:
            cell.number_format = fmt
        if editable:
            cell.fill = INPUT_FILL
        for c in border_cols:
            self.ws.cell(row=r, column=c).border = BORDER
        self.row += 1
        return r

    def note(self, text: str, span_cols: int = 6) -> None:
        r = self.row
        self.ws.cell(row=r, column=1, value=text).font = NOTE_FONT
        if span_cols > 1:
            self.ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span_cols)
        self.row += 1

    def blank(self, n: int = 1) -> None:
        self.row += n

    def table_header(self, headers: list[str]) -> int:
        r = self.row
        for c, h in enumerate(headers, start=1):
            self.ws.cell(row=r, column=c, value=h)
        _style_header_row(self.ws, r, len(headers))
        self.row += 1
        return r


# ============================== Стандартный блок "1. Срок контракта" ==============================

def build_contract_term_section(sheet: SmetaSheet, contract_months: int, area_total_sqm: Optional[float] = None,
                                 extra_rows: Optional[list[tuple[str, object]]] = None) -> dict[str, int]:
    """
    Стандарт: срок контракта — ПЕРВАЯ секция сметы. Возвращает {"months": row, ...}
    со ссылками на строки, чтобы остальные секции могли ссылаться формулами.
    """
    sheet.section("1. Срок контракта")
    refs = {}
    refs["months"] = sheet.value_row("Срок оказания услуг, мес", contract_months, editable=True)
    if area_total_sqm is not None:
        refs["area_total"] = sheet.value_row("Площадь всего, м²", area_total_sqm, editable=True)
    for label, value in (extra_rows or []):
        refs[label] = sheet.value_row(label, value, editable=True)
    sheet.blank()
    return refs


# ============================== Стандартный блок "2. Цена в месяц/год" ==============================

def build_price_summary_section(sheet: SmetaSheet, *, cost_month_formula: str, months_row: int,
                                 vat_rate: float, title: str = "2. Цена") -> dict[str, int]:
    """
    Стандарт: сразу после срока — цена/мес и цена/год (ДО численности и
    материалов, см. smeta-template-standard в памяти проекта).
    cost_month_formula — Excel-формула строки "себестоимость/цена в месяц без НДС"
    (например "=B12+F30/B5"), которую вызывающий код уже построил из своих
    секций численности/материалов; эта функция достраивает НДС и год.
    """
    sheet.section(title)
    refs = {}
    refs["vat_rate"] = sheet.value_row("Ставка НДС, %", vat_rate, editable=True, fmt=PERCENT_FMT)
    r_month = sheet.row
    sheet.ws.cell(row=r_month, column=1, value="ЦЕНА В МЕСЯЦ (без НДС)").font = FINAL_FONT
    sheet.ws.cell(row=r_month, column=2, value=cost_month_formula).font = FINAL_FONT
    sheet.ws.cell(row=r_month, column=2).number_format = CURRENCY_FMT
    sheet.row += 1
    refs["month_novat"] = r_month
    r_month_vat = sheet.row
    sheet.ws.cell(row=r_month_vat, column=1, value="ЦЕНА В МЕСЯЦ (с НДС)").font = FINAL_FONT
    sheet.ws.cell(row=r_month_vat, column=2,
                  value=f"=B{r_month}*(1+B{refs['vat_rate']})").font = FINAL_FONT
    sheet.ws.cell(row=r_month_vat, column=2).number_format = CURRENCY_FMT
    sheet.row += 1
    refs["month_vat"] = r_month_vat
    r_year = sheet.row
    sheet.ws.cell(row=r_year, column=1, value="ЦЕНА В ГОД (с НДС)").font = FINAL_FONT
    sheet.ws.cell(row=r_year, column=2,
                  value=f"=B{r_month_vat}*MIN(B{months_row},12)").font = FINAL_FONT
    sheet.ws.cell(row=r_year, column=2).number_format = CURRENCY_FMT
    sheet.row += 2
    refs["year_vat"] = r_year
    return refs


# ============================== Формулы для количества материалов (норматив × база) ==============================
# Каждый формула-билдер получает ссылки на editable-ячейки общих параметров
# (params_refs) и editable-ячейку коэффициента конкретной нормы (coeff_ref),
# и возвращает строку Excel-формулы. Логика 1:1 повторяет knowledge_base.py
# estimate_consumable_quantity() — если там появится новый тип формулы, его
# нужно завести и здесь (иначе для новых норм будет статическое число).

def _f_per_area_per_cleaning_ml(p, coeff_ref):
    return f"=ROUND({coeff_ref}*{p['area_sqm']}*{p['cleaning_days']}/1000,1)"


def _f_per_sanitary_area_per_cleaning_ml(p, coeff_ref, sanitary_fraction):
    return f"=ROUND({coeff_ref}*{p['area_sqm']}*{sanitary_fraction}*{p['cleaning_days']}/1000,1)"


def _f_per_area_per_general_ml(p, coeff_ref):
    return f"=ROUND({coeff_ref}*{p['area_sqm']}*{p['general_days']}/1000,1)"


def _f_per_cleaning_day_fixed_pieces(p, coeff_ref):
    return f"={coeff_ref}*{p['cleaning_days']}"


def _f_fixed_monthly_liters(p, coeff_ref):
    return f"=ROUND({coeff_ref}*{p['contract_months']},1)"


def _f_per_staff_per_week_pieces(p, coeff_ref):
    return f"=ROUND({coeff_ref}*{p['staff_count']}*{p['contract_weeks']},0)"


def _f_per_staff_fixed_plus_spare(p, coeff_ref):
    return f"={p['staff_count']}+{coeff_ref}"


def _f_per_cleaning_day_per_bin(p, coeff_ref, bin_density):
    return f"=ROUND({p['area_sqm']}/{bin_density},0)*{coeff_ref}*{p['cleaning_days']}"


@dataclass
class MaterialsTableResult:
    first_row: int
    last_row: int
    total_row: int
    total_cell: str          # напр. "F30" — ссылаться из других секций
    unresolved: list[str] = field(default_factory=list)  # позиции без совпадения в прайсе — на проверку эксперту


def build_materials_table(
    sheet: SmetaSheet,
    params: ObjectParams,
    price_lists: list,
    *,
    supplier_adapters: Optional[dict] = None,
    exclude_norms: set[str] = frozenset(),
    no_correction_norms: set[str] = frozenset(),
    correction_factor: float = 1.0,
    correction_note: Optional[str] = None,
    title: str = "3. Расходные материалы",
) -> MaterialsTableResult:
    """
    Стандартный блок расходных материалов. Реализует все пункты стандарта
    сметы из памяти проекта:
      - сопоставление с прайс-листом ВСЕГДА (resolver.resolve_all), не заглушка;
      - количество — ФОРМУЛА норматив×база (для норм с простыми линейными
        формулами — фактическая формула Excel, привязанная к editable-ячейкам
        площади/дней/численности/коэффициента; для формул с промежуточным
        округлением через "корзины"/доли — тоже формула, см. _f_* выше);
      - наименование поставщика — всегда отдельная колонка;
      - строки без совпадения в прайсе — INPUT_FILL, чтобы эксперт ввёл вручную,
        не тихий 0.
    correction_factor — единый поправочный коэффициент (например 1/7 для
    крупных объектов с механизированной уборкой, см. mlrz/krown) — тоже
    editable-ячейка, а не захардкожен в формуле.
    """
    norms = load_consumable_norms()
    defaults = load_object_defaults()
    sheet.section(title)

    # --- общие editable-параметры расчёта (одна ячейка на весь блок, а не
    #     константа в коде — эксперт может подвинуть площадь/дни и увидеть
    #     пересчёт всех строк материалов сразу) ---
    p = {}
    p["area_sqm"] = f"$B${sheet.value_row('База: площадь, м²', params.area_sqm, editable=True)}"
    p["cleaning_days"] = f"$B${sheet.value_row('База: дней уборки в месяц', params.cleaning_days, editable=True)}"
    p["general_days"] = f"$B${sheet.value_row('База: генеральных уборок в месяц', params.general_days, editable=True)}"
    p["contract_months"] = f"$B${sheet.value_row('База: срок контракта, мес', params.contract_months, editable=True)}"
    p["staff_count"] = f"$B${sheet.value_row('База: численность, чел.', params.staff_count, editable=True)}"
    p["contract_weeks"] = f"$B${sheet.value_row('База: недель в контракте', round(params.contract_weeks, 1), editable=True)}"
    r_correction = sheet.value_row("Поправочный коэффициент (крупный объект/мех.уборка и т.п.)",
                                    correction_factor, editable=True)
    correction_ref = f"$B${r_correction}"
    if correction_note:
        sheet.note(correction_note)
    sheet.blank()

    headers = ["Материал", "Норматив (коэфф.)", "Потребность, ед.", "Ед.", "Найдено в прайсе",
               "Поставщик", "Упаковок", "Цена/уп, ₽", "Сумма/мес, ₽"]
    sheet.table_header(headers)
    first_row = sheet.row

    # Сначала считаем реальные (Python) количества — нужны для подбора
    # позиции в прайсе (resolve_all сопоставляет по названию+кол-ву) и как
    # запасное статическое значение для формул, которые не формула-изируем.
    from knowledge_base import estimate_consumable_quantity
    included = [nid for nid in norms if nid not in exclude_norms]
    tz_items, py_qtys, formula_strs = [], [], []
    for norm_id in included:
        norm = norms[norm_id]
        qty, unit, _ = estimate_consumable_quantity(norm_id, params)
        if norm_id not in no_correction_norms:
            qty = round(qty * correction_factor, 2)
        tz_items.append(TZItem(raw_name=norm["display_name"], qty=qty, unit=unit))
        py_qtys.append((qty, unit))

    match_results = resolve_all(tz_items, price_lists, supplier_adapters=supplier_adapters or {})

    unresolved = []
    for norm_id, r, (qty, unit) in zip(included, match_results, py_qtys):
        norm = norms[norm_id]
        formula = norm["formula"]
        coeff = norm["coefficient"]
        row = sheet.row
        sheet.ws.cell(row=row, column=1, value=norm["display_name"])
        r_coeff = row
        coeff_cell = sheet.ws.cell(row=row, column=2, value=coeff)
        coeff_cell.fill = INPUT_FILL
        coeff_ref = f"B{r_coeff}"

        if formula == "per_area_per_cleaning_ml":
            qty_formula = _f_per_area_per_cleaning_ml(p, coeff_ref)
        elif formula == "per_sanitary_area_per_cleaning_ml":
            qty_formula = _f_per_sanitary_area_per_cleaning_ml(p, coeff_ref, defaults["sanitary_area_fraction"])
        elif formula == "per_area_per_general_ml":
            qty_formula = _f_per_area_per_general_ml(p, coeff_ref)
        elif formula == "per_cleaning_day_fixed_pieces":
            qty_formula = _f_per_cleaning_day_fixed_pieces(p, coeff_ref)
        elif formula == "fixed_monthly_liters":
            qty_formula = _f_fixed_monthly_liters(p, coeff_ref)
        elif formula == "per_staff_per_week_pieces":
            qty_formula = _f_per_staff_per_week_pieces(p, coeff_ref)
        elif formula == "per_staff_fixed_plus_spare":
            qty_formula = _f_per_staff_fixed_plus_spare(p, coeff_ref)
        elif formula == "per_cleaning_day_per_bin":
            qty_formula = _f_per_cleaning_day_per_bin(p, coeff_ref, defaults["bin_density_sqm"])
        else:
            qty_formula = None  # неизвестный тип формулы в consumable_norms.json — не должно происходить
            # с текущим набором из 8 типов; если появится новый formula-тип
            # без билдера выше, попадаем сюда и переходим на статику ниже.

        qty_cell = sheet.ws.cell(row=row, column=3)
        if qty_formula is not None:
            # qty_formula начинается с "=" — вставляем поправочный коэффициент
            # внутрь формулы (кроме норм из no_correction_norms, где он не нужен).
            body = qty_formula[1:]
            if norm_id not in no_correction_norms:
                body = f"({body})*{correction_ref}"
            qty_cell.value = f"={body}"
        else:
            qty_cell.value = qty
            qty_cell.fill = INPUT_FILL  # не удалось формула-изировать — эксперт может править вручную как число

        sheet.ws.cell(row=row, column=4, value=unit)
        matched_name = r.matched_item.name if r.matched_item else "НЕ НАЙДЕНО"
        sheet.ws.cell(row=row, column=5, value=matched_name)
        sheet.ws.cell(row=row, column=6, value=r.matched_item.supplier if r.matched_item else "—")

        packs = estimate_purchase_quantity(qty, unit, r.matched_item) if r.matched_item and unit in ("л", "кг", "шт") else None
        packs_cell = sheet.ws.cell(row=row, column=7, value=packs if packs is not None else qty)
        if packs is None:
            packs_cell.fill = INPUT_FILL

        price_cell = sheet.ws.cell(row=row, column=8, value=r.matched_item.final_price if r.matched_item else 0)
        price_cell.number_format = CURRENCY_FMT
        if not r.matched_item:
            price_cell.fill = INPUT_FILL
            unresolved.append(norm["display_name"])

        sum_cell = sheet.ws.cell(row=row, column=9, value=f"=G{row}*H{row}")
        sum_cell.number_format = CURRENCY_FMT

        _style_body_row(sheet.ws, row, len(headers), font=SMALL_FONT)
        if not r.matched_item or r.requires_expert_review:
            for c in range(1, len(headers) + 1):
                sheet.ws.cell(row=row, column=c).fill = REVIEW_FILL
            if not r.matched_item:
                price_cell.fill = INPUT_FILL  # приоритет над REVIEW_FILL — это то, что реально нужно ввести

        sheet.row += 1

    last_row = sheet.row - 1
    total_row = sheet.row
    sheet.ws.cell(row=total_row, column=1, value="ИТОГО материалы/мес").font = TOTAL_FONT
    total_cell = sheet.ws.cell(row=total_row, column=9, value=f"=SUM(I{first_row}:I{last_row})")
    total_cell.font = TOTAL_FONT
    total_cell.number_format = CURRENCY_FMT
    sheet.row += 2

    return MaterialsTableResult(
        first_row=first_row, last_row=last_row, total_row=total_row,
        total_cell=f"I{total_row}", unresolved=unresolved,
    )
