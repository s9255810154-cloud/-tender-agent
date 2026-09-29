"""
База знаний агента: нормативы расхода расходных материалов, параметры
объекта по умолчанию, заготовки под ставки труда и оценку численности.

Хранится как JSON в knowledge_base/, а не как константы в коде — чтобы:
  - эксперт мог поправить коэффициент, не трогая Python;
  - одни и те же нормы переиспользовались для ЛЮБОГО следующего ТЗ, а не
    копировались в каждый demo-скрипт заново;
  - было явно видно, что ещё НЕ заполнено (staff_norms/salary_rules — это
    открытые TODO из tender-agent-architecture.md), а не терялось в коде
    в виде "0" без объяснений.

Формула расчёта количества расходника берётся из поля "formula" записи
в consumable_norms.json:
  per_cleaning_day_fixed_pieces      coefficient (шт/день) × cleaning_days
  per_cleaning_day_per_bin           coefficient (замен/день/корзину) × кол-во корзин × cleaning_days
  per_area_per_cleaning_ml           coefficient (мл/м²/уборку) × area × cleaning_days / 1000 → л
  per_sanitary_area_per_cleaning_ml  то же, но area = area × sanitary_area_fraction
  per_area_per_general_ml            coefficient (мл/м²/генеральную) × area × general_days / 1000 → л
  fixed_monthly_liters               coefficient (л/мес) × contract_months
  per_staff_per_week_pieces          coefficient (шт/сотрудника/неделю) × staff_count × contract_weeks
  per_staff_fixed_plus_spare         staff_count + coefficient (запас)
"""
import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

_KB_DIR = Path(__file__).parent / "knowledge_base"


@dataclass
class ObjectParams:
    """Параметры конкретного объекта — как правило, выводятся из самого ТЗ
    (см. tender-agent-architecture.md: площадь — из объёмов услуг по месяцам,
    не из первого попавшегося числа)."""
    area_sqm: float
    cleaning_days: int
    general_days: int
    staff_count: int
    contract_months: int

    @property
    def contract_weeks(self) -> float:
        return self.contract_months * 52 / 12


def _load_json(name: str) -> dict:
    with open(_KB_DIR / name, encoding="utf-8") as f:
        return json.load(f)


def load_consumable_norms() -> dict[str, dict]:
    norms = _load_json("consumable_norms.json")
    return {n["id"]: n for n in norms}


def load_object_defaults() -> dict:
    return _load_json("object_defaults.json")


def load_staff_norms() -> dict:
    return _load_json("staff_norms.json")


def load_salary_rules() -> dict:
    return _load_json("salary_rules.json")


def estimate_consumable_quantity(
    norm_id: str,
    params: ObjectParams,
    defaults: Optional[dict] = None,
) -> tuple[float, str, str]:
    """
    Возвращает (количество, единица, текстовое описание формулы для аудита).
    KeyError, если norm_id не найден в базе — явная ошибка лучше тихого нуля.
    """
    norms = load_consumable_norms()
    if norm_id not in norms:
        raise KeyError(
            f"Норма '{norm_id}' не найдена в consumable_norms.json. "
            f"Известные: {', '.join(sorted(norms))}"
        )
    norm = norms[norm_id]
    defaults = defaults or load_object_defaults()
    formula = norm["formula"]
    coeff = norm["coefficient"]
    unit = norm["unit"]

    if formula == "per_cleaning_day_fixed_pieces":
        qty = coeff * params.cleaning_days
        desc = f"{coeff} шт/день × {params.cleaning_days} дней уборки"

    elif formula == "per_cleaning_day_per_bin":
        bins = round(params.area_sqm / defaults["bin_density_sqm"])
        qty = coeff * bins * params.cleaning_days
        desc = (f"{bins} корзин (площадь {params.area_sqm} / {defaults['bin_density_sqm']} м²/корзину) "
                f"× {coeff} замена/день × {params.cleaning_days} дней")

    elif formula == "per_area_per_cleaning_ml":
        qty = round(coeff * params.area_sqm * params.cleaning_days / 1000, 1)
        desc = f"{coeff} мл/м² × {params.area_sqm} м² × {params.cleaning_days} дней / 1000"

    elif formula == "per_sanitary_area_per_cleaning_ml":
        sanitary_area = params.area_sqm * defaults["sanitary_area_fraction"]
        qty = round(coeff * sanitary_area * params.cleaning_days / 1000, 1)
        desc = (f"{coeff} мл/м² × {round(sanitary_area, 1)} м² санитарных помещений "
                f"({defaults['sanitary_area_fraction'] * 100:.0f}% от площади) "
                f"× {params.cleaning_days} дней / 1000")

    elif formula == "per_area_per_general_ml":
        qty = round(coeff * params.area_sqm * params.general_days / 1000, 1)
        desc = f"{coeff} мл/м² × {params.area_sqm} м² × {params.general_days} генеральных уборок / 1000"

    elif formula == "fixed_monthly_liters":
        qty = round(coeff * params.contract_months, 1)
        desc = f"{coeff} л/мес × {params.contract_months} мес"

    elif formula == "per_staff_per_week_pieces":
        qty = round(coeff * params.staff_count * params.contract_weeks)
        desc = f"{coeff} шт/сотрудника/неделю × {params.staff_count} чел. × {params.contract_weeks:.0f} нед."

    elif formula == "per_staff_fixed_plus_spare":
        qty = params.staff_count + coeff
        desc = f"{params.staff_count} чел. + {coeff} запас"

    else:
        raise ValueError(f"Неизвестная формула '{formula}' в норме '{norm_id}'")

    return qty, unit, desc


def get_salary(region: str, object_complexity: str, schedule_complexity: str) -> Optional[dict]:
    """
    Возвращает правило целиком (monthly_salary_range, monthly_salary_cap,
    official_employment_overhead_pct и т.д.) или None, если правило не
    найдено — таблица ставок пока не заполнена для этой комбинации
    (открытый TODO, см. salary_rules.json)."""
    rules = load_salary_rules()
    for rule in rules.get("rules", []):
        if (rule["region"] == region
                and rule["object_complexity"] == object_complexity
                and rule["schedule_complexity"] == schedule_complexity):
            return rule
    return None


def get_staff_estimation_norms() -> Optional[dict]:
    """None, если sqm_per_person/hours_per_person_per_shift ещё не заполнены
    (открытый TODO, см. staff_norms.json) — тогда оценивать численность по
    площади/часам нельзя, нужно брать явно указанную в ТЗ, если есть."""
    norms = load_staff_norms()
    if norms.get("sqm_per_person") is None or norms.get("hours_per_person_per_shift") is None:
        return None
    return norms


def get_combined_shift_salary(
    region: str,
    object_complexity: str,
    schedule_complexity: str,
    override_base_salary: Optional[float] = None,
) -> Optional[dict]:
    """
    Готовит базовую ставку и надбавку за оформление из salary_rules.json:
    середина диапазона (или override_base_salary), ограниченная потолком.
    Часы по видам уборки сюда НЕ входят — это факт конкретного объекта
    (из графика в его ТЗ), а не свойство рыночной ставки; передавайте их
    в LaborCostInput отдельно, из данных самого ТЗ.
    Возвращает None, если правило не найдено.
    """
    rule = get_salary(region, object_complexity, schedule_complexity)
    if rule is None:
        return None
    rng = rule["monthly_salary_range"]
    base = override_base_salary if override_base_salary is not None else (rng["min"] + rng["max"]) / 2
    base = min(base, rule["monthly_salary_cap"])
    return {
        "base_shift_salary": base,
        "official_overhead_pct": rule["official_employment_overhead_pct"],
        "official_employment_default": rule.get("official_employment_default", False),
    }


def load_real_salary_benchmarks() -> list[dict]:
    return _load_json("real_salary_benchmarks.json")["benchmarks"]


def find_nearest_salary_benchmarks(
    category: str, workdays_per_week: int, daily_hours: float, top_n: int = 3
) -> list[dict]:
    """
    Ищет ближайшие по графику реальные зарплатные ориентиры из Табель.xlsx
    (не единый рыночный диапазон эксперта, а фактическая выплаченная
    зарплата у конкретных сотрудников с похожим графиком). Не интерполирует
    и не выдумывает промежуточное значение — просто возвращает ближайшие
    реальные точки, чтобы эксперт мог сам решить, что из этого релевантно.
    """
    benchmarks = load_real_salary_benchmarks()
    same_category = [b for b in benchmarks if b["category"] == category]
    if not same_category:
        return []

    def distance(b: dict) -> float:
        # часы — основной фактор; несовпадение графика (5/2 vs 7/0) штрафуется отдельно
        hour_diff = abs(b["daily_hours"] - daily_hours)
        workday_penalty = 0 if b["workdays_per_week"] == workdays_per_week else 3
        return hour_diff + workday_penalty

    return sorted(same_category, key=distance)[:top_n]


def load_additional_service_rates() -> dict[str, dict]:
    """Расценки на допуслуги (химчистка, отмывка окон, вывоз мусора, мастика
    и т.п.) — не входят в стандартную уборку, заказываются отдельно."""
    data = _load_json("additional_services_rates.json")
    return {s["id"]: s for s in data["services"]}


def get_service_rate(service_id: str) -> dict:
    """Одна расценка допуслуги по id. KeyError, если не найдена — явная
    ошибка лучше тихого нуля."""
    rates = load_additional_service_rates()
    if service_id not in rates:
        raise KeyError(
            f"Услуга '{service_id}' не найдена в additional_services_rates.json. "
            f"Известные: {', '.join(sorted(rates))}"
        )
    return rates[service_id]


def load_tz_document_templates() -> dict[str, dict]:
    """Известные структуры типов ТЗ (какие приложения что значат, где
    настоящая площадь, где ловушка) — см. tz_document_templates.json."""
    data = _load_json("tz_document_templates.json")
    return {t["type_id"]: t for t in data["templates"]}


def get_tz_template(type_id: str) -> dict:
    templates = load_tz_document_templates()
    if type_id not in templates:
        raise KeyError(
            f"Тип документа '{type_id}' не найден в tz_document_templates.json. "
            f"Известные: {', '.join(sorted(templates))}"
        )
    return templates[type_id]


def load_tender_search_keywords() -> list[str]:
    """Стемы для TenderFilter.keywords_any — резервный канал фильтрации
    закупок на клининг, когда код классификатора недоступен/неточен.
    См. tender_search_keywords.json."""
    return _load_json("tender_search_keywords.json")["keywords_stems"]


def load_premises_labor_time_norms() -> dict:
    return _load_json("premises_labor_time_norms.json")


def estimate_staff_count_by_area(
    area_sqm: float, cleanings_per_shift: int = 1, sqm_per_person: Optional[float] = None
) -> Optional[int]:
    """
    Оценка численности по площади. По умолчанию использует норматив из
    staff_norms.json (800-1200 м²/чел, среднее 1000 — предоставлено
    пользователем как фактическая производительность его компании).
    Явно передайте sqm_per_person, чтобы посчитать по другому значению
    (например, по нижней/верхней границе диапазона для вилки оценки).
    """
    norms = load_staff_norms()
    if area_sqm <= 0:
        return None
    base = sqm_per_person if sqm_per_person is not None else norms["sqm_per_person"]
    if cleanings_per_shift > 1:
        base = base / 1.1  # поправка на уборку чаще раза за смену
    return math.ceil(area_sqm / base)


def load_premises_gesn_rates() -> dict:
    """Официальные расценки (затратный метод) на уборку ПОМЕЩЕНИЙ — 31 вид
    работ с нормами ЗТР (чел-час/ед.) и расхода материалов, плюс официальная
    статистическая ЗП (ЕМИСС/Росстат) для клининга в Москве. См.
    premises_gesn_rates.json."""
    return _load_json("premises_gesn_rates.json")


def load_territory_equipment_rates() -> dict[str, dict]:
    """Расценки на технику/вывоз снега для территориальных смет (аренда
    по вызову) — см. territory_equipment_rates.json."""
    data = _load_json("territory_equipment_rates.json")
    return {r["id"]: r for r in data["rates"]}


def get_territory_equipment_rate(rate_id: str) -> dict:
    rates = load_territory_equipment_rates()
    if rate_id not in rates:
        raise KeyError(
            f"Расценка '{rate_id}' не найдена в territory_equipment_rates.json. "
            f"Известные: {', '.join(sorted(rates))}"
        )
    return rates[rate_id]
