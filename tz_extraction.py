"""
Document Extraction (шаг 4 архитектуры) — структурированное извлечение из
ТЗ через Claude API, вместо ручного чтения.

Разделение ответственности (тот же принцип "LLM не считает деньги", что уже
заложен в архитектуре для сметы — здесь применён и к извлечению): Claude
извлекает ФАКТЫ из текста (таблицу объёмов по периодам как есть, график,
список расходников, явную численность, если есть) — но НЕ вычисляет площадь
объекта сам. Площадь считается ОТДЕЛЬНО детерминированным кодом
(compute_real_area) — тем же приёмом, что был проверен вручную на реальном
ТЗ (Большой Головин, 15): объёмы по периодам делятся на площадь объекта
нацело — площадь восстанавливается как НОД(объёмы), приведённые к целым.

ВАЖНО: в этой среде (чат) нет доступа в сеть — вызов Claude API
(extract_tz_structured) не протестирован живым запросом здесь. Промпт и
JSON-схема составлены под структуру реального документа (проверьте на
вашем сервере с ключом API). Пост-обработка (compute_real_area,
compute_cleaning_days и т.д.) протестирована на РЕАЛЬНЫХ числах из того же
ТЗ — см. demo_tz_extraction.py — и воспроизводит площадь 1002.6 м²/742 дня
уборки, которые мы раньше высчитывали руками.
"""
from dataclasses import dataclass
from datetime import date, datetime
from fractions import Fraction
from functools import reduce
from math import gcd
from typing import Optional

from document_ingestion import IngestedDocument
from models import TZItem

EXTRACTION_TOOL = {
    "name": "record_tz_extraction",
    "description": "Записать структурированные данные, извлечённые из технического задания на уборку помещений.",
    "input_schema": {
        "type": "object",
        "properties": {
            "object_name": {"type": "string", "description": "Краткое название объекта закупки"},
            "region": {"type": "string"},
            "address": {"type": "string"},
            "kpgz_code": {"type": ["string", "null"], "description": "Код КПГЗ/ОКПД2, если указан в документе"},
            "contract_periods": {
                "type": "array",
                "description": (
                    "Строки из раздела 'Перечень объектов закупки' (обычно Приложение 1). "
                    "ВАЖНО: 'Объём (единица измерения)' там — это СУММАРНЫЙ объём услуги ЗА "
                    "УКАЗАННЫЙ ПЕРИОД (например, за месяц), а НЕ площадь объекта. Не путайте эти "
                    "понятия и не пытайтесь сами вычислить площадь объекта — извлеките периоды и "
                    "объёмы КАК ЕСТЬ, площадь считается отдельно детерминированным кодом после вас."
                ),
                "items": {
                    "type": "object",
                    "properties": {
                        "period_start": {"type": "string", "description": "YYYY-MM-DD"},
                        "period_end": {"type": "string", "description": "YYYY-MM-DD"},
                        "service_volume_sqm": {"type": "number"},
                    },
                    "required": ["period_start", "period_end", "service_volume_sqm"],
                },
            },
            "schedule": {
                "type": "object",
                "properties": {
                    "general_cleaning_per_week": {"type": "number", "description": "Сколько раз в неделю генеральная уборка"},
                    "daily_main_hours_per_day": {"type": ["number", "null"], "description": "Суммарные часы ежедневной-основной уборки в будний день"},
                    "daily_supporting_hours_per_day": {"type": ["number", "null"], "description": "Часы ежедневной-поддерживающей уборки в будний день"},
                    "workdays_per_week": {"type": "number", "description": "Обычно 5"},
                },
            },
            "explicit_staff_count": {
                "type": ["object", "null"],
                "properties": {
                    "value": {"type": "integer"},
                    "condition": {"type": "string", "description": "например 'не менее'"},
                },
            },
            "consumables": {
                "type": "array",
                "description": "Из раздела 'Перечень расходных материалов'.",
                "items": {
                    "type": "object",
                    "properties": {
                        "raw_name": {"type": "string"},
                        "characteristics": {"type": "string", "description": "Все указанные характеристики одной строкой"},
                        "unit": {"type": "string"},
                        "explicit_qty": {"type": ["number", "null"], "description": "null, если в ТЗ количество НЕ указано явным числом — не оценивай, оставляй null"},
                    },
                    "required": ["raw_name", "unit"],
                },
            },
        },
        "required": ["object_name", "region", "contract_periods", "consumables"],
    },
}

EXTRACTION_SYSTEM_PROMPT = (
    "Ты извлекаешь структурированные данные из технического задания на услуги "
    "по уборке помещений (для последующего автоматического расчёта сметы). "
    "Извлекай факты ТОЧНО как написано в документе — не додумывай, НЕ вычисляй "
    "площадь объекта самостоятельно (для этого отдельный детерминированный расчёт "
    "после тебя), не оценивай количество расходников, если оно не указано явным "
    "числом (оставляй explicit_qty = null). Если какого-то раздела в документе "
    "нет — верни пустой список/null для него, не выдумывай данные."
)


def extract_tz_structured(
    doc: IngestedDocument, api_key: Optional[str] = None, model: str = "claude-sonnet-5"
) -> dict:
    """
    Единственная точка входа для шага 4. Требует пакет `anthropic` и ключ API
    (передайте явно или через переменную окружения ANTHROPIC_API_KEY).
    Принудительный tool_choice — чтобы Claude не мог ответить обычным текстом
    вместо структуры.
    """
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model=model,
        max_tokens=8192,
        system=EXTRACTION_SYSTEM_PROMPT,
        tools=[EXTRACTION_TOOL],
        tool_choice={"type": "tool", "name": "record_tz_extraction"},
        messages=[{"role": "user", "content": doc.full_text}],
    )
    for block in response.content:
        if block.type == "tool_use" and block.name == "record_tz_extraction":
            return block.input
    raise RuntimeError("Claude не вернул ожидаемый tool_use блок — проверьте ответ вручную.")


# ============================== Детерминированная пост-обработка ==============================

def compute_real_area(volumes: list[float]) -> float:
    """
    НОД объёмов по периодам = площадь объекта. Работает, потому что
    количества дней уборки в разных месяцах (15, 19, 20, 21, 22, 23...) в
    большинстве своём взаимно просты — их НОД равен 1, поэтому НОД объёмов
    (площадь × дни) восстанавливает именно площадь, а не какой-то другой
    общий делитель. Проверено на реальных данных ТЗ — см. demo.
    """
    fracs = [Fraction(v).limit_denominator(1000) for v in volumes]
    common_denom = reduce(lambda a, b: a * b // gcd(a, b), [f.denominator for f in fracs])
    ints = [int(f * common_denom) for f in fracs]
    g = reduce(gcd, ints)
    return float(Fraction(g, common_denom))


def compute_cleaning_days(volumes: list[float], area: float) -> int:
    """Сумма (объём_периода / площадь) по всем периодам — дни уборки за весь контракт."""
    return round(sum(v / area for v in volumes))


def _parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def compute_contract_months(periods: list[dict]) -> int:
    if not periods:
        return 0
    start = min(_parse_date(p["period_start"]) for p in periods)
    end = max(_parse_date(p["period_end"]) for p in periods)
    return (end.year - start.year) * 12 + (end.month - start.month) + 1


def estimate_general_days(cleaning_days: int, general_per_week: float, workdays_per_week: float) -> int:
    """Пропорция от общего числа дней уборки — та же логика, что и в ручном
    расчёте (742 дня уборки × 1/5 ≈ 148 генеральных, с поправкой на то, что
    часть недель выпадает на праздники наравне с остальными)."""
    if not workdays_per_week:
        return 0
    return round(cleaning_days * general_per_week / workdays_per_week)


def build_tz_items(extraction: dict) -> list[TZItem]:
    """
    Список расходников -> TZItem. qty=None, если explicit_qty не был указан
    в самом ТЗ — честно, а не выдуманное число; такие позиции дальше
    дополняются нормами из knowledge_base (шаг 6), как и раньше.
    """
    items = []
    for c in extraction.get("consumables", []):
        name = c["raw_name"]
        if c.get("characteristics"):
            name = f"{name}. {c['characteristics']}"
        items.append(TZItem(raw_name=name, qty=c.get("explicit_qty"), unit=c.get("unit", "шт")))
    return items


@dataclass
class ExtractedObjectSummary:
    object_name: str
    region: str
    address: Optional[str]
    area_sqm: float
    cleaning_days: int
    general_days: int
    contract_months: int
    schedule: dict
    explicit_staff_count: Optional[int]
    staff_count: int              # итоговое число к расчёту: явное из ТЗ, иначе — по нормативу площади
    staff_count_source: str       # "явно указано в ТЗ" | "норматив Роструда (площадь)"
    tz_items: list[TZItem]


def process_extraction(extraction: dict) -> ExtractedObjectSummary:
    """Собирает весь пост-обработанный результат шага 4 из сырого JSON,
    который вернул Claude (или который вы подставили вручную для теста)."""
    periods = extraction["contract_periods"]
    volumes = [p["service_volume_sqm"] for p in periods]
    area = compute_real_area(volumes)
    cleaning_days = compute_cleaning_days(volumes, area)
    contract_months = compute_contract_months(periods)

    schedule = extraction.get("schedule", {})
    general_days = estimate_general_days(
        cleaning_days,
        schedule.get("general_cleaning_per_week", 0) or 0,
        schedule.get("workdays_per_week", 5) or 5,
    )

    staff = extraction.get("explicit_staff_count")
    explicit_staff_value = staff["value"] if staff else None

    if explicit_staff_value is not None:
        staff_count = explicit_staff_value
        staff_count_source = "явно указано в ТЗ"
    else:
        # Норматив Роструда теперь заполнен (knowledge_base/staff_norms.json) —
        # раньше эта ветка была недостижима (get_staff_estimation_norms()
        # всегда возвращал None). Считаем "уборка чаще раза за смену", если
        # в графике есть и основная, и поддерживающая уборка одновременно.
        from knowledge_base import estimate_staff_count_by_area

        cleanings_per_shift = 2 if (
            schedule.get("daily_main_hours_per_day") and schedule.get("daily_supporting_hours_per_day")
        ) else 1
        staff_count = estimate_staff_count_by_area(area, cleanings_per_shift=cleanings_per_shift)
        staff_count_source = "норматив Роструда (площадь)"

    return ExtractedObjectSummary(
        object_name=extraction["object_name"],
        region=extraction["region"],
        address=extraction.get("address"),
        area_sqm=area,
        cleaning_days=cleaning_days,
        general_days=general_days,
        contract_months=contract_months,
        schedule=schedule,
        explicit_staff_count=explicit_staff_value,
        staff_count=staff_count,
        staff_count_source=staff_count_source,
        tz_items=build_tz_items(extraction),
    )


# ============================== Второй тип ТЗ: уборка ТЕРРИТОРИИ ==============================
#
# Проверка на реальном втором документе (Уборка территории ГБУ «ГЦПиКР») показала:
# схема выше НЕ подходит. Там "Перечень объектов закупки" устроен иначе:
#   - НЕТ помесячной разбивки одной и той же услуги (как было в первом ТЗ) —
#     вместо этого каждая строка = ОТДЕЛЬНЫЙ вид услуги (уборка снега, полив,
#     мойка, очистка урны, посыпка противогололедными материалами,
#     санитарное содержание, подметание) с СЕЗОНОМ (зимний/летний) и СВОИМ
#     объёмом (площадью);
#   - "Срок" — это не месяц, а почти всегда целый год (иногда часть года на
#     стыке контракта), то есть приём "разделить объёмы и найти НОД" здесь
#     неприменим — нет повторяющихся дробных периодов одной и той же услуги,
#     которые можно было бы поделить, чтобы вскрыть площадь;
#   - объём в строке — это, вероятнее всего, УЖЕ площадь, подлежащая
#     конкретному виду обработки (а не площадь × количество обработок), но
#     это неочевидно без дополнительных норм (см. Постановление Правительства
#     Москвы №1018, на которое ссылается сам документ, п.7.10) — как часто
#     нужно, например, убирать снег "по заявке" в течение зимнего периода,
#     в самом ТЗ числом не указано.
#
# Поэтому: для помещений (первый тип) есть рабочая формула расчёта площади и
# дней уборки. Для территорий (второй тип) — только извлечение структуры,
# смета по нормам расхода для НЕЁ ещё не построена (это открытый TODO,
# аналогичный staff_norms.json/salary_rules.json — нужны нормативы по видам
# работ на территории, а не на площадь помещения).
#
# ВАЖНО (инженерный урок из этого теста, не только про схему): на этом
# документе `pdftotext -layout` рвёт многострочную ячейку "Срок" на две
# части, раскиданные по разным строкам вывода (несовпадающее число строк в
# соседних колонках "Характеристики"/"Срок"). Структурные таблицы pdfplumber
# (IngestedDocument.tables) эту ячейку сохраняют целиком и НЕ ломаются — при
# извлечении из документов с многострочными ячейками разной длины в соседних
# колонках нужно опираться на doc.tables, а не на full_text.

# ВАЖНО — исправлено после сверки с реальной официальной сметой заказчика
# (см. чат: контракт ГБУ "ГЦПиКР", файл territory_gesn_rates.json): площадь
# НЕЛЬЗЯ брать из "Перечень объектов закупки" (Приложение 1) — те числа служат
# другой цели (нормирование по ст.19 44-ФЗ) и НЕ являются площадью участков.
# Официальная смета использовала формулу вида "8,402=(490,9+282,3+31,7+35,3)/100"
# — реальные площади суммируются из ОТДЕЛЬНОГО приложения "Перечень территорий,
# прилегающих к нежилым зданиям, подлежащих санитарному содержанию (уборке)"
# (в разных ТЗ может называться Приложением 3 или 5 — искать по этому заголовку,
# не по номеру). Это ТОЧНЫЙ аналог урока про помещения: там тоже площадь не в
# "Перечне объектов закупки", а в "Перечне помещений" (обычно Приложение 3).
EXTRACTION_TOOL_TERRITORY = {
    "name": "record_territory_extraction",
    "description": "Записать структурированные данные из ТЗ на уборку территории (не помещений).",
    "input_schema": {
        "type": "object",
        "properties": {
            "object_name": {"type": "string"},
            "region": {"type": "string"},
            "address": {"type": "string"},
            "kpgz_code": {"type": ["string", "null"]},
            "service_periods": {
                "type": "array",
                "description": (
                    "ИСТОЧНИК ПЛОЩАДИ — приложение с заголовком вида 'Перечень территорий, "
                    "прилегающих к нежилым зданиям, подлежащих санитарному содержанию (уборке)' "
                    "(НЕ 'Перечень объектов закупки' — там другие числа, не площадь, см. выше). "
                    "Если это приложение обрезано/не полностью видно в тексте документа — "
                    "верни service_periods пустым и явно укажи это в отдельном поле warning, "
                    "не подставляй числа из 'Перечня объектов закупки' вместо площади."
                ),
                "items": {
                    "type": "object",
                    "properties": {
                        "service_type": {"type": "string", "description": "'Вид услуги': уборка снега/полив/мойка/очистка урны/санитарное содержание/подметание/посыпка противогололедными материалами"},
                        "season": {"type": ["string", "null"], "description": "'Период уборки территории': зимний/летний"},
                        "area_sqm": {"type": "number"},
                        "period_start": {"type": "string", "description": "YYYY-MM-DD"},
                        "period_end": {"type": "string", "description": "YYYY-MM-DD"},
                    },
                    "required": ["service_type", "area_sqm", "period_start", "period_end"],
                },
            },
            "extraction_warning": {
                "type": ["string", "null"],
                "description": "Если источник площади (см. выше) не найден в тексте документа целиком — опиши это здесь, не молчи об этом.",
            },
            "on_demand": {"type": "boolean", "description": "true, если 'Способ оказания услуг: По заявке' (без фиксированного графика/частоты)"},
        },
        "required": ["object_name", "region", "service_periods"],
    },
}


def find_best_area_candidate(volumes: list[float], tolerance: float = 0.005) -> tuple[float, int, int]:
    """
    Устойчивый к "грязным" значениям поиск площади внутри одной группы (один
    вид услуги), когда группа фактически смешивает несколько подучастков —
    ровно ситуация, которую compute_territory_area_by_service() не решает
    (пуловый НОД по всей группе сразу рушится от одного несовпадающего
    значения). Здесь вместо НОД по всей группе — каждое значение группы
    пробуется как площадь-кандидат, и считается, сколько ДРУГИХ значений
    делятся на него нацело (в пределах tolerance). Побеждает кандидат с
    наибольшей поддержкой — при равенстве предпочитается наименьший (более
    вероятна "базовая единица", а не уже домноженное значение).

    Возвращает (площадь, поддержка, всего_значений_в_группе).
    """
    best = (volumes[0], 0)
    for v in volumes:
        support = 0
        for other in volumes:
            ratio = other / v
            if abs(ratio - round(ratio)) < tolerance and round(ratio) >= 1:
                support += 1
        if support > best[1] or (support == best[1] and v < best[0]):
            best = (v, support)
    return best[0], best[1], len(volumes)


def compute_territory_area_by_service(
    periods: list[dict], min_group_size: int = 2, tolerance: float = 0.005
) -> dict[str, dict]:
    """
    Для каждого вида услуги пробует найти площадь-кандидат устойчивым
    методом (find_best_area_candidate). Возвращает диагностику, а не только
    число — сколько строк группы объяснились этой площадью, сколько нет,
    чтобы было видно, надёжен вывод или нет, а не только "да/нет".
    """
    from collections import defaultdict

    by_service: dict[str, list[float]] = defaultdict(list)
    for p in periods:
        by_service[p["service_type"]].append(p["area_sqm"])

    result: dict[str, dict] = {}
    for service_type, volumes in by_service.items():
        if len(volumes) < min_group_size:
            result[service_type] = {"area": None, "support": 0, "total": len(volumes), "reason": "мало строк"}
            continue
        area, support, total = find_best_area_candidate(volumes, tolerance)
        result[service_type] = {
            "area": area if support >= 2 else None,
            "support": support,
            "total": total,
            "reason": None if support >= 2 else "нет совпадений — вероятно, разные площадки без общего делителя",
        }
    return result


@dataclass
class TerritoryServicePeriod:
    service_type: str
    season: Optional[str]
    area_sqm: float
    period_start: date
    period_end: date


@dataclass
class ExtractedTerritorySummary:
    object_name: str
    region: str
    address: Optional[str]
    on_demand: bool
    periods: list[TerritoryServicePeriod]
    service_types: dict[str, float]   # вид услуги -> суммарная площадь по всем сезонам/годам


def process_territory_extraction(extraction: dict) -> ExtractedTerritorySummary:
    periods = [
        TerritoryServicePeriod(
            service_type=p["service_type"],
            season=p.get("season"),
            area_sqm=p["area_sqm"],
            period_start=_parse_date(p["period_start"]),
            period_end=_parse_date(p["period_end"]),
        )
        for p in extraction["service_periods"]
    ]
    service_types: dict[str, float] = {}
    for p in periods:
        service_types[p.service_type] = service_types.get(p.service_type, 0) + p.area_sqm

    return ExtractedTerritorySummary(
        object_name=extraction["object_name"],
        region=extraction["region"],
        address=extraction.get("address"),
        on_demand=extraction.get("on_demand", False),
        periods=periods,
        service_types=service_types,
    )


def detect_document_type(doc: IngestedDocument) -> str:
    """
    Грубый роутер "какую схему извлечения применить" — до вызова Claude.
    Ключевые слова читаются из knowledge_base/tz_document_templates.json, а
    не захардкожены здесь — при появлении нового типа ТЗ достаточно
    дополнить файл, не трогая код.

    ВАЖНО: ищем по строке "Объект закупки:", а не по первым N символам
    документа целиком — тестирование показало, что служебный текст кода
    КПГЗ (иерархия классификатора) упоминает и "зданий, помещений", и
    "территорий" в ОБОИХ типах документов (это общая родительская
    категория в классификаторе), так что широкий скан по всему началу
    документа путает типы. Название конкретного объекта закупки — надёжнее.
    """
    import json
    import re
    from pathlib import Path

    templates_path = Path(__file__).parent / "knowledge_base" / "tz_document_templates.json"
    with open(templates_path, encoding="utf-8") as f:
        templates = json.load(f)["templates"]

    m = re.search(r"Объект закупки:\s*(.+?)(?:\n\s*\n|\n\s*1\.2\.)", doc.full_text, re.DOTALL)
    title_text = m.group(1) if m else doc.full_text[:500]

    # Заголовок объекта закупки бывает слишком лаконичным (например, "уборка
    # ...корпуса и прилегающей территории" без слова "помещений" вообще) —
    # список приложений надёжнее выдаёт гибридные типы (если среди приложений
    # ЕСТЬ И "Перечень помещений", И "Перечень территорий" — это гибрид).
    appendix_match = re.search(
        r"Приложения к Техническому заданию:(.+?)(?:\n\s*\n\s*\S|\Z)", doc.full_text, re.DOTALL
    )
    appendix_text = appendix_match.group(1) if appendix_match else ""

    target_text = (title_text + "\n" + appendix_text).lower()

    for template in templates:
        if not all(stem in target_text for stem in template["detection_required_all"]):
            continue
        pairs = template.get("detection_any_of_pairs")
        if pairs and not any(all(stem in target_text for stem in pair) for pair in pairs):
            continue
        none_of = template.get("detection_none_of")
        if none_of and any(stem in target_text for stem in none_of):
            continue
        return template["type_id"]
    return "unknown"
