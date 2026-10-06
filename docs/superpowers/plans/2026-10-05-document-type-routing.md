# Маршрутизация типов ТЗ + новый тип (Этап 1b) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Подключить уже существующий, но неиспользуемый роутер типов ТЗ (`detect_document_type()`) к конвейеру и добавить поддержку нового типа документа — «помещения, площадь и периодичность в одной таблице» (без раздела «Перечень объектов закупки»), который не покрывается текущей схемой извлечения.

**Architecture:** `detect_document_type()` получает новую возможность исключающего признака (`detection_none_of`), в реестр `tz_document_templates.json` добавляется новый `type_id`. В `tz_extraction.py` добавляется вторая схема извлечения + функция пост-обработки для нового типа, возвращающая тот же `ExtractedObjectSummary`, что и существующая. `smeta_pipeline.generate_smeta()` вызывает детектор и выбирает нужную пару извлечение/пост-обработка; для неподдерживаемых типов — явная `PipelineError`. `app.py` получает одно новое поле формы (срок контракта вручную, нужен только новому типу).

**Tech Stack:** Python 3.14, существующий стек проекта (anthropic, openpyxl, streamlit) без новых зависимостей.

**Spec:** `docs/superpowers/specs/2026-10-05-document-type-routing-design.md`

## Global Constraints

- `territory_cleaning` и `premises_and_territory_cleaning` — детектор их узнаёт, но `generate_smeta()` для них возвращает явную `PipelineError` «не поддерживается», никакого расчёта. Полноценный расчётный движок для territory — вне объёма этого плана.
- Никогда не выдумывать числа: неподдерживаемый тип / отсутствующие данные → явная ошибка или предупреждение, не тихая подстановка.
- Автоматических тестов (pytest) не вводится — в проекте их нет нигде; проверка — ручная/скриптовая (как в предыдущих планах).
- `process_premises_direct_area_extraction()` обязана возвращать `ExtractedObjectSummary` (тот же dataclass, что и `process_extraction()`) — остальной конвейер не должен знать о различии типов документов после этого шага.
- Существующее поведение для `premises_cleaning` не меняется (регрессия проверяется явно).
- Тестовый документ с реальными данными клиента (конфиденциально, не хранится в репозитории, `/tmp/real_tz.docx`) используется только для локальной проверки — НЕ коммитится в git ни в каком виде (ни сам файл, ни его полный текст).

---

## Task 1: Расширение детектора типа документа + реестр

**Files:**
- Modify: `tz_extraction.py:478-485` (цикл сопоставления внутри `detect_document_type()`)
- Modify: `knowledge_base/tz_document_templates.json` (добавить новую запись в `templates[]`)

**Interfaces:**
- Produces: `detect_document_type(doc: IngestedDocument) -> str` — расширенная версия, поддерживает `detection_none_of` в записях реестра, возвращает `"premises_cleaning_direct_area"` для нового типа. Используется в Task 3.

- [ ] **Step 1: Добавить поддержку `detection_none_of` в `detect_document_type()`**

Открыть `tz_extraction.py`, найти в теле `detect_document_type()` (строка ~478) блок:

```python
    for template in templates:
        if not all(stem in target_text for stem in template["detection_required_all"]):
            continue
        pairs = template.get("detection_any_of_pairs")
        if pairs and not any(all(stem in target_text for stem in pair) for pair in pairs):
            continue
        return template["type_id"]
    return "unknown"
```

Заменить на:

```python
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
```

- [ ] **Step 2: Добавить новую запись в реестр типов**

Открыть `knowledge_base/tz_document_templates.json`. Найти массив `"templates": [`. Вставить новый объект **первым элементом** массива (перед существующим `premises_and_territory_cleaning`):

```json
    {
      "type_id": "premises_cleaning_direct_area",
      "display_name": "Уборка помещений (площадь и периодичность в одной таблице, без area-trap)",
      "detection_required_all": ["убор", "помещен"],
      "detection_none_of": ["перечень объектов закупки"],
      "known_examples": [
        "Реальный ТЗ клиента (конфиденциально, не хранится в репозитории) — площадь по помещениям + периодичность в одной таблице, без раздела 'Перечень объектов закупки'"
      ],
      "typical_appendices": {},
      "area_extraction_method": "Площадь указана напрямую построчно по помещениям в одной таблице вместе с периодичностью уборки каждого помещения — суммируется без НОД/area-trap, см. EXTRACTION_TOOL_PREMISES_DIRECT_AREA / process_premises_direct_area_extraction() в tz_extraction.py.",
      "staffing": "Явной численности в тексте обычно нет — норматив Роструда по площади (staff_norms.json), как и для premises_cleaning.",
      "code_path": "tz_extraction.py: EXTRACTION_TOOL_PREMISES_DIRECT_AREA / process_premises_direct_area_extraction()",
      "cost_reference": null
    },
```

(Запятая в конце — после неё идёт существующая запись `premises_and_territory_cleaning`.) Проверить, что JSON остаётся валидным (см. Step 3).

- [ ] **Step 3: Создать тестовые фикстуры и проверить классификацию локально (без live API — `detect_document_type()` не обращается к Claude)**

Реальный документ клиента уже сохранён контроллером на диске отдельно от git — путь `/tmp/real_tz.docx` (если файла нет на момент выполнения, запросить его у контроллера/пользователя, не выдумывать содержимое).

Создать/обновить синтетическую regression-фикстуру для `premises_cleaning` — она должна содержать буквальную фразу «Перечень объектов закупки» (как настоящие документы этого типа), иначе новый `detection_none_of` не отличит её от нового типа:

```bash
.venv/bin/python3 -c "
from docx import Document

doc = Document()
lines = [
    'Техническое задание на оказание услуг по уборке помещений.',
    'Объект: Оказание услуг по уборке офисных помещений в 2027 году.',
    'Регион: Москва. Адрес: г. Москва, ул. Тестовая, д. 1.',
    'Код КПГЗ: 03.08.01.01.01.07.',
    '',
    'Перечень объектов закупки (Приложение 1). Объёмы оказания услуг по периодам:',
    'Январь 2027: 3100 кв.м.',
    'Февраль 2027: 2800 кв.м.',
    'Март 2027: 3100 кв.м.',
    '',
    'График уборки: основная уборка 5 раз в неделю, генеральная уборка 1 раз в неделю.',
    'Расходные материалы: мыло жидкое — не указано количество явно; туалетная бумага — 500 рулонов.',
]
for line in lines:
    doc.add_paragraph(line)
doc.save('/tmp/sample_tz.docx')
print('saved')
"
```

Запустить проверку классификации на обоих документах:

```bash
.venv/bin/python3 -c "
from document_ingestion import ingest_document
from tz_extraction import detect_document_type

doc1 = ingest_document('/tmp/sample_tz.docx')
print('sample_tz.docx ->', detect_document_type(doc1))

doc2 = ingest_document('/tmp/real_tz.docx')
print('real_tz.docx ->', detect_document_type(doc2))
"
```

Expected:
```
sample_tz.docx -> premises_cleaning
real_tz.docx -> premises_cleaning_direct_area
```

Если результат другой — не переходить к Step 4, разобраться (скорее всего проблема в точности ключевых фраз/регистре).

- [ ] **Step 4: Commit**

```bash
git add tz_extraction.py knowledge_base/tz_document_templates.json
git commit -m "Добавить detection_none_of в детектор типа ТЗ + зарегистрировать тип premises_cleaning_direct_area"
```

---

## Task 2: Новая схема извлечения и пост-обработка для типа `premises_cleaning_direct_area`

**Files:**
- Modify: `tz_extraction.py` (добавить константу, функцию извлечения и функцию пост-обработки — append в конец файла, после `detect_document_type()`)

**Interfaces:**
- Consumes: `IngestedDocument` (из `document_ingestion.py`, без изменений), `ExtractedObjectSummary` dataclass (определён в `tz_extraction.py:191-204`, без изменений), `build_tz_items(extraction: dict) -> list[TZItem]` (существует, `tz_extraction.py:176`), `estimate_staff_count_by_area(area_sqm, cleanings_per_shift=1) -> Optional[int]` (из `knowledge_base.py:262`)
- Produces:
  - `EXTRACTION_TOOL_PREMISES_DIRECT_AREA: dict` — схема для Anthropic tool-use
  - `extract_tz_structured_direct_area(doc: IngestedDocument, api_key: Optional[str] = None, model: str = "claude-sonnet-5") -> dict` — для Task 3
  - `process_premises_direct_area_extraction(extraction: dict, contract_months: int) -> ExtractedObjectSummary` — для Task 3

- [ ] **Step 1: Добавить схему извлечения и функцию вызова API**

Добавить в конец `tz_extraction.py` (после `detect_document_type()`):

```python


# ============================== Третий тип ТЗ: помещения, площадь+периодичность в одной таблице ==============================
#
# Обнаружено на реальном документе клиента: нет раздела "Перечень объектов
# закупки" (area-trap) вообще — площадь и периодичность уборки заданы
# напрямую построчно по помещениям в одной таблице. detect_document_type()
# различает этот тип от premises_cleaning по ОТСУТСТВИЮ фразы "перечень
# объектов закупки" (см. detection_none_of в tz_document_templates.json).

EXTRACTION_TOOL_PREMISES_DIRECT_AREA = {
    "name": "record_premises_direct_area_extraction",
    "description": (
        "Записать структурированные данные из ТЗ на уборку помещений, где "
        "площадь и периодичность уборки заданы НАПРЯМУЮ построчно по "
        "помещениям в одной таблице (не через накопленный объём услуг за "
        "период)."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "object_name": {"type": "string"},
            "region": {"type": "string"},
            "address": {"type": "string"},
            "kpgz_code": {"type": ["string", "null"]},
            "rooms": {
                "type": "array",
                "description": "Строки таблицы с площадью и периодичностью по каждому помещению.",
                "items": {
                    "type": "object",
                    "properties": {
                        "room_name": {"type": "string"},
                        "area_sqm": {"type": "number"},
                        "periodicity_raw": {
                            "type": "string",
                            "description": "Текст периодичности КАК В ДОКУМЕНТЕ, например '1 раз в день'.",
                        },
                        "cleaning_times_per_month": {
                            "type": ["number", "null"],
                            "description": (
                                "Переведи periodicity_raw в среднее число уборок В МЕСЯЦ "
                                "(ежедневно по будням ≈ 21.7; раз в неделю ≈ 4.33; раз в "
                                "2 дня ≈ 10.8 и т.д. — это перевод единиц, а не оценка, "
                                "НЕ придумывай число, если периодичность не указана явно "
                                "текстом — тогда верни null)."
                            ),
                        },
                    },
                    "required": ["room_name", "area_sqm", "periodicity_raw"],
                },
            },
            "general_cleaning_per_week": {
                "type": ["number", "null"],
                "description": "Отдельная периодичность генеральной уборки в неделю, если явно указана, иначе null.",
            },
            "explicit_staff_count": {
                "type": ["object", "null"],
                "properties": {
                    "value": {"type": "integer"},
                    "condition": {"type": "string"},
                },
            },
            "consumables": {
                "type": "array",
                "description": "Из раздела/таблицы расходных материалов.",
                "items": {
                    "type": "object",
                    "properties": {
                        "raw_name": {"type": "string"},
                        "characteristics": {"type": "string"},
                        "unit": {"type": "string"},
                        "explicit_qty": {"type": ["number", "null"]},
                    },
                    "required": ["raw_name", "unit"],
                },
            },
        },
        "required": ["object_name", "region", "rooms", "consumables"],
    },
}


def extract_tz_structured_direct_area(
    doc: IngestedDocument, api_key: Optional[str] = None, model: str = "claude-sonnet-5"
) -> dict:
    """Аналог extract_tz_structured(), но для типа premises_cleaning_direct_area."""
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model=model,
        max_tokens=8192,
        system=EXTRACTION_SYSTEM_PROMPT,
        tools=[EXTRACTION_TOOL_PREMISES_DIRECT_AREA],
        tool_choice={"type": "tool", "name": "record_premises_direct_area_extraction"},
        messages=[{"role": "user", "content": doc.full_text}],
    )
    for block in response.content:
        if block.type == "tool_use" and block.name == "record_premises_direct_area_extraction":
            return block.input
    raise RuntimeError("Claude не вернул ожидаемый tool_use блок — проверьте ответ вручную.")


def process_premises_direct_area_extraction(
    extraction: dict, contract_months: int,
) -> ExtractedObjectSummary:
    """
    Площадь — прямая сумма (без area-trap/НОД, в отличие от
    process_extraction()). «Дней уборки в месяц» — площадь-взвешенное
    среднее cleaning_times_per_month по помещениям: ОЦЕНКА смешанной
    периодичности одним числом для совместимости с build_materials_table()
    (которая ожидает одну базу на весь объект), не точный факт из
    документа. contract_months нельзя вычислить из этого типа документа
    (нет таблицы периодов с датами) — передаётся явно вызывающим кодом.
    """
    rooms = extraction["rooms"]
    area = sum(r["area_sqm"] for r in rooms)

    weighted = [r for r in rooms if r.get("cleaning_times_per_month") is not None]
    if weighted:
        total_weighted_area = sum(r["area_sqm"] for r in weighted)
        cleaning_days = round(
            sum(r["area_sqm"] * r["cleaning_times_per_month"] for r in weighted)
            / total_weighted_area
        ) if total_weighted_area else 0
    else:
        cleaning_days = 0

    general_days = round(extraction.get("general_cleaning_per_week") or 0)

    staff = extraction.get("explicit_staff_count")
    explicit_staff_value = staff["value"] if staff else None
    if explicit_staff_value is not None:
        staff_count = explicit_staff_value
        staff_count_source = "явно указано в ТЗ"
    else:
        from knowledge_base import estimate_staff_count_by_area
        staff_count = estimate_staff_count_by_area(area, cleanings_per_shift=1)
        staff_count_source = "норматив Роструда (площадь)"

    tz_items = build_tz_items(extraction)

    return ExtractedObjectSummary(
        object_name=extraction["object_name"],
        region=extraction["region"],
        address=extraction.get("address"),
        area_sqm=area,
        cleaning_days=cleaning_days,
        general_days=general_days,
        contract_months=contract_months,
        schedule={"general_cleaning_per_week": extraction.get("general_cleaning_per_week")},
        explicit_staff_count=explicit_staff_value,
        staff_count=staff_count,
        staff_count_source=staff_count_source,
        tz_items=tz_items,
    )
```

- [ ] **Step 2: Проверить вживую на реальном документе (живой вызов Claude API)**

```bash
.venv/bin/python3 -c "
from dotenv import load_dotenv
load_dotenv()
from document_ingestion import ingest_document
from tz_extraction import extract_tz_structured_direct_area, process_premises_direct_area_extraction

doc = ingest_document('/tmp/real_tz.docx')
extraction = extract_tz_structured_direct_area(doc)
print('rooms found:', len(extraction['rooms']))
print('consumables found:', len(extraction['consumables']))

summary = process_premises_direct_area_extraction(extraction, contract_months=12)
print('object_name:', summary.object_name)
print('area_sqm:', summary.area_sqm)
print('cleaning_days:', summary.cleaning_days)
print('general_days:', summary.general_days)
print('staff_count:', summary.staff_count, '|', summary.staff_count_source)
print('tz_items count:', len(summary.tz_items))
"
```

Expected: без исключений. `area_sqm` ≈ 1487.0 (371.0 + 1116.0 — сумма площадей двух адресов из документа). `rooms found` и `consumables found` — положительные числа (не ноль). `cleaning_days` — положительное число (большинство помещений убираются ежедневно, поэтому ожидается число, близкое к ~20).

- [ ] **Step 3: Commit**

```bash
git add tz_extraction.py
git commit -m "Добавить EXTRACTION_TOOL_PREMISES_DIRECT_AREA и process_premises_direct_area_extraction() — новый тип ТЗ"
```

---

## Task 3: Диспетчер в `smeta_pipeline.generate_smeta()`

**Files:**
- Modify: `smeta_pipeline.py:23` (импорт)
- Modify: `smeta_pipeline.py:58-96` (сигнатура функции + блок извлечения/пост-обработки + блок предупреждений)

**Interfaces:**
- Consumes: `detect_document_type`, `extract_tz_structured_direct_area`, `process_premises_direct_area_extraction` (все из Task 1/2, `tz_extraction.py`)
- Produces: `generate_smeta(file_path, *, region=, object_complexity=, schedule_complexity=, vat_rate=, output_dir=, contract_months_override=None) -> SmetaPipelineResult` — новый необязательный параметр `contract_months_override`, используется в Task 4 (`app.py`)

- [ ] **Step 1: Обновить импорт**

В `smeta_pipeline.py` заменить строку 23:

```python
from tz_extraction import extract_tz_structured, process_extraction
```

на:

```python
from tz_extraction import (
    detect_document_type, extract_tz_structured, extract_tz_structured_direct_area,
    process_extraction, process_premises_direct_area_extraction,
)
```

- [ ] **Step 2: Добавить параметр `contract_months_override` в сигнатуру**

Заменить (строки 58-69):

```python
def generate_smeta(
    file_path: str,
    *,
    region: str = "Москва",
    object_complexity: str = "стандартный",
    schedule_complexity: str = (
        "комбинированная смена (ежедневная-основная + "
        "ежедневная-поддерживающая, один сотрудник)"
    ),
    vat_rate: float = 0.20,
    output_dir: str = "Outputs/tender-agent-smeta",
) -> SmetaPipelineResult:
```

на:

```python
def generate_smeta(
    file_path: str,
    *,
    region: str = "Москва",
    object_complexity: str = "стандартный",
    schedule_complexity: str = (
        "комбинированная смена (ежедневная-основная + "
        "ежедневная-поддерживающая, один сотрудник)"
    ),
    vat_rate: float = 0.20,
    output_dir: str = "Outputs/tender-agent-smeta",
    contract_months_override: Optional[int] = None,
) -> SmetaPipelineResult:
```

- [ ] **Step 3: Заменить блок извлечения на диспетчер по типу документа**

Заменить (строки 80-90):

```python
    # 2. Живой вызов Claude API — структурированное извлечение
    extraction = extract_tz_structured(doc)

    # 3. Детерминированная пост-обработка
    if not extraction.get("contract_periods"):
        raise PipelineError(
            "В документе не найдены объёмы услуг по периодам (раздел "
            "«Перечень объектов закупки»). Автоматический расчёт "
            "невозможен — нужна ручная проверка документа."
        )
    summary = process_extraction(extraction)
```

на:

```python
    # 2. Определить тип документа и выбрать схему извлечения
    doc_type = detect_document_type(doc)

    if doc_type == "premises_cleaning":
        extraction = extract_tz_structured(doc)
        if not extraction.get("contract_periods"):
            raise PipelineError(
                "В документе не найдены объёмы услуг по периодам (раздел "
                "«Перечень объектов закупки»). Автоматический расчёт "
                "невозможен — нужна ручная проверка документа."
            )
        summary = process_extraction(extraction)

    elif doc_type == "premises_cleaning_direct_area":
        if contract_months_override is None:
            raise PipelineError(
                "Для этого типа ТЗ срок контракта не задан в виде таблицы "
                "периодов — укажите срок контракта в месяцах в форме и "
                "повторите расчёт."
            )
        extraction = extract_tz_structured_direct_area(doc)
        if not extraction.get("rooms"):
            raise PipelineError(
                "В документе не найдены строки с площадью и периодичностью "
                "по помещениям. Автоматический расчёт невозможен — нужна "
                "ручная проверка документа."
            )
        summary = process_premises_direct_area_extraction(extraction, contract_months_override)

    else:
        raise PipelineError(
            f"Тип ТЗ «{doc_type}» пока не поддерживается автоматическим "
            "расчётом сметы (поддерживаются только документы с прямой "
            "площадью по помещениям или с накопленным объёмом услуг за "
            "период). Нужна ручная проверка документа."
        )
```

- [ ] **Step 4: Добавить предупреждения для нового типа**

Найти блок (сразу после только что изменённого, оригинальные строки 92-96):

```python
    if summary.staff_count_source == "норматив Роструда (площадь)":
        from knowledge_base import load_staff_norms
        scope_warning = load_staff_norms().get("_scope_warning")
        if scope_warning:
            warnings.append(f"Численность оценена по нормативу площади: {scope_warning}")
```

Добавить непосредственно ПЕРЕД этим блоком:

```python
    if doc_type == "premises_cleaning_direct_area":
        warnings.append(
            "«Дней уборки» для этого ТЗ — площадь-взвешенная оценка "
            "смешанной периодичности по помещениям, не точный факт из "
            "документа. Проверьте вручную перед использованием сметы."
        )
        if summary.cleaning_days == 0:
            warnings.append(
                "Не удалось определить периодичность уборки ни для одного "
                "помещения (текст периодичности не распознан) — база "
                "материалов в смете будет занижена, требуется ручная "
                "проверка документа и корректировка жёлтых ячеек."
            )

```

- [ ] **Step 5: Проверить регрессию на `premises_cleaning` (живой вызов API)**

```bash
.venv/bin/python3 -c "
from dotenv import load_dotenv
load_dotenv()
from smeta_pipeline import generate_smeta

result = generate_smeta('/tmp/sample_tz.docx')
print('area_sqm:', result.area_sqm)
print('cleaning_days:', result.cleaning_days)
print('contract_months:', result.contract_months)
print('warnings:', result.warnings)
"
```

Expected: без исключений, `area_sqm` ≈ 100.0, `contract_months` = 3, `warnings` НЕ содержит строк про «площадь-взвешенную оценку» (это регрессия для `premises_cleaning`, не для нового типа).

- [ ] **Step 6: Проверить новый тип через `generate_smeta()` целиком (живой вызов API)**

```bash
.venv/bin/python3 -c "
from dotenv import load_dotenv
load_dotenv()
from smeta_pipeline import generate_smeta

result = generate_smeta('/tmp/real_tz.docx', contract_months_override=12)
print('object_name:', result.object_name)
print('area_sqm:', result.area_sqm)
print('cleaning_days:', result.cleaning_days)
print('contract_months:', result.contract_months)
print('staff_count:', result.staff_count, '|', result.staff_count_source)
print('fot_month:', result.fot_month)
print('materials_unresolved:', result.materials_unresolved)
print('xlsx_path:', result.xlsx_path)
print('warnings:')
for w in result.warnings:
    print(' -', w)
"
```

Expected: без исключений. `area_sqm` ≈ 1487.0, `contract_months` = 12, `warnings` содержит строку про «площадь-взвешенную оценку». `xlsx_path` указывает на файл в `Outputs/tender-agent-smeta/`.

- [ ] **Step 7: Проверить отсутствие ошибки `contract_months_override` при вызове без него**

```bash
.venv/bin/python3 -c "
from dotenv import load_dotenv
load_dotenv()
from smeta_pipeline import generate_smeta, PipelineError

try:
    generate_smeta('/tmp/real_tz.docx')
    print('ОШИБКА: исключение не было поднято')
except PipelineError as e:
    print('OK, PipelineError:', e)
"
```

Expected: печатает `OK, PipelineError: ...` с текстом про срок контракта.

- [ ] **Step 8: Проверить сгенерированный Excel на ошибки формул (LibreOffice recalc)**

```bash
XLSX=$(ls -t Outputs/tender-agent-smeta/*.xlsx | head -1)
mkdir -p /tmp/recalc_check_1b
libreoffice --headless --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir /tmp/recalc_check_1b "$XLSX"
.venv/bin/python3 -c "
import glob
import openpyxl

path = sorted(glob.glob('/tmp/recalc_check_1b/*.xlsx'))[-1]
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
errors = [c.value for row in ws.iter_rows() for c in row
          if isinstance(c.value, str) and c.value.startswith('#')]
print('Ошибки формул:', errors)
"
```

Expected: `Ошибки формул: []`

- [ ] **Step 9: Commit**

```bash
git add smeta_pipeline.py
git commit -m "Подключить диспетчер типов ТЗ в generate_smeta() — поддержка premises_cleaning_direct_area"
```

---

## Task 4: Поле формы «Срок контракта» в `app.py`

**Files:**
- Modify: `app.py:70-90` (форма параметров)
- Modify: `app.py:92-104` (вызов `generate_smeta`)

**Interfaces:**
- Consumes: `generate_smeta(..., contract_months_override=...)` — параметр из Task 3

- [ ] **Step 1: Добавить поле ввода срока контракта в форму**

В `app.py` найти блок (строки 74-85):

```python
with col2:
    schedule_complexity = st.selectbox(
        "График",
        [
            "комбинированная смена (ежедневная-основная + "
            "ежедневная-поддерживающая, один сотрудник)"
        ],
        index=0,
    )
    vat_rate_pct = st.number_input(
        "Ставка НДС, %", min_value=0.0, max_value=100.0, value=20.0, step=1.0,
    )
```

Заменить на:

```python
with col2:
    schedule_complexity = st.selectbox(
        "График",
        [
            "комбинированная смена (ежедневная-основная + "
            "ежедневная-поддерживающая, один сотрудник)"
        ],
        index=0,
    )
    vat_rate_pct = st.number_input(
        "Ставка НДС, %", min_value=0.0, max_value=100.0, value=20.0, step=1.0,
    )
    contract_months_input = st.number_input(
        "Срок контракта, мес (только если не определяется из документа)",
        min_value=0, max_value=120, value=0, step=1,
        help=(
            "Нужно заполнить только для ТЗ, где площадь задана напрямую "
            "по помещениям (без таблицы периодов с датами) — для обычных "
            "ТЗ с «Перечнем объектов закупки» это поле игнорируется, срок "
            "берётся из документа. 0 = не указывать."
        ),
    )
```

- [ ] **Step 2: Передать значение в `generate_smeta()`**

Найти блок (строки 98-104):

```python
            result = generate_smeta(
                tmp_path,
                region=region,
                object_complexity=object_complexity,
                schedule_complexity=schedule_complexity,
                vat_rate=vat_rate_pct / 100,
            )
```

Заменить на:

```python
            result = generate_smeta(
                tmp_path,
                region=region,
                object_complexity=object_complexity,
                schedule_complexity=schedule_complexity,
                vat_rate=vat_rate_pct / 100,
                contract_months_override=contract_months_input or None,
            )
```

- [ ] **Step 3: Проверить компиляцию и локальный запуск сервера**

```bash
.venv/bin/python3 -m py_compile app.py
```
Expected: без ошибок.

```bash
nohup .venv/bin/streamlit run app.py --server.address 127.0.0.1 --server.port 8502 --server.headless true > /tmp/streamlit_1b_check.log 2>&1 &
sleep 5
curl -s -o /dev/null -w "HTTP %{http_code}\n" http://127.0.0.1:8502
pkill -f "streamlit run app.py --server.port 8502"
```
Expected: `HTTP 200`. (Порт 8502, а не 8501, — чтобы не конфликтовать с уже работающей systemd-службой `tender-agent-web.service` на 8501 во время проверки.)

- [ ] **Step 4: Commit**

```bash
git add app.py
git commit -m "Добавить поле «Срок контракта» в форму — для типа ТЗ без таблицы периодов"
```

---

## Self-Review Notes

- **Покрытие спецификации:** Компонент 1 (детектор + `detection_none_of`) — Task 1. Компонент 2 (реестр) — Task 1. Компонент 3 (схема извлечения) — Task 2. Компонент 4 (пост-обработка) — Task 2. Компонент 5 (диспетчер + предупреждения) — Task 3. Компонент 6 (поле формы) — Task 4. Раздел «Тестирование» спеки — регрессия на `premises_cleaning` (Task 3 Step 5), новый тип (Task 2 Step 2, Task 3 Step 6), проверка диспетчера на неподдерживаемом типе (Task 3 Step 7, хоть и через отсутствие `contract_months_override`, а не через territory — предполагается, что отсутствие реального territory-документа делает эту проверку менее приоритетной; логика идентична для обоих случаев отказа).
- **Проверка типов/сигнатур:** `ExtractedObjectSummary` используется идентично в обеих функциях пост-обработки (Task 2 ссылается на поля, определённые в существующем коде `tz_extraction.py:192-204`). `generate_smeta()`'s новый параметр `contract_months_override: Optional[int] = None` используется в Task 4 ровно с этим именем.
- **Приватность:** реальный документ клиента используется только по пути `/tmp/real_tz.docx`, вне worktree — ни один Step не коммитит и не записывает его содержимое в файлы репозитория.
