# Маршрутизация типов ТЗ + новый тип «помещения, площадь в таблице» (Этап 1b)

Дата: 2026-10-05

## Контекст и цель

Живая проверка опубликованного Этапа 1 на реальном документе клиента
(конфиденциально, не хранится в репозитории) показала: `smeta_pipeline.generate_smeta()` всегда использует
ОДНУ схему извлечения (`EXTRACTION_TOOL` / `process_extraction()` из
`tz_extraction.py`), рассчитанную на документы с «Перечнем объектов
закупки» (накопленный объём услуг за период, площадь восстанавливается
через НОД — см. `compute_real_area()`). Реальный документ оказался
документом **другой структуры**: площадь и периодичность уборки заданы
напрямую, построчно по помещениям, в одной таблице — без раздела
«Перечень объектов закупки» вообще.

Обнаружилось также, что в проекте уже существует роутер типов документов
— `detect_document_type()` в `tz_extraction.py`, читающий реестр типов из
`knowledge_base/tz_document_templates.json` — но он **нигде не вызывается**
(подтверждено: `grep` по всему проекту не находит ни одного вызова за
пределами определения функции). Реестр уже содержит 3 типа
(`premises_cleaning`, `territory_cleaning`,
`premises_and_territory_cleaning`), но ни один не соответствует структуре
реального документа — нужен четвёртый тип.

Попутно при живой проверке найден и исправлен отдельный баг в
`document_ingestion.py` (таблицы Word не попадали в текст для Claude —
см. коммит `b3e5606`), который скрывал масштаб проблемы: без этого фикса
расходные материалы для любого похожего документа тоже не извлекались
бы.

## Явно вне объёма этой итерации

- **`territory_cleaning`** (уборка территории) — для него нет вообще
  никакой логики сборки сметы (другая методика — расценки ГЭСН по видам
  работ, не нормативы материалов). Решено отдельно: `detect_document_type()`
  научится его узнавать, но `generate_smeta()` для него будет явно
  сообщать «не поддерживается», а не пытаться считать. Полноценный
  расчётный движок для territory — отдельная задача, когда появится
  реальный документ для проверки.
- **`premises_and_territory_cleaning`** (гибрид) — та же судьба: уже
  зафиксирован в реестре как «схема ещё не написана», ничего не меняем.
- Автоматическое определение произвольных НОВЫХ, ещё не встречавшихся
  типов документов — вне объёма; `detect_document_type()` возвращает
  `"unknown"`, и это явная ошибка, не попытка угадать.

## Архитектура

### Компонент 1 — Расширение `detect_document_type()` (`tz_extraction.py`)

Текущая реализация поддерживает только `detection_required_all` (все
основы слов должны быть в тексте) и `detection_any_of_pairs` (хотя бы одна
пара основ). Этого недостаточно, чтобы отличить новый тип от
`premises_cleaning` — у обоих заголовок «Объект закупки:» содержит
одинаковые ключевые слова («убор», «помещен»). Единственный надёжный
признак различия — **наличие или отсутствие** фразы «перечень объектов
закупки» в тексте документа (у `premises_cleaning` она есть как
заголовок раздела/таблицы с накопленным объёмом; у нового типа её нет
вообще — площадь сразу в теле ТЗ).

Добавляется необязательное поле `detection_none_of` (список основ слов,
которых быть НЕ должно) в формат записи `templates[]` и в логику
`detect_document_type()`:

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

`target_text` уже строится из заголовка + списка приложений (существующая
логика) — этого достаточно, так как фраза «перечень объектов закупки»
обычно встречается именно в списке приложений/заголовке раздела, а не
только внутри таблицы.

### Компонент 2 — Новый тип в реестре (`knowledge_base/tz_document_templates.json`)

Новая запись `premises_cleaning_direct_area`, вставленная в список
`templates[]` **сразу после** `premises_and_territory_cleaning` (гибрид
проверяется первым — специфичные типы проверяются раньше общих, уже
существующее правило, см. `_detection_lesson` в файле; итоговый порядок
скорректирован в ревью, см. коммит `2483852`):

```json
{
  "type_id": "premises_cleaning_direct_area",
  "display_name": "Уборка помещений (площадь и периодичность в одной таблице, без area-trap)",
  "detection_required_all": ["убор", "помещен"],
  "detection_none_of": ["перечень объектов закупки"],
  "known_examples": [
    "Реальный ТЗ клиента (конфиденциально, не хранится в репозитории) — площадь по помещениям + периодичность в одной таблице, без раздела «Перечень объектов закупки»"
  ],
  "typical_appendices": {},
  "area_extraction_method": "Площадь указана напрямую построчно по помещениям в одной таблице вместе с периодичностью уборки каждого помещения — суммируется без НОД/area-trap, см. EXTRACTION_TOOL_PREMISES_DIRECT_AREA / process_premises_direct_area_extraction() в tz_extraction.py.",
  "staffing": "Явной численности в тексте обычно нет — норматив Роструда по площади (staff_norms.json), как и для premises_cleaning.",
  "code_path": "tz_extraction.py: EXTRACTION_TOOL_PREMISES_DIRECT_AREA / process_premises_direct_area_extraction()",
  "cost_reference": null
}
```

### Компонент 3 — Новая схема извлечения (`tz_extraction.py`)

```python
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
                            "type": "number",
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
                "description": "Из раздела/таблицы расходных материалов — та же структура, что и в основной схеме.",
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
```

### Компонент 4 — Детерминированная пост-обработка (`tz_extraction.py`)

```python
def process_premises_direct_area_extraction(
    extraction: dict, contract_months: int,
) -> ExtractedObjectSummary:
    """
    Площадь — прямая сумма (без area-trap/НОД, в отличие от
    process_extraction()). «Дней уборки в месяц» — площадь-взвешенное
    среднее cleaning_times_per_month по помещениям: это ОЦЕНКА смешанной
    периодичности одним числом для совместимости с build_materials_table()
    (которая ожидает одну базу на весь объект), не точный факт — помечается
    предупреждением в смете. contract_months нельзя вычислить из этого типа
    документа (нет таблицы периодов с датами) — передаётся явно вызывающим
    кодом (из формы пользователя).
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

    tz_items = build_tz_items(extraction)  # уже существует, работает с extraction["consumables"]

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

Возвращает ТОТ ЖЕ `ExtractedObjectSummary`, что и `process_extraction()` —
остальной конвейер (`smeta_pipeline.py`) не должен знать о различии типов
после этого шага.

### Компонент 5 — Диспетчер в `smeta_pipeline.generate_smeta()`

```python
def generate_smeta(
    file_path: str,
    *,
    region: str = "Москва",
    object_complexity: str = "стандартный",
    schedule_complexity: str = (...),  # без изменений
    vat_rate: float = 0.20,
    output_dir: str = "Outputs/tender-agent-smeta",
    contract_months_override: Optional[int] = None,
) -> SmetaPipelineResult:
    ...
    doc = ingest_document(file_path)
    if len(doc.full_text.strip()) < 50:
        raise PipelineError(...)

    doc_type = detect_document_type(doc)

    if doc_type == "premises_cleaning":
        extraction = extract_tz_structured(doc)
        if not extraction.get("contract_periods"):
            raise PipelineError(...)
        summary = process_extraction(extraction)

    elif doc_type == "premises_cleaning_direct_area":
        if contract_months_override is None:
            raise PipelineError(
                "Для этого типа ТЗ срок контракта не указан в виде таблицы "
                "периодов — укажите срок контракта в месяцах в форме и "
                "повторите расчёт."
            )
        extraction = extract_tz_structured_direct_area(doc)  # новая функция, аналог extract_tz_structured, но с EXTRACTION_TOOL_PREMISES_DIRECT_AREA
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

    if summary.staff_count_source == "норматив Роструда (площадь)":
        ...  # без изменений
    ...  # остальной конвейер без изменений
```

`extract_tz_structured_direct_area()` — тонкая копия существующей
`extract_tz_structured()`, отличается только `tools=[EXTRACTION_TOOL_PREMISES_DIRECT_AREA]`
и `tool_choice={"name": "record_premises_direct_area_extraction"}`;
остальное (клиент, модель, system-промпт) переиспользуется как есть.

### Компонент 6 — `app.py`: поле «Срок контракта, мес»

Добавляется `st.number_input("Срок контракта, мес (если не определяется из документа)", ...)`
рядом с существующими полями формы, с пояснением, что это нужно только для
документов без таблицы периодов (`premises_cleaning_direct_area`) —
для `premises_cleaning` поле игнорируется (срок по-прежнему берётся из
документа). Значение передаётся в `generate_smeta(..., contract_months_override=...)`.

## Поток данных и обработка ошибок (сводка)

```
ingest_document → detect_document_type()
  ├─ premises_cleaning → extract_tz_structured() → process_extraction()
  ├─ premises_cleaning_direct_area → extract_tz_structured_direct_area() → process_premises_direct_area_extraction()
  │     (требует contract_months_override из формы — иначе PipelineError)
  └─ остальное (territory_cleaning / гибрид / unknown) → PipelineError, явный текст
→ (общий путь без изменений) ФОТ по нормативу → материалы по нормативу → Excel → сохранение
```

Принцип не меняется: отсутствие данных/неподдерживаемый тип → явная
ошибка или предупреждение, никогда не выдуманное число. Площадь-взвешенная
частота уборки для нового типа — это посчитанная (не выдуманная) оценка,
но помечается как оценка (предупреждение в `warnings`), а не точный факт.

## Тестирование

Как и для Этапа 1 — ручная проверка, в проекте нет pytest:
- Регрессия на `premises_cleaning`: тот же синтетический `/tmp/sample_tz.docx`,
  что использовался для Этапа 1 — убедиться, что `detect_document_type()`
  классифицирует его как `premises_cleaning` и результат не изменился.
- Новый тип: реальный документ клиента (конфиденциально, уже есть) —
  убедиться, что `detect_document_type()` вернёт
  `premises_cleaning_direct_area`, площадь посчитается как сумма (≈1487 м²),
  расходники извлекутся (уже подтверждено после фикса таблиц), смета
  соберётся без ошибок формул (LibreOffice recalc).
- Проверка диспетчера на явно неподдерживаемом типе: подать текст с
  ключевым словом «территор» без «перечень объектов закупки» — убедиться,
  что возвращается понятная `PipelineError`, а не попытка посчитать.

## Открытые ограничения (переносятся/добавляются)

- «Дней уборки» для нового типа — площадь-взвешенная ОЦЕНКА смешанной
  периодичности, не точный факт (так помечено в смете).
- `territory_cleaning` и гибрид — определяются роутером, но
  автоматический расчёт для них по-прежнему недоступен (явная ошибка).
- Срок контракта для нового типа — всегда ручной ввод пользователя, не
  извлекается из документа (в документе его может не быть в принципе,
  если дата окончания привязана к дате подписания).
