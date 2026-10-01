# Веб-интерфейс тендер-агента (Этап 1) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Дать нескольким коллегам веб-интерфейс (Streamlit), который принимает ТЗ тендера на уборку, запускает уже проверенный конвейер извлечения/расчёта и отдаёт готовую Excel-смету для скачивания.

**Architecture:** Новый модуль `smeta_pipeline.py` связывает существующие проверенные кирпичики (`document_ingestion`, `tz_extraction`, `knowledge_base`, `smeta_builder`, `price_parser`) в одну функцию `generate_smeta()` для НОВОГО объекта (оценка по нормативам, не по реальным зафиксированным данным выигранных контрактов). Новый файл `app.py` — однофайловый Streamlit-интерфейс поверх неё с простой парольной защитой.

**Tech Stack:** Python 3.14, Streamlit, python-dotenv, openpyxl (уже используется), существующий код проекта без изменений.

**Spec:** `docs/superpowers/specs/2026-10-01-web-interface-design.md`

## Global Constraints

- Streamlit выбран как технология интерфейса (не FastAPI, не Telegram-бот) — см. спецификацию, секция «Архитектура».
- Никогда не выдумывать числа: если норматив (ФОТ, позиция в прайсе) не найден — явное предупреждение в UI и в самой Excel-смете (жёлтая editable-ячейка), а не придуманное значение.
- Автоматических тестов (pytest) не вводится — в проекте их нет нигде; проверка всех задач — ручная/скриптовая, но не через pytest.
- Публичный доступ из интернета — вне объёма; сервер запускается только на `127.0.0.1`, новые порты наружу не открываются.
- `.env` (секреты `ANTHROPIC_API_KEY`, `APP_PASSWORD`) не коммитится в git — уже в `.gitignore`.
- Существующие файлы проекта (`tz_extraction.py`, `smeta_builder.py`, `knowledge_base.py`, `document_ingestion.py`, `price_parser.py` и т.д.) не меняются — только используются как есть.
- Все коммиты — отдельные, с сообщением на русском (как уже в истории проекта).

---

## Task 1: Зависимости — Streamlit и python-dotenv

**Files:**
- Modify: `requirements.txt`

**Interfaces:**
- Produces: установленные пакеты `streamlit`, `python-dotenv` в `.venv` — нужны Task 3 (`app.py`).

- [ ] **Step 1: Добавить зависимости в requirements.txt**

Открыть `requirements.txt` и добавить в конец файла:

```
streamlit>=1.30         # app.py — веб-интерфейс (Этап 1, см. docs/superpowers/specs/2026-10-01-web-interface-design.md)
python-dotenv>=1.0       # app.py — загрузка .env (ANTHROPIC_API_KEY, APP_PASSWORD)
```

- [ ] **Step 2: Установить зависимости**

Run:
```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
source .venv/bin/activate
pip install -r requirements.txt
```
Expected: установка завершается без ошибок (`Successfully installed streamlit ... python-dotenv ...` или `Requirement already satisfied`, если уже стоит).

- [ ] **Step 3: Проверить импорт**

Run:
```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
source .venv/bin/activate
python3 -c "import streamlit, dotenv; print('streamlit', streamlit.__version__)"
```
Expected: печатает версию streamlit, без `ImportError`.

- [ ] **Step 4: Commit**

```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
git add requirements.txt
git commit -m "Добавить streamlit и python-dotenv в зависимости (веб-интерфейс Этап 1)"
```

---

## Task 2: `smeta_pipeline.py` — конвейер «ТЗ → оценочная смета»

**Files:**
- Create: `smeta_pipeline.py`
- Modify: `.gitignore`

**Interfaces:**
- Consumes:
  - `document_ingestion.ingest_document(path: str) -> IngestedDocument` (свойство `.full_text: str`, `.warnings: list[str]`)
  - `tz_extraction.extract_tz_structured(doc: IngestedDocument) -> dict`
  - `tz_extraction.process_extraction(extraction: dict) -> ExtractedObjectSummary` (поля `object_name, region, area_sqm, cleaning_days, general_days, contract_months, staff_count, staff_count_source`)
  - `knowledge_base.get_combined_shift_salary(region, object_complexity, schedule_complexity) -> Optional[dict]` (ключ `base_shift_salary`)
  - `knowledge_base.ObjectParams(area_sqm, cleaning_days, general_days, staff_count, contract_months)`
  - `price_parser.load_general_opt_price_list(supplier=...) -> list`, `price_parser.load_retail_suppliers_manual_additions() -> list`
  - `smeta_builder.SmetaSheet`, `.build_contract_term_section()`, `.build_materials_table()` (возвращает `MaterialsTableResult` с полями `total_cell: str`, `unresolved: list[str]`)
- Produces:
  - `smeta_pipeline.PipelineError` (Exception) — для Task 3 (`app.py` ловит и показывает `st.error`)
  - `smeta_pipeline.SmetaPipelineResult` dataclass (поля: `xlsx_path, object_name, region, area_sqm, cleaning_days, contract_months, staff_count, staff_count_source, fot_month, materials_unresolved, warnings`) — для Task 3
  - `smeta_pipeline.generate_smeta(file_path, *, region="Москва", object_complexity="стандартный", schedule_complexity=..., vat_rate=0.20, output_dir="Outputs/tender-agent-smeta") -> SmetaPipelineResult` — главная точка входа для Task 3

- [ ] **Step 1: Написать `smeta_pipeline.py`**

Создать файл `smeta_pipeline.py` с полным содержимым:

```python
"""
smeta_pipeline.py — универсальный конвейер «ТЗ → оценочная смета» для
НОВОГО тендера. Оценка по нормативам (knowledge_base), а не по реальным
зафиксированным данным уже выигранных контрактов — см. CLAUDE.md и
*_smeta_export.py (там реальные цифры конкретных договоров, здесь —
оценка для объекта, по которому ещё нет контракта).
"""
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook

from document_ingestion import ingest_document
from knowledge_base import ObjectParams, get_combined_shift_salary
from price_parser import load_general_opt_price_list, load_retail_suppliers_manual_additions
from smeta_builder import (
    CURRENCY_FMT, FINAL_FONT, PERCENT_FMT, TOTAL_FONT,
    SmetaSheet, build_contract_term_section, build_materials_table,
)
from tz_extraction import extract_tz_structured, process_extraction


class PipelineError(Exception):
    """Понятная для UI ошибка на любом шаге конвейера."""


@dataclass
class SmetaPipelineResult:
    xlsx_path: str
    object_name: str
    region: str
    area_sqm: float
    cleaning_days: int
    contract_months: int
    staff_count: int
    staff_count_source: str
    fot_month: Optional[float]
    materials_unresolved: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _slugify(name: str) -> str:
    safe = re.sub(r"[^\w\-]+", "_", name, flags=re.UNICODE).strip("_")
    return safe[:60] or "smeta"


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
    warnings: list[str] = []

    # 1. Извлечение текста из документа
    doc = ingest_document(file_path)
    if len(doc.full_text.strip()) < 50:
        raise PipelineError(
            "Не удалось извлечь текст из документа (пусто или слишком мало "
            f"текста). Предупреждения: {'; '.join(doc.warnings) or 'нет'}."
        )

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

    # 4. ФОТ — оценка по нормативу
    salary = get_combined_shift_salary(region, object_complexity, schedule_complexity)
    fot_month: Optional[float] = None
    if salary is not None:
        fot_month = salary["base_shift_salary"] * summary.staff_count
    else:
        warnings.append(
            f"Норматив ФОТ не найден для региона «{region}», сложности "
            f"«{object_complexity}», графика «{schedule_complexity}». "
            "Введите ФОТ вручную в сгенерированном файле (жёлтая ячейка) "
            "или дополните knowledge_base/salary_rules.json."
        )

    # 5. Сборка Excel
    wb = Workbook()
    sheet = SmetaSheet.new(
        wb, "Смета",
        {c: w for c, w in zip("ABCDEFGHI", [46, 16, 16, 8, 34, 22, 12, 14, 14])},
    )
    sheet.title_block(
        summary.object_name,
        subtitle=(
            "Автоматически сгенерировано (оценка по нормативам, НЕ "
            "финальная цена — накладные расходы и маржу добавляет эксперт)."
        ),
        region=region,
    )

    term_refs = build_contract_term_section(
        sheet, summary.contract_months, area_total_sqm=summary.area_sqm,
    )

    sheet.section("2. Численность и ФОТ (оценка по нормативу)")
    sheet.value_row(
        f"Численность персонала ({summary.staff_count_source})",
        summary.staff_count, editable=True,
    )
    if fot_month is not None:
        r_fot = sheet.value_row(
            "ФОТ, ₽/мес (норматив × численность)", fot_month,
            editable=True, fmt=CURRENCY_FMT,
        )
    else:
        r_fot = sheet.value_row(
            "ФОТ, ₽/мес — НОРМАТИВ НЕ НАЙДЕН, ввести вручную", 0,
            editable=True, fmt=CURRENCY_FMT,
        )
        sheet.note(
            f"Норматив не найден для «{region}» / «{object_complexity}» / "
            f"«{schedule_complexity}». Введите значение вручную в жёлтую "
            "ячейку выше."
        )
    sheet.blank()

    params = ObjectParams(
        area_sqm=summary.area_sqm,
        cleaning_days=summary.cleaning_days,
        general_days=summary.general_days,
        staff_count=summary.staff_count,
        contract_months=summary.contract_months,
    )
    price_lists = (
        load_general_opt_price_list(supplier="ТК Сервис")
        + load_retail_suppliers_manual_additions()
    )
    mat = build_materials_table(sheet, params, price_lists, title="3. Расходные материалы")
    if mat.unresolved:
        warnings.append(
            f"Не найдено в прайсе (нужен ручной ввод цены): {', '.join(mat.unresolved)}"
        )

    sheet.section("4. ИТОГО — оценочная себестоимость")
    r_total_month = sheet.row
    sheet.ws.cell(row=r_total_month, column=1,
                  value="Себестоимость в месяц (ФОТ + материалы), без НДС").font = TOTAL_FONT
    sheet.ws.cell(row=r_total_month, column=2,
                  value=f"=B{r_fot}+{mat.total_cell}").number_format = CURRENCY_FMT
    sheet.ws.cell(row=r_total_month, column=2).font = TOTAL_FONT
    sheet.row += 1
    r_vat = sheet.value_row("Ставка НДС", vat_rate, editable=True, fmt=PERCENT_FMT)
    r_total_month_vat = sheet.row
    sheet.ws.cell(row=r_total_month_vat, column=1,
                  value="Себестоимость в месяц, с НДС").font = FINAL_FONT
    sheet.ws.cell(row=r_total_month_vat, column=2,
                  value=f"=B{r_total_month}*(1+B{r_vat})").number_format = CURRENCY_FMT
    sheet.ws.cell(row=r_total_month_vat, column=2).font = FINAL_FONT
    sheet.row += 1
    r_total_year = sheet.row
    sheet.ws.cell(row=r_total_year, column=1,
                  value="Себестоимость в год, с НДС").font = FINAL_FONT
    sheet.ws.cell(row=r_total_year, column=2,
                  value=f"=B{r_total_month_vat}*MIN(B{term_refs['months']},12)").number_format = CURRENCY_FMT
    sheet.ws.cell(row=r_total_year, column=2).font = FINAL_FONT
    sheet.row += 1
    sheet.note(
        "ОЦЕНОЧНАЯ себестоимость (ФОТ по нормативу + материалы по "
        "нормативу), НЕ финальная цена для тендера — накладные расходы и "
        "маржу эксперт добавляет отдельно (human-in-the-loop, см. CLAUDE.md)."
    )

    # 6. Сохранение
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    xlsx_path = out_dir / f"{_slugify(summary.object_name)}_{timestamp}.xlsx"
    wb.save(xlsx_path)

    return SmetaPipelineResult(
        xlsx_path=str(xlsx_path),
        object_name=summary.object_name,
        region=region,
        area_sqm=summary.area_sqm,
        cleaning_days=summary.cleaning_days,
        contract_months=summary.contract_months,
        staff_count=summary.staff_count,
        staff_count_source=summary.staff_count_source,
        fot_month=fot_month,
        materials_unresolved=mat.unresolved,
        warnings=warnings,
    )
```

- [ ] **Step 2: Исключить сгенерированные сметы из git**

Открыть `.gitignore` и добавить строку:
```
Outputs/
```

- [ ] **Step 3: Создать тестовый ТЗ-документ для проверки**

Run:
```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
source .venv/bin/activate
python3 -c "
from docx import Document

doc = Document()
lines = [
    'Техническое задание на оказание услуг по уборке помещений.',
    'Объект: Оказание услуг по уборке офисных помещений в 2027 году.',
    'Регион: Москва. Адрес: г. Москва, ул. Тестовая, д. 1.',
    'Код КПГЗ: 03.08.01.01.01.07.',
    '',
    'Приложение 1. Объёмы оказания услуг по периодам:',
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
Expected: печатает `saved`, файл `/tmp/sample_tz.docx` создан.

- [ ] **Step 4: Запустить конвейер end-to-end и проверить результат**

Run:
```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
source .venv/bin/activate
set -a && source .env && set +a
python3 -c "
from smeta_pipeline import generate_smeta

result = generate_smeta('/tmp/sample_tz.docx')
print('xlsx_path:', result.xlsx_path)
print('object_name:', result.object_name)
print('area_sqm:', result.area_sqm)
print('cleaning_days:', result.cleaning_days)
print('contract_months:', result.contract_months)
print('staff_count:', result.staff_count, '(', result.staff_count_source, ')')
print('fot_month:', result.fot_month)
print('materials_unresolved:', result.materials_unresolved)
print('warnings:', result.warnings)
"
```
Expected: без исключений; `area_sqm` около 100.0 (площадь из тестового документа), `cleaning_days` и `contract_months` — небольшие положительные числа, `xlsx_path` указывает на файл внутри `Outputs/tender-agent-smeta/`.

- [ ] **Step 5: Проверить, что в сгенерированном Excel нет ошибок формул**

Run:
```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
XLSX=$(ls -t Outputs/tender-agent-smeta/*.xlsx | head -1)
mkdir -p /tmp/recalc_check
libreoffice --headless --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir /tmp/recalc_check "$XLSX"
python3 -c "
import glob
import openpyxl

path = sorted(glob.glob('/tmp/recalc_check/*.xlsx'))[-1]
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active
errors = [c.value for row in ws.iter_rows() for c in row
          if isinstance(c.value, str) and c.value.startswith('#')]
print('Ошибки формул:', errors)
"
```
Expected: `Ошибки формул: []`

- [ ] **Step 6: Commit**

```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
git add smeta_pipeline.py .gitignore
git commit -m "Добавить smeta_pipeline.py — конвейер ТЗ -> оценочная смета для нового тендера"
```

---

## Task 3: `app.py` — Streamlit-интерфейс

**Files:**
- Create: `app.py`
- Modify: `.env` (добавить `APP_PASSWORD`)
- Modify: `README.md` (добавить раздел «Как запустить веб-интерфейс»)

**Interfaces:**
- Consumes: `smeta_pipeline.generate_smeta()`, `smeta_pipeline.PipelineError`, `smeta_pipeline.SmetaPipelineResult` (все из Task 2)

- [ ] **Step 1: Спросить у пользователя пароль для входа в интерфейс и добавить в `.env`**

Прежде чем писать код, спросить пользователя (через обычное сообщение в чате, не придумывать самостоятельно): «Какой пароль использовать для входа коллег в веб-интерфейс?». Получив ответ, добавить в `.env` строку:
```
APP_PASSWORD=<значение, которое назвал пользователь>
```
(`.env` уже в `.gitignore` — секрет в git не попадёт).

- [ ] **Step 2: Написать `app.py`**

Создать файл `app.py` с полным содержимым:

```python
import os
import tempfile
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

from smeta_pipeline import PipelineError, generate_smeta

load_dotenv()

st.set_page_config(page_title="Тендер-агент — смета", layout="centered")

APP_PASSWORD = os.environ.get("APP_PASSWORD")


def _check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    st.title("Тендер-агент")
    pwd = st.text_input("Пароль", type="password")
    if st.button("Войти"):
        if APP_PASSWORD and pwd == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Неверный пароль")
    return False


if not _check_password():
    st.stop()

st.title("Тендер-агент — расчёт сметы")
st.caption(
    "Загрузите ТЗ нового тендера — конвейер извлечёт данные через Claude API "
    "и соберёт оценочную смету (ФОТ и материалы — по нормативам)."
)

uploaded = st.file_uploader("Загрузите ТЗ (PDF/DOCX/XLSX)", type=["pdf", "docx", "xlsx"])

col1, col2 = st.columns(2)
with col1:
    region = st.selectbox("Регион", ["Москва"], index=0)
    object_complexity = st.selectbox("Сложность объекта", ["стандартный"], index=0)
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

st.caption(
    "Регион/сложность/график сейчас ограничены единственной заполненной "
    "комбинацией в базе нормативов (knowledge_base/salary_rules.json)."
)

if uploaded is not None and st.button("Рассчитать"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(uploaded.name).suffix) as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = tmp.name
    try:
        with st.spinner("Извлекаем данные из ТЗ и считаем смету (может занять до минуты)..."):
            result = generate_smeta(
                tmp_path,
                region=region,
                object_complexity=object_complexity,
                schedule_complexity=schedule_complexity,
                vat_rate=vat_rate_pct / 100,
            )
    except PipelineError as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"Неожиданная ошибка: {e}")
    else:
        st.success(f"Смета готова: {result.object_name}")
        m1, m2, m3 = st.columns(3)
        m1.metric("Площадь, м²", f"{result.area_sqm:,.1f}")
        m1.metric("Дней уборки", result.cleaning_days)
        m2.metric("Срок, мес", result.contract_months)
        m2.metric("Численность", result.staff_count)
        m3.metric(
            "ФОТ, ₽/мес",
            f"{result.fot_month:,.0f}" if result.fot_month is not None else "не найден",
        )
        for w in result.warnings:
            st.warning(w)
        with open(result.xlsx_path, "rb") as f:
            st.download_button(
                "Скачать смету (.xlsx)",
                data=f.read(),
                file_name=Path(result.xlsx_path).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    finally:
        os.unlink(tmp_path)
```

- [ ] **Step 3: Проверить, что сервер запускается и отвечает**

Run:
```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
source .venv/bin/activate
nohup streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --server.headless true > /tmp/streamlit.log 2>&1 &
sleep 5
curl -s -o /dev/null -w "%{http_code}\n" http://127.0.0.1:8501
```
Expected: `200`. Если не `200` — посмотреть `/tmp/streamlit.log` на ошибку импорта/синтаксиса.

- [ ] **Step 4: Остановить тестовый сервер**

Run:
```bash
pkill -f "streamlit run app.py"
```

- [ ] **Step 5: Добавить инструкцию запуска в README.md**

Добавить в конец `README.md` новый раздел:

```markdown
## Шаг 8 — веб-интерфейс (Этап 1)

```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
source .venv/bin/activate
streamlit run app.py --server.address 127.0.0.1 --server.port 8501
```

Открыть `http://127.0.0.1:8501` (локально на сервере или через SSH-туннель:
`ssh -L 8501:127.0.0.1:8501 <сервер>` с компьютера коллеги). Пароль — в
`.env` (`APP_PASSWORD`). Публичный доступ из интернета — отдельная задача
(см. `docs/superpowers/specs/2026-10-01-web-interface-design.md`).
```

- [ ] **Step 6: Ручная проверка полного цикла (загрузка → расчёт → скачивание)**

Запустить сервер как в Step 3 (без `&sleep&curl`, в интерактивном режиме), открыть `http://127.0.0.1:8501` в браузере (локально или через SSH-туннель), пройти: ввод пароля → загрузка `/tmp/sample_tz.docx` → «Рассчитать» → убедиться, что появились метрики и кнопка «Скачать смету (.xlsx)» → скачать файл и открыть, убедиться, что он открывается без ошибок Excel/LibreOffice. Остановить сервер (`Ctrl+C` или `pkill -f "streamlit run app.py"`).

- [ ] **Step 7: Commit**

```bash
cd /srv/cloudcli-users/ecoadmin/Projects/tender-agent
git add app.py README.md
git commit -m "Добавить app.py — Streamlit-интерфейс для расчёта сметы (Этап 1)"
```

---

## Self-Review Notes

- **Покрытие спецификации:** архитектура (2 файла), конвейер (Task 2, все 7 шагов из спеки — ingestion/extraction/process/ФОТ/материалы/сборка/сохранение), UI (Task 3, все элементы — пароль/форма/спиннер/ошибки/метрики/предупреждения/скачивание), поток ошибок (PipelineError на шагах 1 и 3, предупреждения на шагах 4 и 5 — не падаем), тестирование (ручные команды в каждом Task вместо pytest, как решено в спеке), открытые ограничения (видны через `warnings` в UI и в самой Excel-смете).
- **Отклонение от буквального порядка разделов в спеке/CLAUDE.md** («срок → цена → численность → материалы»): в существующем стандарте раздел «цена» — это ИЗВЕСТНАЯ входная величина (реальная ставка по уже заключённому договору, как в `mlrz_smeta_export_v2.py`). В нашем случае для НОВОГО тендера эта величина, наоборот, ВЫЧИСЛЯЕТСЯ из ФОТ+материалов, поэтому раздел «ИТОГО» логически должен идти последним (после того, как компоненты посчитаны) — реализовано по аналогии с разделом «5. Сверка себестоимости» в `mlrz_smeta_export_v2.py`, а не через неиспользуемый нигде в проекте `build_price_summary_section()`.
- **Проверка типов/сигнатур:** `SmetaPipelineResult`, `PipelineError`, `generate_smeta()` — используются в Task 3 ровно с теми именами и сигнатурами, что определены в Task 2.
