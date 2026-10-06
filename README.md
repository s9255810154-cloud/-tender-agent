# Тендер-агент для клининговой компании — инструкция по развёртыванию

Проект собирался итеративно в чате claude.ai (без git, без сервера) — тестировался
на реальных документах (ТЗ, договоры, сметы, счета), которые присылал пользователь.
Эта инструкция — перенос уже проверенного кода на сервер в Claude Code
(`/srv/cloudcli-users/ecoadmin/...`, см. `tender-agent-architecture.md`).

## Что уже проверено, а что нет — важно перед переносом

- **Работает и проверено на реальных данных**: извлечение из ТЗ (area-trap, ГЭСН/затратный
  метод, гибридные документы), расчёт ФОТ (0,11% расхождение с реальным табелем),
  расчёт материалов (сверен по ₽/м² с 2 независимыми реальными объектами), сопоставление
  с прайс-листами (resolver.py), экспорт в Excel (smeta_builder.py, 0 ошибок recalc).
- **НИ РАЗУ не тестировалось вживую**: `extract_tz_structured()` в `tz_extraction.py` —
  реальный вызов Anthropic API. В чат-песочнице не было прямого доступа в интернет к
  api.anthropic.com. Это первое, что нужно проверить на сервере.
- **Не начато**: Step 1-2 архитектуры (сбор тендеров с ЕИС/коммерческих площадок) —
  `tender_models.py` есть (модели данных, фильтрация, дедуп), но не подключён ни к
  одному живому источнику. Разобрались, как минимум, с двумя каналами (см.
  `tender-agent-architecture.md` и историю чата): ЕИС SOAP-интеграция (бесплатно, но
  кривая документация) и Контур.Закупки (платно, без публичной документации).
- **Открытый пробел**: территория/озеленение (лужайки, летнее благоустройство) —
  модель материалов один раз занизила смету на ~31% (объект Звезда/Главкино).

## Шаг 1 — распаковать и инициализировать git

```bash
cd /srv/cloudcli-users/ecoadmin/           # или куда решите класть проект
unzip tender_matcher_export.zip
mv tender_matcher_export tender-agent      # или другое имя
cd tender-agent
git init
git add .
git commit -m "Импорт наработок из чат-сессии claude.ai"
```

## Шаг 2 — системные зависимости (не через pip)

```bash
# Debian/Ubuntu
sudo apt-get update
sudo apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-rus libreoffice
```

- `poppler-utils` (`pdftotext -layout`) — извлечение текста из PDF
- `tesseract-ocr` + `tesseract-ocr-rus` — OCR сканов (когда PDF — картинка, не текст)
- `libreoffice` (headless) — конвертация старого `.doc` (Word 97-2003) в `.docx`,
  и починка битых `.xlsx` (пересохранением)

Проверка:
```bash
pdftotext -v
tesseract --list-langs | grep rus
libreoffice --version
```

## Шаг 3 — Python-окружение

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

`requirements.txt` сгенерирован сканированием реальных импортов во всех `.py` файлах
проекта (не вручную составлен — так надёжнее, ничего не забыто).

## Шаг 4 — переменные окружения

```bash
export ANTHROPIC_API_KEY="sk-..."
```

Нужен для `tz_extraction.extract_tz_structured()` — единственного места, где код
реально дёргает Claude API для извлечения структурированных данных из ТЗ.
**Первым делом после установки — проверить именно этот вызов** на одном из реальных
ТЗ (они есть в `data/` или можно взять свежий с zakupki.gov.ru) — это единственная
непроверенная вживую часть пайплайна.

## Шаг 5 — быстрая проверка, что всё живо

```bash
source .venv/bin/activate
python3 -c "
from price_parser import load_general_opt_price_list, load_retail_suppliers_manual_additions
from knowledge_base import load_consumable_norms, load_salary_rules
print('Прайс-лист:', len(load_general_opt_price_list(supplier='ТК Сервис')), 'позиций')
print('Розничные добавки:', len(load_retail_suppliers_manual_additions()), 'позиций')
print('Нормы расходников:', len(load_consumable_norms()), 'штук')
print('Правила окладов:', len(load_salary_rules()), 'штук')
"
```

Должно отработать без ошибок и напечатать ненулевые числа. Если что-то падает —
скорее всего забыт шаг 2 или 3.

Дальше — сгенерировать тестовую смету и прогнать проверку формул:
```bash
python3 mlrz_smeta_export_v2.py
python3 -c "
import subprocess
# recalc.py — скрипт из xlsx-скилла claude.ai, на сервере его нет по умолчанию;
# либо скопировать его логику (LibreOffice headless recalculation), либо просто
# открыть smeta_mlrz_v2.xlsx в Excel/LibreOffice и проверить глазами на #REF!/#DIV0!
"
```
(`recalc.py` — вспомогательный скрипт из среды claude.ai для пересчёта формул и
поиска ошибок в xlsx через LibreOffice headless; на сервере такого скрипта нет —
нужно либо перенести его логику самостоятельно (headless LibreOffice recalculation
+ чтение cached values), либо просто открывать сгенерированные файлы в Excel/LibreOffice
и проверять глазами.)

## Шаг 6 — структура проекта (что где лежит)

```
tender-agent/
├── knowledge_base/          # JSON: нормы, расценки, реальные бенчмарки (24 файла)
├── data/                    # прайс-листы поставщиков (general_opt_price.xlsx — 4093 позиции,
│                             #   tk_service_manual_additions.json, retail_suppliers_manual_additions.json)
├── smeta_builder.py          # ОБЩИЙ модуль для новых смет (см. ниже)
├── *_smeta_export.py         # скрипты по конкретным объектам (частично мигрированы на smeta_builder.py)
├── knowledge_base.py          # доступ к JSON-базе + формулы норм
├── tz_extraction.py            # извлечение структуры из ТЗ (детекция типа документа, GCD-восстановление площади)
├── document_ingestion.py        # PDF/DOCX/XLSX/скан → единый текст
├── matcher.py, resolver.py, price_parser.py, unit_price.py  # сопоставление с прайсами
└── tender_models.py             # модели тендера — ПОКА не подключено к живому источнику
```

Для новых смет предпочтительно использовать `smeta_builder.py`
(`SmetaSheet`, `build_materials_table()`, ...) вместо копирования старого скрипта —
он даёт настоящие Excel-формулы для количества материалов и меньше дублирования кода.
Пример миграции — сравните `mlrz_smeta_export.py` (старый) и `mlrz_smeta_export_v2.py` (новый).

## Шаг 7 — дальше решить: CLI-скрипты или сервис

Архитектурный документ (`tender-agent-architecture.md`) предполагает FastAPI +
APScheduler/Celery + PostgreSQL + Telegram + Google Sheets — это следующий уровень
после проверки, что пайплайн работает end-to-end на сервере как есть (шаги 1-5).
Не обязательно делать всё сразу — можно погонять как есть скриптами, пока не
понадобится автоматический график/уведомления.

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

Поддерживаются только ТЗ на уборку помещений (не территории/благоустройство) —
см. ограничение в `docs/superpowers/specs/2026-10-01-web-interface-design.md`.
