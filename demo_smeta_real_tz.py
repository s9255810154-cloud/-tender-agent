from datetime import date

from models import TZItem
from price_parser import load_almin_price_list, load_general_opt_price_list
from resolver import resolve_all
from smeta_export import LaborCostInput, SmetaMeta, export_smeta_to_xlsx
from supplier_fallback import stub_supplier_adapter

# --- "Приложение 6. Перечень расходных материалов" из реального ТЗ ---
# ВАЖНО: в самом ТЗ (Приложение 6) НЕ указаны числовые количества — колонка
# "Количество" везде содержит только текст "количество расходных материалов,
# необходимое для бесперебойного и качественного оказания услуг...".
# Это значит, что qty здесь принципиально неизвестен из документа — его
# должен оценить эксперт (по нормам расхода на кв.м/чел., которых пока нет
# в проекте — см. TODO в tender-agent-architecture.md), а не агент.
tz_items = [
    TZItem(raw_name="Пакеты для мусора. Объём 120 л. Размер 70x110 см. Материал ПВД. Плотность 50 мкм", unit="шт"),
    TZItem(raw_name="Пакеты для мусора. Объём 30 л. Размер 48x57 см. Материал ПНД. Плотность 7 мкм. Цвет чёрный", unit="шт"),
    TZItem(raw_name="Моющее средство для полов и стен. Антимикробный эффект. Без хлора. Без отдушки", unit="шт"),
    TZItem(raw_name="Чистящее средство для унитаза. С дезинфекцией", unit="шт"),
    TZItem(raw_name="Чистящее средство для сантехники. Без хлора", unit="шт"),
    TZItem(raw_name="Дезинфицирующее средство для твёрдых поверхностей, инвентаря, приборов", unit="шт"),
    TZItem(raw_name="Чистящее средство для интерьера", unit="шт"),
    TZItem(raw_name="Чистящее средство для оргтехники", unit="шт"),
    TZItem(raw_name="Средство для мойки стёкол и зеркал", unit="шт"),
    TZItem(raw_name="Салфетки универсальные. Материал микрофибра", unit="шт"),
    TZItem(raw_name="Комплект для уборки (швабра, ведро) для влажной уборки", unit="шт"),
]

price_lists = (
    load_almin_price_list(discount_pct=5.0)
    + load_general_opt_price_list(supplier="ТК Сервис", discount_pct=0.0)
)

results = resolve_all(tz_items, price_lists, supplier_adapters={"almin": stub_supplier_adapter})

# --- Трудозатраты ---
# ТЗ прямо указывает численность (п. после Приложения 3):
#   "Количество работников, выполняющих оказание услуг по уборке помещений:
#    не менее 2 (двух) человек" — берём как given, не оцениваем по площади/
#    графику, оценка не нужна раз число указано явно в самом документе.
# ВАЖНО: оклад по региону/сложности объекта посчитать не могу — таблицы
# ставок (region × object_complexity × schedule_complexity → base_salary)
# в проекте ещё нет (открытый TODO в архитектуре). Ставлю 0 и жёлтую
# заливку — это осознанный пропуск, а не забытая деталь.
labor = LaborCostInput(
    monthly_salary=0,  # ЗАПОЛНИТЬ: нет таблицы ставок по региону/сложности
    staff_count_by_area=0,     # не считали — численность дана в ТЗ явно
    staff_count_by_hours=0,    # не считали — численность дана в ТЗ явно
    staff_count_recommended=2,  # "не менее 2 человек" — прямая цитата из ТЗ
    contract_duration_months=36,  # "в 2026-2029 годах" ≈ 3 года — ПРОВЕРИТЬ по датам контракта
)

meta = SmetaMeta(
    object_name="Оказание услуг по уборке помещений в 2026-2029 годах — Москва, Большой Головин переулок, 15",
    region="Москва",
    tender_url="Техническое_задание_17204828-1.pdf",
    prepared_date=date.today(),
)

export_smeta_to_xlsx(results, labor, meta, "smeta_golovin.xlsx")
print("Смета сохранена: smeta_golovin.xlsx")
