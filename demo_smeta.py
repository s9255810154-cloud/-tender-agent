from datetime import date

from models import TZItem
from price_parser import load_almin_price_list, load_general_opt_price_list
from resolver import resolve_all
from smeta_export import LaborCostInput, SmetaMeta, export_smeta_to_xlsx
from supplier_fallback import stub_supplier_adapter

tz_items = [
    TZItem(raw_name="Туалетная бумага от 16 м. 2х ГОСТ Р 52354-2025", qty=500, unit="рул"),
    TZItem(raw_name="Мыло жидкое туалетное 5 л с антибактериальным эффектом", qty=20, unit="шт"),
    TZItem(raw_name="Пакеты для мусора 120 л, плотность не менее 20 мкм", qty=100, unit="шт"),
    TZItem(raw_name="Полотенце бумажное листовое V-сложения 2-слойное белое", qty=300, unit="пач"),
    TZItem(raw_name="Перчатки латексные хозяйственные размер XL", qty=50, unit="пар"),
    TZItem(raw_name="Средство для мытья полов, объём 5 л, нейтральный pH", qty=10, unit="шт"),
    TZItem(raw_name="Дозирующая насадка для жидкого мыла, пенообразующая", qty=15, unit="шт"),
]

price_lists = (
    load_almin_price_list(discount_pct=5.0)
    + load_general_opt_price_list(supplier="ТК Сервис", discount_pct=0.0)
)

results = resolve_all(tz_items, price_lists, supplier_adapters={"almin": stub_supplier_adapter})

# Пример трудозатрат — по логике из tender-agent-architecture.md:
# get_monthly_salary(region, object_complexity, schedule_complexity, ...) уже
# даёт готовый оклад; estimate_staff_count() даёт две оценки. Здесь — заглушка
# с правдоподобными числами для проверки экспорта.
labor = LaborCostInput(
    monthly_salary=52000,
    staff_count_by_area=4,
    staff_count_by_hours=5,
    staff_count_recommended=5,  # max(4, 5) — подлежит подтверждению экспертом
    contract_duration_months=12,
)

meta = SmetaMeta(
    object_name="Комплексная уборка, Бутырская/Новый Арбат",
    region="Москва",
    tender_url="zakupki.gov.ru/... (пример)",
    prepared_date=date.today(),
)

export_smeta_to_xlsx(results, labor, meta, "smeta_example.xlsx")
print("Смета сохранена: smeta_example.xlsx")
