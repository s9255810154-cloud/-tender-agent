"""
Смета по реальному ТЗ (Большой Головин переулок, 15) — теперь коэффициенты
расхода берутся из knowledge_base/, а не зашиты в скрипт. Площадь и дни
уборки по-прежнему выводятся математически из самого ТЗ (см. предыдущую
версию скрипта и чат — здесь это не повторяется, просто задаётся как факт).
"""
from datetime import date

from knowledge_base import ObjectParams, estimate_consumable_quantity, get_salary, get_staff_estimation_norms
from models import TZItem
from price_parser import load_almin_price_list, load_general_opt_price_list
from resolver import resolve_all
from smeta_export import LaborCostInput, SmetaMeta, export_smeta_to_xlsx
from supplier_fallback import stub_supplier_adapter

# --- Параметры объекта — выведены математически из Приложения 1 ТЗ (см. чат) ---
params = ObjectParams(
    area_sqm=1002.6,
    cleaning_days=742,
    general_days=148,
    staff_count=2,          # "не менее 2 человек" — прямая цитата из ТЗ
    contract_months=36,     # 01.11.2026 — 31.10.2029
)

# --- Позиции ТЗ ↔ нормы из базы знаний ---
# Соответствие "формулировка в ТЗ" -> "id нормы в consumable_norms.json"
# в реальном пайплайне будет расставлять шаг извлечения ТЗ (LLM), сейчас —
# вручную, как и раньше расставлялись формулы.
CONSUMABLE_MAPPING = [
    ("Пакеты для мусора. Объём 120 л. Размер 70x110 см. Материал ПВД. Плотность 50 мкм", "мусорный_мешок_крупный"),
    ("Пакеты для мусора. Объём 30 л. Размер 48x57 см. Материал ПНД. Плотность 7 мкм. Цвет чёрный", "мусорный_мешок_мелкий"),
    ("Моющее средство для полов и стен. Антимикробный эффект. Без хлора. Без отдушки", "моющее_средство_полы_стены"),
    ("Чистящее средство для унитаза. С дезинфекцией", "чистящее_унитаз"),
    ("Чистящее средство для сантехники. Без хлора", "чистящее_сантехника"),
    ("Дезинфицирующее средство для твёрдых поверхностей, инвентаря, приборов", "дезинфицирующее_средство"),
    ("Чистящее средство для интерьера", "чистящее_интерьер"),
    ("Чистящее средство для оргтехники", "чистящее_оргтехника"),
    ("Средство для мойки стёкол и зеркал", "средство_стекла_зеркала"),
    ("Салфетки универсальные. Материал микрофибра", "салфетки_универсальные"),
    ("Комплект для уборки (швабра, ведро) для влажной уборки", "комплект_уборки"),
]

tz_items = []
print("Расчёт количества по базе знаний:")
for raw_name, norm_id in CONSUMABLE_MAPPING:
    qty, unit, formula_desc = estimate_consumable_quantity(norm_id, params)
    print(f"  {norm_id:35} {qty:>10} {unit:4}  ({formula_desc})")
    tz_items.append(TZItem(raw_name=raw_name, qty=qty, unit=unit))

price_lists = (
    load_almin_price_list(discount_pct=5.0)
    + load_general_opt_price_list(supplier="ТК Сервис", discount_pct=0.0)
)

results = resolve_all(tz_items, price_lists, supplier_adapters={"almin": stub_supplier_adapter})

# --- Трудозатраты: оклад через базу знаний, если заполнена (пока — TODO) ---
salary_rule = get_salary(region="Москва", object_complexity="стандартный", schedule_complexity="стандартный (без ночных/выходных смен)")
monthly_salary = salary_rule["base_salary"] if salary_rule else 0

staff_norms = get_staff_estimation_norms()  # None, пока не заполнено — используем численность из самого ТЗ

labor = LaborCostInput(
    monthly_salary=monthly_salary,
    staff_count_by_area=0,   # не считали — норм sqm_per_person ещё нет (staff_norms.json)
    staff_count_by_hours=0,
    staff_count_recommended=params.staff_count,
    contract_duration_months=params.contract_months,
)

meta = SmetaMeta(
    object_name="Оказание услуг по уборке помещений в 2026-2029 годах — Москва, Большой Головин переулок, 15",
    region="Москва",
    tender_url="Техническое_задание_17204828-1.pdf",
    prepared_date=date.today(),
)

export_smeta_to_xlsx(results, labor, meta, "smeta_golovin_v3_kb.xlsx")
print("\nСмета сохранена: smeta_golovin_v3_kb.xlsx")
if salary_rule is None:
    print("Оклад: TODO — записи для (Москва, стандартный, стандартный) нет в salary_rules.json")
