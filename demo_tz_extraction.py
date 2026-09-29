"""
Демонстрация шага 4 БЕЗ живого вызова API (в этой среде нет сети) —
подставляем extraction-словарь вручную, ровно в том виде, который должен
вернуть Claude по схеме EXTRACTION_TOOL для реального ТЗ (Большой Головин).
Это тот же самый текст документа, который мы разбирали в чате — значения
периодов и расходников взяты из него, а не придуманы для теста.

Цель: проверить, что compute_real_area/compute_cleaning_days и весь
пост-обработка НЕ зависят от Claude и дают ТЕ ЖЕ числа (1002.6 м², 742 дня,
36 месяцев), что мы посчитали вручную в чате.
"""
from tz_extraction import process_extraction

# Периоды — как они реально выглядят в Приложении 1 ТЗ (объект 1: ноя-дек
# 2026; объект 2: весь 2027; объект 3: весь 2028; объект 4: янв-окт 2029).
mock_extraction = {
    "object_name": "Оказание услуг по уборке помещений в 2026-2029 годах",
    "region": "Москва",
    "address": "г. Москва, Большой Головин переулок, 15",
    "kpgz_code": "03.08.01.01.01.07",
    "contract_periods": [
        {"period_start": "2026-11-01", "period_end": "2026-11-30", "service_volume_sqm": 20052},
        {"period_start": "2026-12-01", "period_end": "2026-12-31", "service_volume_sqm": 22057.2},
        {"period_start": "2027-01-01", "period_end": "2027-01-31", "service_volume_sqm": 15039},
        {"period_start": "2027-02-01", "period_end": "2027-02-28", "service_volume_sqm": 19049.4},
        {"period_start": "2027-03-01", "period_end": "2027-03-31", "service_volume_sqm": 22057.2},
        {"period_start": "2027-04-01", "period_end": "2027-04-30", "service_volume_sqm": 22057.2},
        {"period_start": "2027-05-01", "period_end": "2027-05-31", "service_volume_sqm": 19049.4},
        {"period_start": "2027-06-01", "period_end": "2027-06-30", "service_volume_sqm": 21054.6},
        {"period_start": "2027-07-01", "period_end": "2027-07-31", "service_volume_sqm": 22057.2},
        {"period_start": "2027-08-01", "period_end": "2027-08-31", "service_volume_sqm": 22057.2},
        {"period_start": "2027-09-01", "period_end": "2027-09-30", "service_volume_sqm": 22057.2},
        {"period_start": "2027-10-01", "period_end": "2027-10-31", "service_volume_sqm": 21054.6},
        {"period_start": "2027-11-01", "period_end": "2027-11-30", "service_volume_sqm": 20052},
        {"period_start": "2027-12-01", "period_end": "2027-12-31", "service_volume_sqm": 22057.2},
        {"period_start": "2028-01-01", "period_end": "2028-01-31", "service_volume_sqm": 16041.6},
        {"period_start": "2028-02-01", "period_end": "2028-02-29", "service_volume_sqm": 20052},
        {"period_start": "2028-03-01", "period_end": "2028-03-31", "service_volume_sqm": 22057.2},
        {"period_start": "2028-04-01", "period_end": "2028-04-30", "service_volume_sqm": 20052},
        {"period_start": "2028-05-01", "period_end": "2028-05-31", "service_volume_sqm": 20052},
        {"period_start": "2028-06-01", "period_end": "2028-06-30", "service_volume_sqm": 21054.6},
        {"period_start": "2028-07-01", "period_end": "2028-07-31", "service_volume_sqm": 21054.6},
        {"period_start": "2028-08-01", "period_end": "2028-08-31", "service_volume_sqm": 23059.8},
        {"period_start": "2028-09-01", "period_end": "2028-09-30", "service_volume_sqm": 21054.6},
        {"period_start": "2028-10-01", "period_end": "2028-10-31", "service_volume_sqm": 22057.2},
        {"period_start": "2028-11-01", "period_end": "2028-11-30", "service_volume_sqm": 21054.6},
        {"period_start": "2028-12-01", "period_end": "2028-12-31", "service_volume_sqm": 21054.6},
        {"period_start": "2029-01-01", "period_end": "2029-01-31", "service_volume_sqm": 17044.2},
        {"period_start": "2029-02-01", "period_end": "2029-02-28", "service_volume_sqm": 19049.4},
        {"period_start": "2029-03-01", "period_end": "2029-03-31", "service_volume_sqm": 20052},
        {"period_start": "2029-04-01", "period_end": "2029-04-30", "service_volume_sqm": 21054.6},
        {"period_start": "2029-05-01", "period_end": "2029-05-31", "service_volume_sqm": 20052},
        {"period_start": "2029-06-01", "period_end": "2029-06-30", "service_volume_sqm": 20052},
        {"period_start": "2029-07-01", "period_end": "2029-07-31", "service_volume_sqm": 22057.2},
        {"period_start": "2029-08-01", "period_end": "2029-08-31", "service_volume_sqm": 23059.8},
        {"period_start": "2029-09-01", "period_end": "2029-09-30", "service_volume_sqm": 20052},
        {"period_start": "2029-10-01", "period_end": "2029-10-31", "service_volume_sqm": 23059.8},
    ],
    "schedule": {
        "general_cleaning_per_week": 1,
        "daily_main_hours_per_day": 6.5,
        "daily_supporting_hours_per_day": 1.5,
        "workdays_per_week": 5,
    },
    "explicit_staff_count": {"value": 2, "condition": "не менее"},
    "consumables": [
        {"raw_name": "Пакеты для мусора", "characteristics": "Объём 120 л, Размер 70x110 см, ПВД, 50 мкм", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Пакеты для мусора", "characteristics": "Объём 30 л, Размер 48x57 см, ПНД, 7 мкм, чёрный", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Моющее средство для полов и стен", "characteristics": "Антимикробный эффект, без хлора, без отдушки", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Чистящее средство для унитаза", "characteristics": "С дезинфекцией", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Чистящее средство для сантехники", "characteristics": "Без хлора", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Дезинфицирующее средство", "characteristics": "Для твёрдых поверхностей, инвентаря, приборов", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Чистящее средство для интерьера", "characteristics": "", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Чистящее средство для оргтехники", "characteristics": "", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Средство для мойки стёкол и зеркал", "characteristics": "", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Салфетки универсальные", "characteristics": "Материал микрофибра", "unit": "шт", "explicit_qty": None},
        {"raw_name": "Комплект для уборки", "characteristics": "Швабра, ведро, для влажной уборки", "unit": "шт", "explicit_qty": None},
    ],
}

summary = process_extraction(mock_extraction)

print(f"Объект: {summary.object_name}")
print(f"Регион: {summary.region}, адрес: {summary.address}")
print(f"Площадь (вычислена НОДом объёмов): {summary.area_sqm} кв.м  [ожидалось: 1002.6]")
print(f"Дней уборки за контракт: {summary.cleaning_days}  [ожидалось: 742]")
print(f"Из них генеральных: {summary.general_days}  [ожидалось: ~148]")
print(f"Срок контракта, мес: {summary.contract_months}  [ожидалось: 36]")
print(f"Явная численность: {summary.explicit_staff_count}  [ожидалось: 2]")
print(f"Позиций расходников: {len(summary.tz_items)}  [ожидалось: 11]")
print()
print("Первые 3 позиции расходников:")
for item in summary.tz_items[:3]:
    print(f"  {item.raw_name} (qty={item.qty}, unit={item.unit})")
