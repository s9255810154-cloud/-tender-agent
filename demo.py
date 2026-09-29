from models import TZItem
from price_parser import load_almin_price_list, load_general_opt_price_list
from resolver import resolve_all
from supplier_fallback import stub_supplier_adapter
from unit_price import normalized_unit_price

# Позиции "как в ТЗ" — формулировки нарочно ближе к ГОСТ-стилю госзакупок,
# не к маркетинговым названиям поставщика.
tz_items = [
    TZItem(raw_name="Туалетная бумага от 16 м. 2х ГОСТ Р 52354-2025", qty=500, unit="рул"),
    TZItem(raw_name="Мыло жидкое туалетное 5 л с антибактериальным эффектом", qty=20, unit="шт"),
    TZItem(raw_name="Пакеты для мусора 120 л, плотность не менее 20 мкм", qty=100, unit="шт"),
    TZItem(raw_name="Полотенце бумажное листовое V-сложения 2-слойное белое", qty=300, unit="пач"),
    TZItem(raw_name="Перчатки латексные хозяйственные размер XL", qty=50, unit="пар"),
    TZItem(raw_name="Средство для мытья полов, объём 5 л, нейтральный pH", qty=10, unit="шт"),
    TZItem(raw_name="Дозирующая насадка для жидкого мыла, пенообразующая", qty=15, unit="шт"),  # заведомо нет в прайсе
]

price_lists = (
    load_almin_price_list(discount_pct=5.0)  # пример: скидка 5% по договору
    + load_general_opt_price_list(supplier="ТК Сервис", discount_pct=0.0)
)

results = resolve_all(
    tz_items,
    price_lists,
    supplier_adapters={"almin": stub_supplier_adapter},
)

print(f"{'ТЗ →':55} {'Поставщик':22} {'Найдено':38} {'цена':>8} {'₽/ед':>8} expert?")
print("-" * 155)
for r in results:
    tz = r.tz_item.raw_name[:53]
    supplier = (r.matched_item.supplier[:20] if r.matched_item else "—")
    matched = (r.matched_item.name[:36] if r.matched_item else "—")
    price = f"{r.matched_item.final_price:.2f}" if r.matched_item else "—"
    if r.matched_item:
        unit_price, unit_label = normalized_unit_price(r.matched_item)
        per_unit = f"{unit_price:.2f}{unit_label}"
    else:
        per_unit = "—"
    note = "" if r.price_comparable else "  ⚠ разные ед.изм — сравнение ненадёжно"
    print(f"{tz:55} {supplier:22} {matched:38} {price:>8} {per_unit:>8} {r.requires_expert_review}{note}")
    for alt in r.alternatives:
        alt_unit_price, alt_unit_label = normalized_unit_price(alt)
        print(f"   ↳ дороже, не выбрана: {alt.supplier:20} {alt.name[:40]:40} "
              f"{alt.final_price:>8.2f} ({alt_unit_price:.2f}{alt_unit_label})")

print("\nJSON (для передачи в шаг расчёта сметы):")
import json
print(json.dumps([r.to_dict() for r in results], ensure_ascii=False, indent=2))
