"""
Fallback-поиск позиции на сайте поставщика, если её нет ни в одном загруженном
прайсе. Точечный запрос по конкретному наименованию — не обход всего каталога.

В этой среде нет доступа в интернет, поэтому здесь — интерфейс и файловый кэш;
`fetch_from_supplier_site` в проде реализуется под конкретного поставщика
(запрос к внутреннему поиску сайта / парсинг карточки товара) и подключается
через `suppliers_config`.
"""
import json
from pathlib import Path
from typing import Callable, Optional

from models import PriceListItem, Source, TZItem

CACHE_PATH = Path("price_cache.json")
CACHE_TTL_DAYS = 7  # обновлять кэш раз в N дней


def _load_cache() -> dict:
    if CACHE_PATH.exists():
        return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    return {}


def _save_cache(cache: dict) -> None:
    CACHE_PATH.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def check_cache(tz_item: TZItem) -> Optional[PriceListItem]:
    cache = _load_cache()
    entry = cache.get(tz_item.raw_name)
    if not entry:
        return None
    # TTL-проверку по entry["fetched_at"] опущена здесь для краткости —
    # в проде: если старше CACHE_TTL_DAYS, вернуть None и перезапросить сайт.
    return PriceListItem(
        supplier=entry["supplier"],
        name=entry["name"],
        sku=entry.get("sku"),
        unit=entry.get("unit", "шт"),
        price=entry["price"],
        discount_pct=entry.get("discount_pct", 0.0),
        source=Source.CACHE,
    )


def save_to_cache(tz_item: TZItem, item: PriceListItem) -> None:
    cache = _load_cache()
    cache[tz_item.raw_name] = {
        "supplier": item.supplier,
        "name": item.name,
        "sku": item.sku,
        "unit": item.unit,
        "price": item.price,
        "discount_pct": item.discount_pct,
    }
    _save_cache(cache)


# Тип функции-адаптера под конкретного поставщика: (query: str) -> PriceListItem | None
SupplierSearchFn = Callable[[str], Optional[PriceListItem]]


def search_supplier_sites(
    tz_item: TZItem, supplier_adapters: dict[str, SupplierSearchFn]
) -> Optional[PriceListItem]:
    """Пробует точечный поиск по каждому настроенному поставщику по очереди."""
    for supplier_name, search_fn in supplier_adapters.items():
        try:
            result = search_fn(tz_item.raw_name)
        except Exception:
            # сайт недоступен / изменил вёрстку — не роняем весь пайплайн
            result = None
        if result:
            save_to_cache(tz_item, result)
            return result
    return None


def stub_supplier_adapter(query: str) -> Optional[PriceListItem]:
    """
    Пример адаптера-заглушки. В проде: HTTP-запрос к поиску на сайте поставщика,
    парсинг карточки товара (BeautifulSoup/httpx), возврат PriceListItem.
    Здесь всегда возвращает None — сети нет в этой среде.
    """
    return None
