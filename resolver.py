"""
Главная точка входа: resolve_item(tz_item) -> MatchResult

Порядок поиска:
  1. Fuzzy-match по всем загруженным прайсам (через предпостроенный индекс
     категорий — см. resolve_all).
  2. Если лучшая позиция уверенная (>= AUTO) — берём автоматически.
  3. Если найдена, но неуверенная (>= REVIEW, < AUTO) — берём, но помечаем
     на подтверждение эксперта.
  4. Если в прайсах вообще ничего похожего — идём в кэш, затем на сайты
     поставщиков (точечный запрос).
  5. Если и там пусто — NOT_FOUND, обязательно требует эксперта.
"""
from dataclasses import dataclass
from typing import Optional

from matcher import CONFIDENCE_THRESHOLD_AUTO, build_category_index, match_against_pricelists
from models import MatchResult, PriceListItem, Source, TZItem
from supplier_fallback import SupplierSearchFn, check_cache, search_supplier_sites
from unit_price import select_cheapest_equivalent


def resolve_item(
    tz_item: TZItem,
    price_lists: list[PriceListItem],
    supplier_adapters: dict[str, SupplierSearchFn] | None = None,
    category_index: Optional[dict] = None,
) -> MatchResult:
    candidates = match_against_pricelists(tz_item, price_lists, category_index=category_index)

    if candidates:
        best_item, best_score, equivalent, comparable = select_cheapest_equivalent(candidates)
        alternatives = [item for item, _ in equivalent if item is not best_item]
        return MatchResult(
            tz_item=tz_item,
            matched_item=best_item,
            confidence=best_score,
            source=Source.PRICE_LIST,
            requires_expert_review=best_score < CONFIDENCE_THRESHOLD_AUTO or not comparable,
            alternatives=alternatives,
            price_comparable=comparable,
        )

    # ничего в прайсах — сначала кэш (то, что уже когда-то нашли на сайте)
    cached = check_cache(tz_item)
    if cached:
        return MatchResult(
            tz_item=tz_item,
            matched_item=cached,
            confidence=1.0,  # кэш = ранее подтверждённая находка с сайта
            source=Source.CACHE,
            requires_expert_review=False,
        )

    # затем — точечный поиск на сайтах поставщиков
    if supplier_adapters:
        found = search_supplier_sites(tz_item, supplier_adapters)
        if found:
            return MatchResult(
                tz_item=tz_item,
                matched_item=found,
                confidence=1.0,
                source=Source.SUPPLIER_SITE,
                requires_expert_review=True,  # первая находка с сайта — всегда на проверку
            )

    return MatchResult(
        tz_item=tz_item,
        matched_item=None,
        confidence=0.0,
        source=Source.NOT_FOUND,
        requires_expert_review=True,
    )


def resolve_all(
    tz_items: list[TZItem],
    price_lists: list[PriceListItem],
    supplier_adapters: dict[str, SupplierSearchFn] | None = None,
) -> list[MatchResult]:
    # Индекс категорий строится один раз на весь прайс и переиспользуется для
    # каждой позиции ТЗ — без этого на больших прайсах (тысячи строк) каждая
    # позиция ТЗ заново пересчитывала бы score по всему прайсу целиком.
    category_index = build_category_index(price_lists)
    return [
        resolve_item(item, price_lists, supplier_adapters, category_index=category_index)
        for item in tz_items
    ]
