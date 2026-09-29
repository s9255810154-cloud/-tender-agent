"""
Единые модели данных для сопоставления позиций ТЗ с прайсами поставщиков.
"""
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Optional


class Source(str, Enum):
    PRICE_LIST = "price_list"      # найдено в загруженном прайсе (Excel/CSV/PDF)
    SUPPLIER_SITE = "supplier_site"  # найдено точечным запросом на сайте поставщика
    CACHE = "cache"                  # найдено в локальном кэше (ранее спарсено с сайта)
    NOT_FOUND = "not_found"          # не найдено нигде — уходит эксперту


@dataclass
class PriceListItem:
    """Одна позиция из прайса поставщика (из файла или с сайта)."""
    supplier: str
    name: str                  # полное наименование, как в прайсе
    sku: Optional[str] = None  # артикул, если есть
    unit: str = "шт"           # единица измерения
    price: float = 0.0
    pack_size: Optional[str] = None   # например "1/12" (штук в упаковке/коробе)
    discount_pct: float = 0.0         # скидка поставщика, %
    source: Source = Source.PRICE_LIST
    fetched_at: datetime = field(default_factory=datetime.utcnow)

    @property
    def final_price(self) -> float:
        return round(self.price * (1 - self.discount_pct / 100), 2)


@dataclass
class TZItem:
    """Одна позиция расходников, извлечённая из ТЗ."""
    raw_name: str               # формулировка как в ТЗ, например:
                                 # "Туалетная бумага от 16 м. 2х ГОСТ Р 52354-2025"
    qty: Optional[float] = None
    unit: Optional[str] = None
    notes: Optional[str] = None


@dataclass
class MatchResult:
    tz_item: TZItem
    matched_item: Optional[PriceListItem]
    confidence: float          # 0.0–1.0
    source: Source
    requires_expert_review: bool = False
    alternatives: list[PriceListItem] = field(default_factory=list)
    # другие поставщики, чья позиция признана тем же товаром (score близок к
    # лучшему), но final_price выше выбранной — для прозрачности перед экспертом
    price_comparable: bool = True
    # False = у равнозначных кандидатов не удалось привести цену к одной
    # единице измерения (шт/л/кг/м) — сравнение прошло по цене за упаковку
    # "как есть" и может быть некорректным при разных размерах фасовки

    def to_dict(self) -> dict:
        return {
            "tz_name": self.tz_item.raw_name,
            "matched_name": self.matched_item.name if self.matched_item else None,
            "supplier": self.matched_item.supplier if self.matched_item else None,
            "final_price": self.matched_item.final_price if self.matched_item else None,
            "confidence": round(self.confidence, 3),
            "source": self.source.value,
            "requires_expert_review": self.requires_expert_review,
            "price_comparable": self.price_comparable,
            "alternatives": [
                {"supplier": alt.supplier, "name": alt.name, "final_price": alt.final_price}
                for alt in self.alternatives
            ],
        }
