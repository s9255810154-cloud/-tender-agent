"""
Источник-независимая модель закупки (тендера).

Смысл: шаг 1 (сбор) будет состоять из НЕСКОЛЬКИХ адаптеров под разные
источники (официальный SOAP-сервис ЕИС, платный агрегатор, партнёрский API
площадки, в крайнем случае — скрапинг). У каждого источника свой формат
ответа. Чтобы шаги 2+ (фильтрация, дедупликация, дальше по пайплайну —
document ingestion и т.д.) не переписывались под каждый новый источник,
каждый адаптер обязан уметь одно: превратить то, что он получил, в
TenderRaw -> Tender. Дальше всё работает с Tender и не знает про источники.

Похожий принцип уже применён в моделях прайсов (models.py: PriceListItem
не знает, был ли он распарсен из PDF Альмина или из xlsx на 4093 строки).
"""
from dataclasses import dataclass, field
from datetime import date, datetime
from enum import Enum
from typing import Optional


class TenderSource(str, Enum):
    EIS_OFFICIAL = "eis_official"   # официальный SOAP-сервис ЕИС (getDocsIP/getDocsLE2)
    AGGREGATOR = "aggregator"        # платный агрегатор (DaMIA, Мультитендер, NewDB...)
    PLATFORM = "platform"            # партнёрский API коммерческой площадки (B2B-Center и т.п.)
    SCRAPE = "scrape"                # скрапинг страницы поиска — последний вариант


class ProcurementLaw(str, Enum):
    FZ_44 = "44-ФЗ"
    FZ_223 = "223-ФЗ"
    COMMERCIAL = "коммерческая"       # площадки вне 44/223-ФЗ


class ProcurementMethod(str, Enum):
    AUCTION = "аукцион"
    COMPETITION = "конкурс"
    QUOTATION_REQUEST = "запрос котировок"
    SINGLE_SUPPLIER = "единственный поставщик"
    OTHER = "прочее"


class TenderStatus(str, Enum):
    ANNOUNCED = "объявлена"
    SUBMISSION_OPEN = "приём заявок"
    SUBMISSION_CLOSED = "приём заявок завершён"
    SUMMARIZING = "подведение итогов"
    COMPLETED = "завершена"
    CANCELLED = "отменена"
    UNKNOWN = "неизвестно"


@dataclass
class TenderRaw:
    """
    То, что вернул конкретный источник, почти без изменений (кроме упаковки
    в общий конверт). source_payload хранит исходный ответ целиком — для
    отладки и на случай, если normalize_tender() пропустит поле, которое
    понадобится позже (не придётся перезапрашивать источник).
    """
    source: TenderSource
    source_id: str            # идентификатор в системе источника (может отличаться от reestr_number)
    fetched_at: datetime
    source_payload: dict


@dataclass
class Tender:
    """
    Канонический источник-независимый вид закупки — то, с чем работают
    фильтрация, дедупликация и весь дальнейший пайплайн.
    """
    reestr_number: str                   # реестровый номер ЕИС — главный ключ дедупликации
    law: ProcurementLaw
    method: ProcurementMethod
    status: TenderStatus

    subject: str                          # предмет закупки
    kpgz_code: Optional[str] = None       # код КПГЗ/ОКПД2 — надёжнее ключевых слов для фильтра по категории
    customer_name: Optional[str] = None
    customer_inn: Optional[str] = None
    region: Optional[str] = None
    address: Optional[str] = None

    initial_price: Optional[float] = None  # НМЦК
    currency: str = "RUB"

    publish_date: Optional[date] = None
    submission_deadline: Optional[date] = None
    contract_start: Optional[date] = None
    contract_end: Optional[date] = None

    documents: list[str] = field(default_factory=list)   # ссылки на ТЗ и прочие приложения
    source_url: Optional[str] = None

    sources: list[TenderSource] = field(default_factory=list)  # из скольких источников подтверждена (после дедупа)
    raw: list[TenderRaw] = field(default_factory=list)          # сырые версии от каждого источника


def dedup_key(t: Tender) -> str:
    """
    Реестровый номер уникален в рамках ЕИС и остаётся тем же, даже если
    одна и та же закупка попала к нам и через официальный сервис, и через
    агрегатор, и через площадку.
    """
    if t.reestr_number:
        return t.reestr_number.strip()
    # Фолбэк для источников без реестрового номера (некоторые коммерческие
    # площадки вне 44/223-ФЗ его не публикуют). Составной ключ не идеален
    # (не поймает лёгкие расхождения в формулировке предмета закупки), но
    # лучше, чем не дедуплицировать вовсе.
    return "|".join([
        (t.customer_inn or "").strip(),
        (t.subject or "").strip().lower()[:100],
        str(t.initial_price or ""),
    ])


def merge_duplicates(tenders: list[Tender]) -> list[Tender]:
    """
    Схлопывает закупки с одинаковым dedup_key(), пришедшие из разных
    источников, в одну — объединяя sources/raw и заполняя пропуски полей
    данными из более полной версии (если один источник не дал поле,
    берём из другого, не перезаписывая уже заполненное).
    """
    merged: dict[str, Tender] = {}
    for t in tenders:
        key = dedup_key(t)
        if key not in merged:
            merged[key] = t
            continue
        existing = merged[key]
        for f in ("kpgz_code", "customer_name", "customer_inn", "region", "address",
                  "initial_price", "publish_date", "submission_deadline",
                  "contract_start", "contract_end", "source_url"):
            if getattr(existing, f) is None:
                setattr(existing, f, getattr(t, f))
        existing.documents = list(dict.fromkeys(existing.documents + t.documents))
        existing.sources = list(dict.fromkeys(existing.sources + t.sources))
        existing.raw = existing.raw + t.raw
    return list(merged.values())


@dataclass
class TenderFilter:
    """
    Критерии шага 2. kpgz_prefixes — основной, надёжный фильтр (код КПГЗ не
    меняется от формулировок конкретного заказчика — то же самое "03.08...",
    что было в реальном ТЗ Большого Головина). keywords_any — резерв на
    случай источников, которые код КПГЗ не отдают.
    """
    kpgz_prefixes: list[str] = field(default_factory=list)
    keywords_any: list[str] = field(default_factory=list)
    regions: list[str] = field(default_factory=list)
    min_price: Optional[float] = None
    max_price: Optional[float] = None
    statuses: list[TenderStatus] = field(
        default_factory=lambda: [TenderStatus.ANNOUNCED, TenderStatus.SUBMISSION_OPEN]
    )


def matches_filter(t: Tender, f: TenderFilter) -> bool:
    if f.statuses and t.status not in f.statuses:
        return False

    if f.kpgz_prefixes:
        kpgz_ok = t.kpgz_code is not None and any(t.kpgz_code.startswith(p) for p in f.kpgz_prefixes)
        if not kpgz_ok:
            keyword_ok = f.keywords_any and any(kw.lower() in t.subject.lower() for kw in f.keywords_any)
            if not keyword_ok:
                return False
    elif f.keywords_any:
        if not any(kw.lower() in t.subject.lower() for kw in f.keywords_any):
            return False

    if f.regions and t.region not in f.regions:
        return False
    if f.min_price is not None and (t.initial_price is None or t.initial_price < f.min_price):
        return False
    if f.max_price is not None and (t.initial_price is None or t.initial_price > f.max_price):
        return False
    return True


def filter_tenders(tenders: list[Tender], f: TenderFilter) -> list[Tender]:
    return [t for t in tenders if matches_filter(t, f)]
