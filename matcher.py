"""
Fuzzy-сопоставление формулировки из ТЗ с позициями прайсов поставщиков.

Проблема: ТЗ описывает товар через ГОСТ/спецификацию
  "Туалетная бумага от 16 м. 2х ГОСТ Р 52354-2025"
а прайс — через бренд и маркетинговое название
  "Бумага туалетная 3-сл 16,8 м/рул 140 лист/рул PAPIA PROF ... 'PAPIA'"

Чистое строковое сходство (difflib) здесь слабое: общих слов мало, а важные
числа (длина рулона, слойность) могут быть в разных позициях строки.

Итоговый score — комбинация текстового сходства, категориального совпадения
(тип товара) и сходства "атрибутов". Атрибуты бывают двух родов:
  - числовые (объём, вес, длина рулона, слойность) — для "жидких"/"бумажных"
    товаров это самый надёжный сигнал;
  - категориальные (размер XL/L/M, материал — латекс/винил/микрофибра) —
    для товаров вроде перчаток или насадок числовых характеристик почти нет,
    и раньше это занижало score. Теперь вес атрибутов распределяется по
    ТИПУ товара: если у обеих сторон нет ни одного сравнимого числового
    признака, но есть размер/материал — на них и опираемся; если вообще
    нет сравнимых атрибутов — больше веса уходит на текст и категорию.

В проде рекомендуется заменить SequenceMatcher на rapidfuzz (token_sort_ratio /
WRatio) для скорости на больших прайсах — интерфейс scoring-функции такой же.
"""
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Optional

from models import MatchResult, PriceListItem, Source, TZItem

# Категории по НАБОРАМ ОСНОВ СЛОВ (стемов), а не по фразам целиком. Раньше
# словарь был словарём точных фраз ("средство чистящ", "пакет мусорн") — это
# ломалось на любой другой словоформе или порядке слов ("чистящее средство",
# "пакеты для мусора"). Реальные ТЗ формулируют почти всегда иначе, чем
# прайсы, поэтому нужна устойчивость к словоформам, а не точное совпадение.
#
# Правило — список ТРЕБОВАНИЙ (AND), каждое требование — список АЛЬТЕРНАТИВ
# (OR). Альтернатива это либо ("prefix", "стем") — любое слово в тексте
# начинается с этой основы, либо ("exact", {набор форм}) — слово совпадает
# ровно с одной из перечисленных словоформ (нужно для коротких основ вроде
# "пол", которые иначе ложно сработают внутри "полотенце").
#
# Порядок важен: более специфичные правила (стиральный/чистящий порошок,
# дезинфицирующее средство) должны стоять раньше общего "чистящее средство",
# иначе общее правило перехватит их первым.
_FLOOR_FORMS = {"пол", "пола", "полу", "полом", "поле", "полы", "полов", "полам", "полах", "полами"}

CATEGORY_RULES: list[tuple[str, list[list[tuple[str, object]]]]] = [
    ("мусорный мешок", [
        [("prefix", "пакет"), ("prefix", "мешк"), ("exact", {"мешок"})],
        [("prefix", "мусор")],
    ]),
    ("жидкое мыло", [[("prefix", "мыл")], [("prefix", "жидк")]]),
    ("туалетное мыло", [[("prefix", "мыл")], [("prefix", "туалет")]]),
    ("хозяйственное мыло", [[("prefix", "мыл")], [("prefix", "хозяйств")]]),
    ("туалетная бумага", [[("prefix", "туалет")], [("prefix", "бумаг")]]),
    ("бумажное полотенце", [[("prefix", "полотенц")]]),
    ("перчатки", [[("prefix", "перчат")]]),
    ("губка", [[("prefix", "губк")]]),
    ("комплект для уборки", [[("prefix", "комплект")], [("prefix", "уборк")]]),
    ("моп/насадка для швабры", [[("prefix", "моп"), ("prefix", "насадк"), ("prefix", "швабр")]]),
    ("щётка/совок", [[("prefix", "щетк"), ("prefix", "совок")]]),
    ("средство для мытья пола", [
        [("prefix", "мыт"), ("prefix", "мою"), ("prefix", "мойк")],
        [("exact", _FLOOR_FORMS)],
    ]),
    ("средство для мытья посуды", [
        [("prefix", "мыт"), ("prefix", "мою"), ("prefix", "мойк")],
        [("prefix", "посуд")],
    ]),
    ("средство для стёкол", [[("prefix", "стек"), ("prefix", "зеркал")]]),
    ("дезинфицирующее средство", [[("prefix", "дезинфиц")]]),
    ("стиральный порошок", [[("prefix", "стира")], [("prefix", "порош")]]),
    ("чистящий порошок", [[("prefix", "чист")], [("prefix", "порош")]]),
    ("отбеливатель", [[("prefix", "отбелив"), ("prefix", "белизн")]]),
    ("освежитель воздуха", [[("prefix", "освежит")]]),
    ("салфетка", [[("prefix", "салфетк")]]),
    # Общая категория — намеренно в самом конце: "Чистящее средство для
    # унитаза/сантехники/интерьера/оргтехники" и т.п. не имеют устойчивого
    # отдельного стема, поэтому ловятся сюда, если не подошло ничего точнее.
    ("чистящее средство", [[("prefix", "чист")], [("prefix", "средств")]]),
]

# Числовые характеристики, которые важно сопоставлять точно.
_NUM_PATTERNS = {
    "length_m": re.compile(r"(\d+[.,]?\d*)\s*м(?!г|л)\b"),      # длина в метрах (не мг/мл)
    "volume_l": re.compile(r"(\d+[.,]?\d*)\s*л\b"),               # объём в литрах
    "volume_ml": re.compile(r"(\d+[.,]?\d*)\s*мл\b"),             # объём в мл
    "weight_g": re.compile(r"(\d+[.,]?\d*)\s*г\b"),               # вес в граммах
    "weight_kg": re.compile(r"(\d+[.,]?\d*)\s*кг\b"),             # вес в кг
    "layers": re.compile(r"(\d)[-\s]?сл|(\d)[xх]\b"),             # слойность 2-сл / 2х
}

# Размеры одежды/СИЗ — латинские буквы, не путаются с кириллическими
# единицами измерения (л, м, г — те распознаются отдельным паттерном выше).
_SIZE_PATTERN = re.compile(r"\b(xxl|xxs|xl|xs|[sml])\b")

# Материалы — категориальный атрибут, важен там, где почти нет чисел
# (перчатки, насадки, полотно).
_MATERIAL_KEYWORDS = [
    "латекс", "винил", "нитрил", "пвх", "хлопок", "микрофибра",
    "поролон", "целлюлоза", "полиэстер", "резин",
]

_WORD_PATTERN = re.compile(r"[a-zа-я0-9]+")


def normalize(text: str) -> str:
    text = text.lower().replace("ё", "е")
    text = re.sub(r'["\']', " ", text)
    text = re.sub(r"[^\w\s.,/-]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _tokens(norm_text: str) -> list[str]:
    return _WORD_PATTERN.findall(norm_text)


def _alt_matches(tokens: list[str], alt: tuple[str, object]) -> bool:
    kind, value = alt
    if kind == "prefix":
        return any(t.startswith(value) for t in tokens)
    return any(t in value for t in tokens)  # kind == "exact"


def extract_category(text: str) -> Optional[str]:
    norm = normalize(text)
    tokens = _tokens(norm)
    for category, requirements in CATEGORY_RULES:
        if all(any(_alt_matches(tokens, alt) for alt in group) for group in requirements):
            return category
    return None


def extract_numeric_specs(text: str) -> dict:
    norm = normalize(text)
    specs = {}
    for key, pattern in _NUM_PATTERNS.items():
        m = pattern.search(norm)
        if m:
            val = next(g for g in m.groups() if g)
            specs[key] = float(val.replace(",", "."))

    # "N л." в НАЗВАНИЯХ БУМАЖНЫХ ИЗДЕЛИЙ часто значит "N листов" (сокращение
    # "л." = "листов"), а не литры — например "250л." у туалетной бумаги.
    # Ограничиваем это контекстом бумаги/полотенец: у мешков для мусора
    # "120 л" — это настоящие литры (и легитимно больше 30), их трогать нельзя.
    if "бумаг" in norm or "полотенц" in norm:
        specs.pop("volume_l", None)

    return specs


def extract_size(text: str) -> Optional[str]:
    norm = normalize(text)
    m = _SIZE_PATTERN.search(norm)
    return m.group(1) if m else None


def extract_materials(text: str) -> set[str]:
    norm = normalize(text)
    return {kw for kw in _MATERIAL_KEYWORDS if kw in norm}


def _numeric_similarity(a: dict, b: dict) -> Optional[float]:
    """Доля совпадающих (в пределах допуска) числовых характеристик среди общих ключей.
    None, если сравнивать нечего (нет общих числовых ключей у обеих сторон)."""
    common = set(a) & set(b)
    if not common:
        return None
    hits = 0
    for key in common:
        tol = 0.15 if key not in ("layers",) else 0.0  # 15% допуск, слойность — точно
        if abs(a[key] - b[key]) <= tol * max(a[key], 1):
            hits += 1
    return hits / len(common)


def _attribute_similarity(tz_item: TZItem, candidate: PriceListItem) -> Optional[float]:
    """
    Сводит числовые и категориальные атрибуты (размер, материал) в один сигнал.
    Возвращает None, если вообще нет ни одной сравнимой пары атрибутов —
    тогда score_candidate компенсирует это весом текста/категории.

    Важно: если ТЗ явно называет атрибут (например "латексные" или "XL"), а у
    кандидата этот атрибут не указан вовсе — это считается расхождением (0.0),
    а не пропускается. Раньше отсутствие данных у кандидата не штрафовалось
    никак, из-за чего товар без указания материала мог обойти по score товар
    с точно совпадающим материалом (пекарские перчатки против хозяйственных
    латексных — оба одного размера, но у пекарских просто не сказано, из
    чего они сделаны, и это не должно засчитываться как совпадение).
    """
    tz_specs = extract_numeric_specs(tz_item.raw_name)
    cand_specs = extract_numeric_specs(candidate.name)
    num_sim = _numeric_similarity(tz_specs, cand_specs)

    tz_size = extract_size(tz_item.raw_name)
    cand_size = extract_size(candidate.name)
    size_sim = None
    if tz_size:
        size_sim = 1.0 if tz_size == cand_size else 0.0  # cand_size=None -> не подтверждено -> 0.0

    tz_mats = extract_materials(tz_item.raw_name)
    cand_mats = extract_materials(candidate.name)
    material_sim = None
    if tz_mats:
        if cand_mats:
            union = tz_mats | cand_mats
            material_sim = len(tz_mats & cand_mats) / len(union) if union else 0.0
        else:
            material_sim = 0.0  # ТЗ требует материал явно, кандидат не подтверждает

    parts = [p for p in (num_sim, size_sim, material_sim) if p is not None]
    if not parts:
        return None
    return sum(parts) / len(parts)


def score_candidate(tz_item: TZItem, candidate: PriceListItem) -> float:
    tz_norm = normalize(tz_item.raw_name)
    cand_norm = normalize(candidate.name)

    text_sim = SequenceMatcher(None, tz_norm, cand_norm).ratio()

    tz_cat = extract_category(tz_item.raw_name)
    cand_cat = extract_category(candidate.name)
    category_match = 1.0 if (tz_cat and tz_cat == cand_cat) else 0.0

    # Категория — обязательное условие (без неё сходство почти всегда шум).
    if tz_cat and not category_match:
        return text_sim * 0.15

    attr_sim = _attribute_similarity(tz_item, candidate)

    if attr_sim is None:
        # Нет ни одного сравнимого атрибута (ни чисел, ни размера, ни материала) —
        # полагаемся на текст и категорию сильнее.
        return 0.55 * text_sim + 0.45 * category_match

    # Есть чем сверить атрибуты — они самый надёжный сигнал.
    return 0.20 * text_sim + 0.55 * attr_sim + 0.25 * category_match


CONFIDENCE_THRESHOLD_AUTO = 0.75   # выше — можно брать автоматически
CONFIDENCE_THRESHOLD_REVIEW = 0.45  # ниже — не показывать вовсе, сразу not_found


def build_category_index(price_lists: list[PriceListItem]) -> dict[Optional[str], list[PriceListItem]]:
    """
    Группирует прайс по категории ОДИН раз — вместо того чтобы на каждую
    позицию ТЗ гонять score_candidate() (с SequenceMatcher и регэкспами) по
    всем тысячам строк прайса. На объединённом прайсе (~4100 позиций) это
    самая дорогая часть пайплайна при десятках/сотнях позиций в одном ТЗ.

    Строится один раз в resolve_all() и переиспользуется для каждой позиции
    ТЗ. Позиции без распознанной категории складываются под ключом None и
    участвуют только тогда, когда категория самого ТЗ тоже не распозналась.
    """
    index: dict[Optional[str], list[PriceListItem]] = {}
    for item in price_lists:
        cat = extract_category(item.name)
        index.setdefault(cat, []).append(item)
    return index


def match_against_pricelists(
    tz_item: TZItem,
    price_lists: list[PriceListItem],
    top_n: int = 8,
    category_index: Optional[dict[Optional[str], list[PriceListItem]]] = None,
) -> list[tuple[PriceListItem, float]]:
    tz_cat = extract_category(tz_item.raw_name)

    if category_index is not None and tz_cat is not None:
        # Предфильтр: считаем полный score только для позиций той же категории.
        # Если категория ТЗ вообще не встретилась в прайсе — пул пуст, и это
        # корректный результат (ничего похожего нет), а не повод сканировать всё.
        candidates_pool = category_index.get(tz_cat, [])
    else:
        candidates_pool = price_lists

    scored = [(item, score_candidate(tz_item, item)) for item in candidates_pool]
    scored.sort(key=lambda pair: pair[1], reverse=True)
    return [pair for pair in scored if pair[1] >= CONFIDENCE_THRESHOLD_REVIEW][:top_n]
