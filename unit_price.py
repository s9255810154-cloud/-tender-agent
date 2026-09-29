"""
Нормализация цены к единой единице измерения — чтобы честно сравнивать цены
между поставщиками, у которых один и тот же товар фасован по-разному
(перчатки поштучно vs коробками по 50/100 шт; полотенца по листам/пачкам;
жидкости в разных объёмах).

Без этого шага "более низкая цена" в прайсе может означать просто меньшую
фасовку, а не более выгодную сделку — см. пример с перчатками: 38.36₽ за
1 пару у одного поставщика vs 587.68₽ за коробку 50 пар у другого (то есть
11.75₽/пару — на самом деле ДЕШЕВЛЕ за единицу, хотя число в прайсе больше).

Приоритет распознавания количества в одной приоретной единице:
  1. количество штук/пар в упаковке ("100 шт", "1 пара", "50 шт") —
     самый частый и однозначный случай для дискретных товаров;
  2. количество листов ("N л."/"N лист" в контексте бумаги/полотенец,
     умноженное на число рулонов, если оно указано отдельно);
  3. объём (л / мл) — для жидкостей;
  4. вес (кг / г) — для сыпучих/твёрдых;
  5. длина рулона (м) — для туалетной бумаги/полотенец в рулонах.
Если ни один признак не найден — нормализовать нечем, возвращается цена
"за упаковку как есть" с пометкой unit="уп" (упаковка) — сравнение в этом
случае менее надёжно, и это должно быть видно эксперту.
"""
import math
import re
from typing import Optional

from matcher import extract_numeric_specs, normalize
from models import PriceListItem

_PIECE_COUNT_PATTERN = re.compile(r"(\d+)\s*(?:шт|пар[а-я]*)\b")

# "N л." / "N лист" рядом с "бумага"/"полотенце" — количество листов
# (часто на один рулон), а не литры. Плюс отдельно — количество рулонов
# в упаковке ("2 рул.", "4рулона"), если есть.
_PAPER_CONTEXT_PATTERN = re.compile(r"бумаг|полотенц")
_SHEET_COUNT_PATTERN = re.compile(r"(\d+)\s*(?:л\.?\b|лист)")
_ROLL_COUNT_PATTERN = re.compile(r"(\d+)\s*рул")


def extract_piece_count(text: str) -> Optional[int]:
    """
    Первое найденное "<N> шт" / "<N> пара/пар" в названии — количество
    физических единиц, за которое указана цена в строке прайса.
    Не путать с pack_size ("1/12") — это код заказа коробками, а не то,
    сколько единиц входит в ЭТУ конкретную позицию по ЭТОЙ цене.
    """
    norm = normalize(text)
    m = _PIECE_COUNT_PATTERN.search(norm)
    return int(m.group(1)) if m else None


def extract_sheet_count(text: str) -> Optional[int]:
    """
    Общее количество листов, за которое указана цена в строке — то есть
    "листов на рулон" × "рулонов в упаковке", если рулоны указаны отдельно.

    Работает только в контексте бумаги/полотенец ("бумаг"/"полотенц" в
    названии) — иначе "N л." в других товарных категориях может означать
    что угодно другое, а не листы.

    Примеры:
      "Туалетная бумага листовая 2сл.250л. 21*10,8 ..." -> рул=1 (не указано),
        листов=250 -> итого 250
      "Полотенца бумажные 2сл. Familia 2 рул. 80л. 22,7*12 ..." -> рул=2,
        листов=80 (на рулон) -> итого 160
      "Бумага туалетная 2-сл 21,6 м/рул 180 лист/рул ... 4рулона ..." -> рул=4,
        листов=180 (на рулон) -> итого 720
    """
    norm = normalize(text)
    if not _PAPER_CONTEXT_PATTERN.search(norm):
        return None

    sheet_match = _SHEET_COUNT_PATTERN.search(norm)
    if not sheet_match:
        return None
    sheets_per_roll = int(sheet_match.group(1))

    roll_match = _ROLL_COUNT_PATTERN.search(norm)
    rolls = int(roll_match.group(1)) if roll_match else 1

    return sheets_per_roll * rolls


def normalized_unit_price(item: PriceListItem) -> tuple[float, str]:
    """Возвращает (цена_за_единицу, метка_единицы) для одной позиции прайса."""
    specs = extract_numeric_specs(item.name)
    piece_count = extract_piece_count(item.name)

    if piece_count:
        # piece_count может быть равен 1 ("1 пара") — это значит, что цена в
        # строке УЖЕ указана за одну единицу, и это тоже валидная нормализация,
        # а не "нет данных". Раньше здесь стояло "> 1", из-за чего "1 пара"
        # (Альмин) не считалась нормализуемой и сравнение с коробкой 50 пар
        # (второй поставщик) откатывалось на сырую цену за упаковку.
        return round(item.final_price / piece_count, 4), "шт"

    sheet_count = extract_sheet_count(item.name)
    if sheet_count:
        return round(item.final_price / sheet_count, 4), "лист"

    if "volume_l" in specs and specs["volume_l"] > 0:
        return round(item.final_price / specs["volume_l"], 4), "л"
    if "volume_ml" in specs and specs["volume_ml"] > 0:
        return round(item.final_price / (specs["volume_ml"] / 1000), 4), "л"
    if "weight_kg" in specs and specs["weight_kg"] > 0:
        return round(item.final_price / specs["weight_kg"], 4), "кг"
    if "weight_g" in specs and specs["weight_g"] > 0:
        return round(item.final_price / (specs["weight_g"] / 1000), 4), "кг"
    if "length_m" in specs and specs["length_m"] > 0:
        return round(item.final_price / specs["length_m"], 4), "м"

    return item.final_price, "уп"  # нечем нормализовать — сравнение по цене за упаковку целиком


# Насколько ниже лучшего score кандидат ещё считается "тем же товаром у другого
# поставщика", а не менее подходящим совпадением. Подобрано эмпирически: два
# реально одинаковых товара (мешки 120л/20мкм у двух поставщиков) обычно
# расходятся в score на 0.01–0.05 из-за разной пунктуации/порядка слов в
# наименовании — не из-за того, что это разные товары.
PRICE_EQUIVALENCE_TOLERANCE = 0.06


def select_cheapest_equivalent(
    scored_candidates: list[tuple[PriceListItem, float]],
) -> tuple[PriceListItem, float, list[tuple[PriceListItem, float]], bool]:
    """
    Среди кандидатов, чей score достаточно близок к лучшему (т.е. это, скорее
    всего, один и тот же товар, просто у разных поставщиков), выбирает самую
    выгодную позицию.

    Сравнение идёт по ЦЕНЕ ЗА ЕДИНИЦУ (normalized_unit_price), а не по сырой
    цене строки — иначе коробка на 50 пар "дороже" одной пары чисто по числу,
    хотя может быть выгоднее в пересчёте на пару.

    Если у равнозначных кандидатов единицы нормализации не совпадают
    (например для одного нашли "шт", для другого — только "уп", то есть
    нечем нормализовать) — честно сравнить нельзя. В этом случае откатываемся
    на сравнение по цене за упаковку как есть и возвращаем comparable=False,
    чтобы вызывающий код мог это показать.

    Возвращает (выбранная_позиция, её_score, все_равнозначные_кандидаты, comparable).
    """
    best_score = scored_candidates[0][1]
    equivalent = [
        pair for pair in scored_candidates
        if pair[1] >= best_score - PRICE_EQUIVALENCE_TOLERANCE
    ]

    normalized = [
        (item, score, *normalized_unit_price(item)) for item, score in equivalent
    ]
    units_present = {unit for *_, unit in normalized}

    if len(units_present) == 1 and "уп" not in units_present:
        # Все кандидаты нормализовались к одной и той же единице — сравнение честное.
        item, score, _, _ = min(normalized, key=lambda row: row[2])
        return item, score, equivalent, True

    # Разные единицы (или нечем нормализовать хотя бы у одного) — сравниваем
    # по цене за упаковку "как есть", предупреждая, что это менее надёжно.
    item, score = min(equivalent, key=lambda pair: pair[0].final_price)
    return item, score, equivalent, False


def estimate_purchase_quantity(
    natural_qty: float, natural_unit: str, matched_item: PriceListItem
) -> Optional[int]:
    """
    Переводит "сколько нужно физически" (натуральная потребность — литры,
    килограммы или штуки) в "сколько упаковок заказать" по факту фасовки
    НАЙДЕННОГО товара — переиспользует ту же экстракцию фасовки, что и
    normalized_unit_price (extract_piece_count / extract_numeric_specs).

    Пример: нужно 3719.6 л моющего средства для полов; найденный товар
    фасован по 5 л; закупка = ceil(3719.6 / 5) = 744 канистры.

    natural_unit: "шт" | "л" | "кг". Возвращает None, если у найденного
    товара нет опознаваемой фасовки в этой единице (тогда придётся
    вводить количество вручную — как раньше).
    """
    specs = extract_numeric_specs(matched_item.name)

    if natural_unit == "шт":
        pack = extract_piece_count(matched_item.name) or 1
    elif natural_unit == "л":
        if specs.get("volume_l"):
            pack = specs["volume_l"]
        elif specs.get("volume_ml"):
            pack = specs["volume_ml"] / 1000
        else:
            return None
    elif natural_unit == "кг":
        if specs.get("weight_kg"):
            pack = specs["weight_kg"]
        elif specs.get("weight_g"):
            pack = specs["weight_g"] / 1000
        else:
            return None
    else:
        return None

    if not pack or pack <= 0:
        return None
    return math.ceil(natural_qty / pack)
