from datetime import date, datetime

from tender_models import (
    ProcurementLaw,
    ProcurementMethod,
    Tender,
    TenderFilter,
    TenderRaw,
    TenderSource,
    TenderStatus,
    filter_tenders,
    merge_duplicates,
)

# --- Позиция 1: реальная закупка (Большой Головин, 15) от ОФИЦИАЛЬНОГО сервиса ЕИС ---
# КПГЗ-код — прямо из документа, который мы разбирали.
tender_from_eis = Tender(
    reestr_number="0173200001426000123",  # пример номера (в самом ТЗ конкретный номер не был виден в тексте)
    law=ProcurementLaw.FZ_44,
    method=ProcurementMethod.AUCTION,
    status=TenderStatus.ANNOUNCED,
    subject="Оказание услуг по уборке помещений в 2026-2029 годах",
    kpgz_code="03.08.01.01.01.07",
    customer_name="Государственное учреждение г. Москвы",
    region="Москва",
    address="г. Москва, Большой Головин переулок, 15",
    initial_price=7939593.68,  # из нашей же сметы, для примера
    publish_date=date(2026, 9, 1),
    submission_deadline=date(2026, 9, 25),
    contract_start=date(2026, 11, 1),
    contract_end=date(2029, 10, 31),
    documents=["Техническое_задание_17204828-1.pdf"],
    source_url="https://zakupki.gov.ru/epz/order/notice/ea44/view/common-info.html?regNumber=0173200001426000123",
    sources=[TenderSource.EIS_OFFICIAL],
    raw=[TenderRaw(
        source=TenderSource.EIS_OFFICIAL,
        source_id="0173200001426000123",
        fetched_at=datetime(2026, 9, 18, 10, 0),
        source_payload={"raw": "полный SOAP-ответ ЕИС здесь"},
    )],
)

# --- Та же закупка, но пришла ВТОРОЙ раз — через платный агрегатор ---
# У агрегатора нет поля kpgz_code в этом примере (не все источники отдают всё),
# зато есть дата публикации и статус — merge_duplicates дозаполнит пробелы.
tender_from_aggregator = Tender(
    reestr_number="0173200001426000123",  # тот же номер — это и есть основа дедупа
    law=ProcurementLaw.FZ_44,
    method=ProcurementMethod.AUCTION,
    status=TenderStatus.ANNOUNCED,
    subject="Оказание услуг по уборке помещений в 2026-2029 годах",
    kpgz_code=None,  # агрегатор в этом примере не отдал код КПГЗ
    initial_price=7939593.68,
    source_url="https://multitender.ru/tender/0173200001426000123",
    sources=[TenderSource.AGGREGATOR],
    raw=[TenderRaw(
        source=TenderSource.AGGREGATOR,
        source_id="agg-98765",
        fetched_at=datetime(2026, 9, 18, 10, 5),
        source_payload={"raw": "ответ агрегатора здесь"},
    )],
)

# --- Шумовая закупка — другая категория, не должна пройти фильтр ---
tender_noise = Tender(
    reestr_number="0173200001426000999",
    law=ProcurementLaw.FZ_44,
    method=ProcurementMethod.QUOTATION_REQUEST,
    status=TenderStatus.ANNOUNCED,
    subject="Поставка канцелярских товаров",
    kpgz_code="02.01.01.01.01.01",
    region="Москва",
    initial_price=150000,
    sources=[TenderSource.EIS_OFFICIAL],
)

all_tenders = [tender_from_eis, tender_from_aggregator, tender_noise]

print(f"Всего получено (до дедупа): {len(all_tenders)}")
merged = merge_duplicates(all_tenders)
print(f"После дедупликации: {len(merged)}")
for t in merged:
    print(f"  {t.reestr_number} | источники: {[s.value for s in t.sources]} | "
          f"кпгз: {t.kpgz_code} | статус: {t.status.value}")

# Фильтр — ровно тот же код КПГЗ, что в реальном ТЗ ("санитарное содержание/уборка")
cleaning_filter = TenderFilter(
    kpgz_prefixes=["03.08"],
    keywords_any=["уборк", "клининг", "санитарное содержание"],  # резерв, если КПГЗ не пришёл
    regions=["Москва"],
    max_price=15_000_000,
)

result = filter_tenders(merged, cleaning_filter)
print(f"\nПрошли фильтр (клининг, Москва, до 15 млн): {len(result)}")
for t in result:
    print(f"  {t.reestr_number} | {t.subject} | НМЦК: {t.initial_price:,.2f} ₽")
