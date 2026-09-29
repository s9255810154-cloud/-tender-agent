import json

with open("medical_periodic_cost.json", encoding="utf-8") as f:
    periodic = json.load(f)

ZP_PERIODIC = periodic["zp_year"]       # 2,107,755.10
MR_PERIODIC = periodic["mr_year"]       # 271,702.67
NR_PCT, SP_PCT = 0.70, 0.10             # конвенция ЭТОЙ сметы (ГЭСН, как в позициях "НР от ЗП"/"СП от ЗП")

NMCK_WITH_VAT = 29_000_000.0
VAT_RATE = 0.22
NMCK_WITHOUT_VAT = NMCK_WITH_VAT / (1 + VAT_RATE)

# Периодическая часть (подтверждённые расценки + допущения по частоте)
periodic_total_no_vat = ZP_PERIODIC * (1 + NR_PCT + SP_PCT) + MR_PERIODIC
print(f"НМЦК без НДС: {NMCK_WITHOUT_VAT:,.2f} руб")
print(f"Периодическая часть (ЗП+НР+СП+материалы): {periodic_total_no_vat:,.2f} руб")

remaining_for_daily = NMCK_WITHOUT_VAT - periodic_total_no_vat
print(f"Остаток на ежедневный штат: {remaining_for_daily:,.2f} руб")

# Ежедневная часть — предполагаем ту же формулу, что у Лужников (затратный метод):
# ЗП * (1 + ФОТ 30.2% + НР 20% + СП 5%) = ЗП * 1.552
FOT_PCT_DAILY, NR_PCT_DAILY, SP_PCT_DAILY = 0.302, 0.20, 0.05
daily_multiplier = 1 + FOT_PCT_DAILY + NR_PCT_DAILY + SP_PCT_DAILY

zp_daily_base = remaining_for_daily / daily_multiplier
print(f"Подразумеваемая базовая ЗП ежедневного штата: {zp_daily_base:,.2f} руб/год")

EMISS_SALARY = 79181.10
staff_implied = zp_daily_base / (EMISS_SALARY * 12)
print(f"Подразумеваемая численность (при ЗП ЕМИСС): {staff_implied:.2f} человек ~ {round(staff_implied)}")

sqm_per_person_implied = 16026.9 / staff_implied
print(f"Подразумеваемый норматив: {sqm_per_person_implied:,.1f} м²/чел")

# ---- Твёрдый пол (та же логика, что у Лужников/дисконта) ----
zp_daily_fot = zp_daily_base * FOT_PCT_DAILY
zp_periodic_fot_equiv = ZP_PERIODIC * NR_PCT  # обязательная часть НР по методике этой сметы трактуем как overhead, не чистую ЗП-надбавку -> в твёрдый пол НЕ включаем, только ЗП+материалы+обязательные страх.взносы

floor = ZP_PERIODIC + MR_PERIODIC + zp_daily_base + zp_daily_fot
print()
print(f"ТВЁРДЫЙ ПОЛ (ЗП периодич. + материалы периодич. + ЗП ежедневная + ФОТ-надбавка 30.2% на ежедневную):")
print(f"  {floor:,.2f} руб (без НДС)")

max_safe_discount = 1 - floor / NMCK_WITHOUT_VAT
print(f"Максимальная безопасная скидка от НМЦК: {max_safe_discount*100:.1f}%")

with open("medical_floor_calc.json", "w", encoding="utf-8") as f:
    json.dump({
        "nmck_without_vat": NMCK_WITHOUT_VAT,
        "periodic_total_no_vat": periodic_total_no_vat,
        "zp_daily_base": zp_daily_base,
        "staff_implied": staff_implied,
        "floor": floor,
        "max_safe_discount": max_safe_discount,
    }, f, ensure_ascii=False, indent=2)
