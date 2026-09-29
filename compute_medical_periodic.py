import json

with open('knowledge_base/medical_premises_gesn_rates.json', encoding='utf-8') as f:
    rates = {it['name']: it for it in json.load(f)['items']}
with open('medical_fixture_totals.json', encoding='utf-8') as f:
    t = json.load(f)

MAPPING = [
    ("Мытье пола в административных кабинетах", t['admin_kabinety_m2'], 100, 12, "ежемесячно (допущение, как у Лужников)"),
    ("Мытье пола в вестибюлях, коридорах, фойе, холлах, рекреациях", t['koridory_m2'] + t['vhodnaya_gruppa_m2'], 100, 12, "ежемесячно"),
    ("Мытье пола в санузлах", t['sanuzly_m2'], 100, 12, "ежемесячно"),
    ("Мытье лестниц с протиркой перил", t['lestnicy_m2'], 100, 12, "ежемесячно"),
    ("Мытье внутренней поверхности окон (кроме помещений класса А)", t['okna_m2'], 100, 4, "1 раз в 3 мес — подтверждено регламентом"),
    ("Мытье наружной поверхности окон", t['okna_m2'], 100, 4, "1 раз в 3 мес — подтверждено регламентом"),
    ("Мытье дверей однопольных (кроме помещений класса А)", t['dveri_sht'], 100, 12, "ежемесячно (допущение)"),
    ("Мытье стен, колонн (кроме помещений класса А)", t['area_total'], 100, 4, "ежеквартально (допущение)"),
]

zp_year, mr_year = 0, 0
print(f"{'Операция':55} {'Кол-во':>10} {'Раз/год':>8} {'ЗП/год':>14} {'МР/год':>12}")
for name, qty, unit_scale, freq, note in MAPPING:
    r = rates.get(name)
    if not r:
        print(f"НЕ НАЙДЕНО: {name}")
        continue
    scale = qty / unit_scale
    zp = r['zp'] * scale * freq
    mr = r['mr'] * scale * freq
    zp_year += zp
    mr_year += mr
    print(f"{name[:55]:55} {qty:>10,.1f} {freq:>8} {zp:>14,.0f} {mr:>12,.0f}")

print()
print(f"ИТОГО ЗП/год (периодические операции): {zp_year:,.2f} руб")
print(f"ИТОГО МР/год (материалы для этих операций): {mr_year:,.2f} руб")

with open("medical_periodic_cost.json", "w", encoding="utf-8") as f:
    json.dump({"zp_year": zp_year, "mr_year": mr_year}, f, ensure_ascii=False, indent=2)
