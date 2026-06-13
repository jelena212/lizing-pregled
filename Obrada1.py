"""
Modul 1 — Lizing: pregled plaćanja i ostatka duga
=================================================
Jedan skript koji uzme sirovi Excel (lizing_pregled.xlsx) i izbaci gotov izveštaj.
Obrazac: 1. UČITAVANJE  2. ČIŠĆENJE  3. OBRADA  4. IZLAZ
"""

import pandas as pd


# ============================================================
# 1. UČITAVANJE
# ============================================================
ugovori = pd.read_excel("lizing_pregled.xlsx", sheet_name="Ugovori")
rate     = pd.read_excel("lizing_pregled.xlsx", sheet_name="Rate")

# Brza provera šta smo učitali (otkomentariši po potrebi)
# print(ugovori.head())
# print(rate.info())
# print(rate.columns)   # tačna imena kolona — kad nisi sigurna kako se zovu


# ============================================================
# 2. ČIŠĆENJE  (uvek PRE računanja!)
# ============================================================

# 2a. Gde su rupe? (informativno)
print("Nedostajuće vrednosti po koloni:")
print(rate.isna().sum())
print()

# 2b. Duplikati — isti red dvaput naduva sumu
print("Broj duplih redova:", rate.duplicated().sum())
rate = rate.drop_duplicates()

# 2c. Sumnjivi unosi — kod ne zna šta je normalno, samo ih POKAŽE
print("\nSumnjive rate (<= 0 ili apsurdno velike):")
print(rate[(rate["Iznos rate (EUR)"] <= 0) | (rate["Iznos rate (EUR)"] > 100000)])
print()

# Prazne rate ostavljamo kao NaN = "nije plaćeno" (to je informacija, ne greška).
# groupby ih ionako preskače pri sabiranju.

# 2d. Tekst -> datum (da bi se sortirao hronološki, ne abecedno)
rate["Mesec"] = pd.to_datetime(rate["Mesec"])
rate["Mesec"] = rate["Mesec"].dt.to_period("M")   # konkretan dan -> ceo mesec


# ============================================================
# 3. OBRADA
# ============================================================

# 3a. Ukupno plaćeno po ugovoru (groupby = saberi po jednoj fioci)
placeno = rate.groupby("Ugovor")["Iznos rate (EUR)"].sum().reset_index()
placeno.columns = ["Ugovor", "Placeno (EUR)"]

# 3b. Spoji sa ugovorima (merge = VLOOKUP za sve redove odjednom)
pregled = ugovori.merge(placeno, on="Ugovor")

# 3c. Nova kolona — ostatak duga (aritmetika nad celim kolonama)
pregled["Ostatak duga (EUR)"] = pregled["Ukupan iznos (EUR)"] - pregled["Placeno (EUR)"]

# 3d. Pivot — plaćanja po mesecu i ugovoru (mreža: redovi × kolone)
pivot = rate.pivot_table(
    index="Ugovor",
    columns="Mesec",
    values="Iznos rate (EUR)",
    aggfunc="sum",
    fill_value=0,
    margins=True,
    margins_name="Ukupno",
)


# ============================================================
# 4. IZLAZ
# ============================================================
print("=== Pregled duga po ugovoru ===")
print(pregled[["Ugovor", "Ukupan iznos (EUR)", "Placeno (EUR)", "Ostatak duga (EUR)"]])
print()
print("=== Plaćanja po mesecu (pivot) ===")
print(pivot)

# Snimi u Excel — dva sheet-a u jednom fajlu
with pd.ExcelWriter("lizing_izvestaj.xlsx") as writer:
    pregled.to_excel(writer, sheet_name="Pregled duga", index=False)
    pivot.to_excel(writer, sheet_name="Po mesecu")

print("\nIzveštaj snimljen u lizing_izvestaj.xlsx")
