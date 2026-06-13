import pandas as pd

# ── 1. UČITAVANJE ─────────────────────────────────
ugovori = pd.read_excel("lizing_pregled.xlsx", sheet_name="Ugovori")
rate = pd.read_excel("lizing_pregled.xlsx", sheet_name="Rate")

# ── 2. ČIŠĆENJE / TIPOVI ──────────────────────────
# tekst → datum → mesec (odmah posle učitavanja, pre svake obrade)
rate["Mesec"] = pd.to_datetime(rate["Mesec"])
rate["Mesec"] = rate["Mesec"].dt.to_period("M")

# ── 3. IZVEŠTAJ 1: ostatak duga po ugovoru ────────
placeno = rate.groupby("Ugovor")["Iznos rate (EUR)"].sum()
placeno = placeno.reset_index()
placeno.columns = ["Ugovor", "Placeno (EUR)"]

pregled = ugovori.merge(placeno, on="Ugovor")
pregled["Ostatak duga (EUR)"] = pregled["Ukupan iznos (EUR)"] - pregled["Placeno (EUR)"]

print(pregled[["Ugovor", "Ukupan iznos (EUR)", "Placeno (EUR)", "Ostatak duga (EUR)"]])
pregled.to_excel("lizing_izvestaj.xlsx", index=False)
print("\nIzvestaj sacuvan: lizing_izvestaj.xlsx")

# ── 4. IZVEŠTAJ 2: pivot — plaćanja po mesecima ───
tabela = rate.pivot_table(
    index="Ugovor",
    columns="Mesec",
    values="Iznos rate (EUR)",
    aggfunc="sum",
    fill_value=0,
    margins = True
)
print(tabela)
print(rate.isna().sum())