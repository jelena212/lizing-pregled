import pandas as pd

ugovori  = pd.read_excel("lizing_pregled.xlsx", sheet_name = "Ugovori")
rate = pd.read_excel("lizing_pregled.xlsx", sheet_name = "Rate")

rate = rate.drop_duplicates()
print("Provera sumnjivih")
print(rate[rate["Iznos rate (EUR)"] <= 0])

# 3. OBRADA
placeno = rate.groupby("Ugovor")["Iznos rate (EUR)"].sum().reset_index()
placeno.columns = ["Ugovor", "Placeno (EUR)"]
pregled = ugovori.merge(placeno, on="Ugovor")
pregled["Ostatak duga (EUR)"] = pregled["Ukupan iznos (EUR)"] - pregled["Placeno (EUR)"]

# 4. IZLAZ
print(pregled[["Ugovor", "Ukupan iznos (EUR)", "Placeno (EUR)", "Ostatak duga (EUR)"]])
pregled.to_excel("lizing_izvestaj.xlsx", index=False)