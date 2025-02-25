import pandas as pd

# Percorsi file
input_path = r"IBSA_Report fatture output.xlsx"
output_path = r"IBSA_Report fatture output - 2.xlsx"

# Caricamento del file Excel senza modificare i tipi
df = pd.read_excel(input_path, dtype=str)

# Pulizia nomi colonne (rimuove spazi extra)
df.columns = df.columns.str.strip()

# Conversione delle colonne numeriche
df["Somma di IMPONIBILE EUR"] = df["Somma di IMPONIBILE EUR"].str.replace(',', '.').astype(float)
df["Somma di QTY PRODOTTO"] = df["Somma di QTY PRODOTTO"].str.replace(',', '.').astype(float)

# Itera su tutte le righe
for idx, row in df.iterrows():
    if pd.notna(row["LOTTO"]) and row["LOTTO"] != "":
        lotto_values = str(row["LOTTO"]).split('-')
        qty = int(abs(row["Somma di QTY PRODOTTO"]))  # Considera sempre il valore assoluto
        
        # La quantità di lotti deve essere esattamente uguale alla quantità di prodotto
        last_lotto = lotto_values[-1]  # Prende l'ultimo lotto
        lotto_values.extend([last_lotto] * (qty - len(lotto_values)))  # Aggiunge il lotto mancante

        # Aggiorna la cella
        df.at[idx, "LOTTO"] = '-'.join(lotto_values)

# Esportiamo il file aggiornato
df.to_excel(output_path, index=False, engine='openpyxl')

print("✅ File aggiornato e salvato in:", output_path)
