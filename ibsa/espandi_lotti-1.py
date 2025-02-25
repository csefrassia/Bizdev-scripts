import pandas as pd 

# Percorsi file
input_path = r"IBSA_Report fatture.csv"
output_path = r"IBSA_Report fatture output.xlsx"

# Caricamento del file CSV senza modificare i tipi
df = pd.read_csv(input_path, dtype=str)

# Pulizia nomi colonne (rimuove spazi extra)
df.columns = df.columns.str.strip()

# Debug: controlliamo i dati originali
print("\U0001F50D Controllo iniziale (Prime 5 righe)")
print(df[["ORDINE #", "SKU COMPLETO", "DOC TYPE", "Somma di IMPONIBILE EUR" ]].head())

# **Sostituzione della virgola con il punto per la corretta conversione numerica**
df["Somma di IMPONIBILE EUR"] = df["Somma di IMPONIBILE EUR"].str.replace(',', '.', regex=False)
df["Somma di IMPONIBILE EUR"] = pd.to_numeric(df["Somma di IMPONIBILE EUR"], errors='coerce')

# **Creiamo una copia per sicurezza**
df_original = df.copy()

# Dizionario per gestire i lotti disponibili
lotti_map = {}

# Creiamo una mappa dei lotti per ogni combinazione ORDINE # + SKU COMPLETO + DOC TYPE con imponibile > 0
for _, row in df[df["Somma di IMPONIBILE EUR"] > 0].iterrows():
    key = (row["ORDINE #"], row["SKU COMPLETO"], row["DOC TYPE"])
    
    if pd.isna(row["LOTTO"]) or row["LOTTO"] == "":
        continue  # Saltiamo le righe senza numeri di lotto
    
    lotti_map.setdefault(key, []).extend(str(row["LOTTO"]).split('-'))

# Assegnazione lotti per righe con imponibile = 0
for idx, row in df[df["Somma di IMPONIBILE EUR"] == 0].iterrows():
    key = (row["ORDINE #"], row["SKU COMPLETO"], row["DOC TYPE"])
    qty = int(float(row["Somma di QTY PRODOTTO"].replace(',', '.')))

    if key in lotti_map and qty > 0:
        lotti_da_assegnare = lotti_map[key][:qty]

        if lotti_da_assegnare:
            df.at[idx, "LOTTO"] = '-'.join(lotti_da_assegnare)
            lotti_map[key] = lotti_map[key][qty:]

            # Se dopo l'assegnazione la cella "LOTTO" è vuota, mantieni almeno un numero
            if df.at[idx, "LOTTO"] == "":
                df.at[idx, "LOTTO"] = lotti_da_assegnare[-1]  # Ultimo lotto disponibile

            # Aggiorniamo la riga con imponibile > 0
            mask = (df["ORDINE #"] == row["ORDINE #"]) & (df["SKU COMPLETO"] == row["SKU COMPLETO"]) & (df["DOC TYPE"] == row["DOC TYPE"]) & (df["Somma di IMPONIBILE EUR"] > 0)
            df.loc[mask, "LOTTO"] = '-'.join(lotti_map[key]) if lotti_map[key] else df.loc[mask, "LOTTO"].fillna(lotti_da_assegnare[-1])


# **Ripristiniamo i valori originali dell'Imponibile senza modificarli**
df["Somma di IMPONIBILE EUR"] = df_original["Somma di IMPONIBILE EUR"]

# Debug: controlliamo i dati prima di esportare
print("\U0001F50D Controllo finale (Prime 5 righe)")
print(df[["ORDINE #", "SKU COMPLETO", "DOC TYPE", "Somma di IMPONIBILE EUR", "LOTTO"]].head())

# **Esportiamo in Excel con formato corretto**
df.to_excel(output_path, index=False, engine='openpyxl')

print("✅ Elaborazione completata. File salvato in:", output_path)
