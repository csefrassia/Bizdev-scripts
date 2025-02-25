import pandas as pd

# Funzione per caricamento e pulizia dei dati
def load_and_clean(file_path):
    try:
        df = pd.read_csv(file_path, dtype=str) if file_path.endswith('.csv') else pd.read_excel(file_path, dtype=str)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        print(f"Errore durante il caricamento del file: {e}")
        return pd.DataFrame()

# Configurazione dei percorsi
INPUT_CSV = "IBSA_Report fatture.csv"
OUTPUT_XLSX = "IBSA_Report fatture output.csv"

# Caricamento dei dati originali
df_original = load_and_clean(INPUT_CSV)
df = df_original.copy()

# Formattazione delle colonne numeriche
def format_numeric_columns(df):
    try:
        numeric_cols = ["Somma di IMPONIBILE EUR", "Somma di QTY PRODOTTO"]
        for col in numeric_cols:
            df[col] = df[col].str.replace(',', '.', regex=True).astype(float)
    except Exception as e:
        print(f"Errore durante la formattazione numerica: {e}")

# Gestione dei lotti
def manage_lots(df, df_original):
    lotti_map = {}
    positive_mask = df["Somma di IMPONIBILE EUR"] > 0
    zero_mask = df["Somma di IMPONIBILE EUR"] == 0

    # Costruisci mappa dei lotti per righe con imponibile > 0
    for _, row in df[positive_mask].iterrows():
        key = (row["ORDINE #"], row["SKU COMPLETO"], row["DOC TYPE"])
        if pd.isna(row["LOTTO"]):
            continue
        lotti = str(row["LOTTO"]).split('-')
        lotti_map[key] = lotti_map.get(key, []) + lotti

    # Assegna lotti per righe con imponibile = 0
    for idx, row in df[zero_mask].iterrows():
        key = (row["ORDINE #"], row["SKU COMPLETO"], row["DOC TYPE"])
        qty = int(row["Somma di QTY PRODOTTO"])
        if key not in lotti_map or qty <= 0:
            continue
        
        available_lots = lotti_map[key]
        if not available_lots:
            continue
        
        assigned_lots = available_lots[:qty]
        remaining_lots = available_lots[qty:]
        
        # Aggiorna la riga corrente
        new_lotto = '-'.join(assigned_lots) if assigned_lots else available_lots[-1]
        df.at[idx, "LOTTO"] = new_lotto
        
        # Aggiorna la mappa
        lotti_map[key] = remaining_lots

    return df

# Esporta i dati
def export_data(df, output_path):
    try:
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"✅ Elaborazione completata. File salvato in: {output_path}")
    except Exception as e:
        print(f"Errore durante l'esportazione: {e}")

# Flusso principale delle operazioni
format_numeric_columns(df)
df = manage_lots(df, df_original)

# Ripristina formati numerici originali (stringhe con virgola)
df["Somma di IMPONIBILE EUR"] = df_original["Somma di IMPONIBILE EUR"]

# Esporta il file
export_data(df, OUTPUT_XLSX)