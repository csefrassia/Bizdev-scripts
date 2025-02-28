import pandas as pd
import argparse

def main(input_csv_path, output_excel_path):
    # Caricamento del file CSV senza modificare i tipi
    df = pd.read_csv(input_csv_path, dtype=str)

    # Pulizia nomi colonne (rimuove spazi extra)
    df.columns = df.columns.str.strip()

    # Conversione delle colonne numeriche
    df["Somma di IMPONIBILE EUR"] = df["Somma di IMPONIBILE EUR"].str.replace(',', '.').astype(float)
    df["Somma di QTY PRODOTTO"] = df["Somma di QTY PRODOTTO"].str.replace(',', '.').astype(float)

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
        qty = int(float(row["Somma di QTY PRODOTTO"]))

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

    df.columns = df.columns.str.strip()

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

    # Espansione dei lotti
    new_rows = []
    # Itera sulle righe del dataframe
    for index, row in df.iterrows():
        # Estrai i numeri di lotto
        lot_numbers = str(row['LOTTO']).split('-')
        
        # Per ogni numero di lotto, crea una nuova riga con tutte le informazioni dell'ordine
        for lot in lot_numbers:
            new_row = row.to_dict()  # Crea una copia del dizionario con tutte le colonne
            new_row['LOTTO'] = lot  # Aggiorna il numero di lotto
            new_rows.append(new_row)

    df_expanded = pd.DataFrame(new_rows)

    # Colonne da dividere
    colonne_da_dividere = [
        "Somma di PREZZO ORIGINALE",
        "Somma di FATTURATO",
        "Somma di FATTURATO EUR",
        "Somma di PREZZO EUR IVATO",
        "Somma di SCONTO EUR IVATO",
        "Somma di IMPONIBILE EUR"
    ]

    # Nome della colonna divisore
    colonna_divisore = "Somma di QTY PRODOTTO"

    # Converte i valori numerici
    for col in colonne_da_dividere + [colonna_divisore]:
        df_expanded[col] = df_expanded[col].astype(str).str.replace(',', '.', regex=True)
        df_expanded[col] = pd.to_numeric(df_expanded[col], errors='coerce')

    # Divisione senza alterare il segno
    def calcola_divisione(row, col, divisor_col):
        num = row[col]
        denom = row[divisor_col]
        
        if pd.isna(num) or pd.isna(denom) or denom == 0:
            return "Non valido"
        
        result = num / denom
        
        if num < 0 and denom < 0:
            return -abs(result)
        elif num < 0 or denom < 0:
            return -abs(result)
        return abs(result)

    # Applica la divisione alle colonne
    for col in colonne_da_dividere:
        df_expanded[col] = df_expanded.apply(lambda row: calcola_divisione(row, col, colonna_divisore), axis=1)

    # Modifica i valori della colonna `Somma di QTY PRODOTTO` per essere sempre 1 mantenendo il segno
    df_expanded["Somma di QTY PRODOTTO"] = df_expanded["Somma di QTY PRODOTTO"].apply(lambda x: 1 if x > 0 else -1)

    # Salva il risultato in un nuovo file Excel
    # df_expanded.to_excel(output_excel_path, index=False, engine='openpyxl')

    print(f"✅ Elaborazione completata. File salvato in: {output_excel_path}")

    grouped_df = df_expanded.groupby(["DOC TYPE", "ORDINE #", "LOTTO"]).agg(
        totale_quantità=("Somma di QTY PRODOTTO", "sum"),
        prezzo_medio=("prezzo_unitario", "mean")
    ).reset_index()


     # Salva il risultato in un nuovo file Excel
    grouped_df.to_excel(output_excel_path, index=False, engine='openpyxl')

    print(f"✅ Elaborazione completata. File salvato in: {output_excel_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Processa un file CSV e salva il risultato in un file Excel.")
    parser.add_argument("input_csv_path", help="Percorso del file CSV di input")
    parser.add_argument("output_excel_path", help="Percorso del file Excel di output")
    args = parser.parse_args()

    main(args.input_csv_path, args.output_excel_path)
     
    # python processa_fatture.py "IBSA_Report fatture.csv" "IBSA_Report fatture_Seba2.xlsx"
   