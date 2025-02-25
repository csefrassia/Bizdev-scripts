import pandas as pd

# Percorsi dei file
input_file = r"IBSA_Report fatture output - 2.xlsx"
output_file = r"IBSA_Report fatture output - 3.xlsx"

# Carica il file CSV
df = pd.read_excel(input_file, dtype=str)

# Assicurati che i nomi delle colonne siano privi di spazi indesiderati
df.columns = df.columns.str.strip()

# Crea una lista vuota per memorizzare i nuovi dati
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

# Crea un nuovo dataframe con le righe modificate
new_df = pd.DataFrame(new_rows)

# Salva il risultato in un file Excel
new_df.to_excel(output_file, index=False)

print(f"File salvato con successo come {output_file}")