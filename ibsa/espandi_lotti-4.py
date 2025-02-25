import pandas as pd

# Percorsi dei file
input_file = r"IBSA_Report fatture output - 3.xlsx"
output_file = r"IBSA_Report fatture output - 4.xlsx"

# Carica il file Excel
df = pd.read_excel(input_file, dtype=str)  # Carica tutto come stringa per evitare problemi iniziali

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

# Converte i valori numerici (mantiene il segno e sostituisce ',' con '.')
for col in colonne_da_dividere + [colonna_divisore]:
    df[col] = df[col].str.replace(',', '.', regex=True)  # Sostituisce ',' con '.'
    df[col] = pd.to_numeric(df[col], errors='coerce')  # Converte in float, mantenendo segni negativi

# Divisione senza alterare il segno
def calcola_divisione(row, col, divisor_col):
    num = row[col]
    denom = row[divisor_col]
    
    # Se il denominatore è 0 o non valido, restituisce 'Non valido'
    if pd.isna(num) or pd.isna(denom) or denom == 0:
        return "Non valido"
    
    # Altrimenti esegue la divisione mantenendo il segno
    result = num / denom
    
    # Mantieni il segno in modo che il risultato sia negativo
    if num < 0 and denom < 0:
        return -abs(result)  # entrambi negativi, risulta negativo
    elif num < 0 or denom < 0:
        return -abs(result)  # uno dei due è negativo, risulta negativo
    return abs(result)  # entrambi positivi, risultato positivo

# Applica la divisione alle colonne
for col in colonne_da_dividere:
    df[col] = df.apply(lambda row: calcola_divisione(row, col, colonna_divisore), axis=1)

# Salva il risultato in un nuovo file Excel
df.to_excel(output_file, index=False)

print(f"File salvato con successo come {output_file}")
