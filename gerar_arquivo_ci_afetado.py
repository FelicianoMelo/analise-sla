import pandas as pd
import json

# Ler Excel
df = pd.read_excel("arquivo.xlsx")

# Ler coluna CI e remover duplicidades
cis = (
    df["CI"]
    .dropna()              # remove vazios
    .drop_duplicates()     # remove duplicados
    .tolist()
)

# Montar estrutura JSON
dados = [
    {
        "CI": ci,
        "otimo": 0,
        "medio": 0,
        "atencao": 0,
        "fora": 0,
        "total": 0
    }
    for ci in cis
]

# Salvar arquivo JSON
with open("ci.json", "w", encoding="utf-8") as f:
    json.dump(dados, f, ensure_ascii=False, indent=4)

print(f"Arquivo criado com {len(dados)} registros.")