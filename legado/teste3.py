import openpyxl
import pandas as pd

# 1. Carrega o arquivo Excel
wb = openpyxl.load_workbook('/Users/famj/projetos/analise_sla/20260804_INT_aberto.xlsx', data_only=True)
sheet = wb.active

dados = []

# 2. Percorre as linhas (exemplo: da linha 2 até a 100)
for row in range(2, 101):
    celula_cor = sheet[f'A{row}']    # Coluna onde está a cor de fundo
    celula_valor = sheet[f'B{row}']  # Coluna onde está o valor a ser somado
    
    # Extrai o código Hex da cor de preenchimento
    if celula_cor.fill and celula_cor.fill.start_color and celula_cor.fill.start_color.rgb:
        hex_cor = str(celula_cor.fill.start_color.rgb)
        
        # Remove o canal Alfa (transparência 'FF') se o openpyxl retornar 8 caracteres (ex: FFC6EFCE -> C6EFCE)
        if len(hex_cor) == 8 and hex_cor.startswith('FF'):
            hex_cor = hex_cor[2:]
    else:
        hex_cor = 'SEM_COR'
        
    dados.append({
        'cor_hex': hex_cor,
        'valor': celula_valor.value or 0
    })

# 3. Cria o DataFrame do Pandas
df = pd.DataFrame(dados)

# 4. Agrupa por cor e soma a coluna 'valor'
resultado = df.groupby('cor_hex')['valor'].sum().reset_index()

print(resultado)