import pandas as pd

# Lendo o arquivo Excel
df_plan1 = pd.read_excel(R'/Users/famj/projetos/analise_sla/20260729_INT_aberto.xlsx')  # Substitua 'arquivo1.xlsx' pelo caminho do seu arquivo
df_plan1.sheet_name=""
# Função para aplicar cores com base nos valores da coluna "% SLA Resolution"
def color_sla(val):
    if val <= 45.00:
        color = 'green'
    elif 45.01 <= val <= 50.00:
        color = 'yellow'
    elif 50.01 <= val <= 90.00:
        color = 'orange'
    elif val >= 100:
        color = 'red'
    else:
        color = 'white'  # Caso não se encaixe em nenhuma condição
    return f'background-color: {color}'

# Aplicando a formatação condicional na coluna "% SLA Resolution"
styled_df = df_plan1.style.applymap(color_sla, subset=['% SLA Resolution'])

# Exibindo o DataFrame estilizado
#styled_df