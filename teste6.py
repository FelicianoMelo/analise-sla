import pandas as pd
import numpy as np

# 1. Ler os dados da planilha
arquivo1 = f'/Users/famj/projetos/analise_sla/SLA_por_Grupos - Fernando Pantuffi(incident).xlsx'
df = pd.read_excel(arquivo1)

# 2. Criar a tabela dinâmica com 2 condições (ex: Região e Produto)
# O parâmetro index recebe as 2 condições de agrupamento
tabela_dinamica = pd.pivot_table(
    df,
    values='Aberto(a)',  # Coluna que será agregada (contada)
    index=['IC afetado',df['Aberto(a)'].strftime("%Y%m%d")], # Colunas que serão usadas como condições de agrupamento
    aggfunc='count',  # Contar o número de ocorrências
    margins=True,       # Adiciona a linha de total geral   
    margins_name='Total Geral' # Nome da linha de total geral
)

# 3. Exportar o resultado para um ficheiro Excel (.xlsx)
tabela_dinamica.to_excel('tabela_2_condicoes.xlsx')
print("Tabela dinâmica criada com sucesso!")

# import pandas as pd
# import numpy as np

# # 1. Criar dados de exemplo
# dados = {
#     'Regiao': ['Norte', 'Norte', 'Sul', 'Sul', 'Norte', 'Sul'],
#     'Produto': ['Caderno', 'Caneta', 'Caderno', 'Caneta', 'Caderno', 'Caneta'],
#     'Vendas': [150, 200, 300, 100, 50, 120]
# }
# df = pd.DataFrame(dados)

# # 2. Criar a tabela dinâmica com 2 condições (ex: Região e Produto)
# # O parâmetro index recebe as 2 condições de agrupamento
# tabela_dinamica = pd.pivot_table(
#     df,
#     values='Vendas',  # Coluna que será agregada (contada)
#     index=['Produto'],
#     aggfunc='count',
#     margins=True,          # Adiciona a linha de total geral
#     margins_name='Total Geral'
# )

# # 3. Exportar o resultado para um ficheiro Excel (.xlsx)
# tabela_dinamica.to_excel('tabela_2_condicoes.xlsx')
# print("Tabela dinâmica criada com sucesso!")
