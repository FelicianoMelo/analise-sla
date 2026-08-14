import pandas as pd
import json
import datetime
import os

data_atual = datetime.datetime.now().strftime("%Y%m%d") # Data atual para salvar o arquivo com um nome único
pasta_trabalho = "/Users/famj/projetos/analise_sla/" # Pasta onde os arquivos serão salvos
os.chdir(pasta_trabalho)
df_plan3 = []


def cruzar_planilhas(arquivo1, arquivo2, arquivo_saida):
    # 1. Ler os dados das planilhas para DataFrames
    df_plan1 = pd.read_excel(arquivo1)
    df_plan2 = pd.read_excel(arquivo2)
  
    # 2. Renomear a coluna 0 para 'Chave' (caso o nome seja diferente)
    df_plan1.rename(columns={df_plan1.columns[0]: 'Chave'}, inplace=True)
    df_plan2.rename(columns={df_plan2.columns[0]: 'Chave'}, inplace=True)

    # 3. Selecionar apenas as colunas desejadas da planilha 2 Número e 6
    df_plan2 = df_plan2.iloc[:, [6,0]  ] # Seleciona a coluna 0 (Chave) e a coluna 6 da planilha 2
    
    # 4. Mesclar (Unir) os dados com base na coluna 0 (Chave)
    # Utilizando 'inner' para apenas onde há correspondência nas duas planilhas
    # Altere para 'outer', 'left' ou 'right' se desejar manter todos os dados independente de correspondência
    df_plan3 = pd.merge(df_plan1, df_plan2, on='Chave', how='left')  # Mantendo todos os dados da planilha 1 e apenas os correspondentes da planilha 2
    
    # 5. Salvar o resultado em uma nova planilha    
    df_plan3.to_excel(arquivo_saida, index=False)
    
    # 6. Exibir mensagem de sucesso 
    print(f"Arquivo '{arquivo_saida}' criado com sucesso!") 

def planilhas_encerrados(arquivo1, arquivo_saida, filtro):
    # 1. Ler os dados da pplanilha 
    df_plan1 = pd.read_excel(arquivo1)
    
    # 2. Filtrar os dados com base no valor da coluna 5
    df_filtrado = df_plan1[df_plan1.iloc[:, 5] == filtro]  # Filtra onde a coluna 5 é igual ao valor do filtro
    
    # 3. Salvar o resultado em uma nova planilha
    df_filtrado.to_excel(arquivo_saida, index=False)    
   
    # 4. Exibir mensagem de sucesso
    print(f"Arquivo '{arquivo_saida}' criado com sucesso!") 


def processar_configuracoes():
    #ler o arquivo json
    with open('configuracao.json', 'r', encoding="utf-8") as f:
        configuracoes = json.load(f) 

    # criar uma nova planilha com base nas configurações do arquivo JSON
    for i in range(len(configuracoes)):

        # Defina os nomes dos arquivos de entrada e saída com base nas configurações
        arquivo1 = configuracoes[i]['ArquivoEntrada1']
        arquivo2 = configuracoes[i]['ArquivoEntrada2']
        arquivo_saida = f"{data_atual}_{configuracoes[i]['arquivoSaida']}"
        filtro = configuracoes[i]['Filtro']

        # Chame a função para cruzar as planilhas
        if arquivo2 != "":
            #
            cruzar_planilhas(arquivo1, arquivo2, arquivo_saida) 
        else:
            # Chame a função para filtrar os dados da planilha 1 com base no filtro
            planilhas_encerrados(arquivo1, arquivo_saida, filtro)       

def main():
    processar_configuracoes()     

if __name__ == "__main__":
    main()