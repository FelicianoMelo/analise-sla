# Projeto Análise SLA

Este projeto consiste em um conjunto de scripts em Python desenvolvidos para automatizar o processamento, análise e geração de relatórios de Service Level Agreement (SLA) a partir de planilhas Excel (exportadas de sistemas de chamados como ServiceNow, por exemplo).

O sistema cruza dados de incidentes e tarefas (SCTASK), aplica filtros baseados em configurações externas, calcula o desempenho de SLA (categorizando em diferentes níveis de alerta) e gera painéis consolidados ("Visão Geral").

## Estrutura do Projeto

### Scripts Principais

- **`gerar_planilhas_sla.py`**: É o script principal (core) do projeto. Ele lê as configurações de `configuracao.json`, carrega as planilhas base, filtra os dados, cruza informações, aplica estilos, calcula os indicadores de SLA (Ótimo, Médio, Atenção, Fora do SLA) e gera os arquivos finais de saída, incluindo o painel de Visão Geral.
- **`gerar_planilhas.py`**: Versão alternativa ou anterior do script principal de geração de relatórios.
- **`configuracao.json`**: Arquivo de configuração em formato JSON que define as regras de cruzamento de dados. Determina quais arquivos de entrada ler, se os dados devem ser mesclados ou cruzados, os filtros a aplicar (ex: "Clearance", "CSATH") e os nomes dos arquivos e abas de saída.
- **`app_sla.py`**: Uma interface gráfica simples construída com `customtkinter` que permite ao usuário selecionar um diretório e listar todos os arquivos `.xlsx` contidos nele.
- **`ajusta_painel_sla.py`**: Módulo utilitário utilizado para desenhar, formatar e estilizar o painel de "Visão Geral" gerado nas planilhas de totais. Utiliza `openpyxl` para criar blocos coloridos e bordas para o cabeçalho, abertos, encerrados e totais de SLA.
- **`relat.py`**: Módulo que contém a classe `ExcelReport`. Facilita a exportação de DataFrames do Pandas para o Excel e a aplicação padronizada de formatações em blocos (títulos, dados e totais) usando o `openpyxl`.
- **`enviar_email_sla.py`**: Script de automação de e-mails. Ele lê uma planilha chamada `emails.xlsx` e, para cada linha, envia um e-mail com os relatórios em anexo utilizando um servidor SMTP (Office 365).
- **`BaixarPlanilhasEPDSM.py`**: Um script simples utilizando `selenium` projetado para automatizar o download das planilhas base a partir de um portal web.
- **Outros utilitários**: O projeto também conta com scripts utilitários e de teste como `ExcelReportHeaderOnly.py`, `ExemplosColorCelula.py` e vários scripts temporários (`teste.py`, `temp.py`, etc.).

### Arquivos de Log e Saída
- **`analise_sla.log`**: Arquivo de log gerado pelo módulo `logging` durante a execução do processo principal, registrando os passos realizados e eventuais erros.
- **Arquivos Excel gerados (`.xlsx`)**: Arquivos como `YYYYMMDD_Total.xlsx`, `YYYYMMDD_DOM_aberto.xlsx`, gerados como resultado do processamento.

## Fluxo de Processamento de SLA

1. **Leitura da Configuração**: O script principal abre e itera sobre a lista definida no `configuracao.json`.
2. **Processamento Baseado em Regras**:
   - Para cruzamentos (`JuntaArq: false`), o script une dados de duas planilhas (por exemplo, incidentes e logs de SLA).
   - Para junções simples (`JuntaArq: true`), o script apenas filtra a planilha base pelos critérios definidos.
3. **Cálculo de SLA**: É calculada a distribuição dos chamados baseada no percentual de SLA:
   - **Verde (Ótimo)**: Até 45%
   - **Amarelo (Em Andamento/Médio)**: 46% a 50%
   - **Laranja (Em Atenção)**: 51% a 99%
   - **Vermelho (Fora do SLA)**: Acima de 99%
4. **Resumos e Agrupamentos**: O sistema agrupa os dados filtrados por métricas chave como "IC afetado", "Resolvido por" ou "Atribuição a" e escreve o resumo na aba de "SLA".
5. **Geração do Painel de Visão Geral**: Ao final de todos os ciclos do json, é gerada uma planilha "Total" (`YYYYMMDD_Total.xlsx`) contendo um dashboard formatado com os números consolidados por grupo (ex: Transporte Doméstico, Clearance Internacional).

## Tecnologias e Bibliotecas Utilizadas
- **Python 3**
- **Pandas**: Para a leitura estruturada, cruzamento, filtragem e mesclagem (`merge`) dos dados originais.
- **OpenPyXL**: Para manipulação fina de estilos no Excel (aplicação de cores nas células, bordas, fontes, alinhamentos) sem perder a formatação base.
- **CustomTkinter**: Para a interface gráfica base do projeto (`app_sla.py`).
- **Selenium**: Automação de browser para extração de dados.
- **smtplib / email**: Para a leitura e disparo automático dos e-mails com relatórios.

## Como Executar

Para gerar os relatórios completos a partir da configuração padrão:
```bash
python gerar_planilhas_sla.py
```

Para listar as planilhas existentes em uma pasta via interface:
```bash
python app_sla.py
```

Para automatizar o disparo de e-mails com os resultados (após popular `emails.xlsx`):
```bash
python enviar_email_sla.py
```
