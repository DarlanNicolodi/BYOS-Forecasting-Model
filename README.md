# BYOS-Forecasting-Model
BYOS Forecasting Model for MIT Capstone Project

import pandas as pd
from datetime import datetime, timedelta
import os
import win32com.client as win32


# Função para prever a próxima coleta
def prever_proxima_coleta(local, datas, data_previsao):
    # Garantir que as datas estão no formato correto (datetime) e remover inválidas
    datas = pd.to_datetime(datas, errors='coerce').dropna()
    
    if len(datas) == 0:
        print(f"Sem dados válidos de coletas para {local}.")
        return None, None, None, 0, None, 'N/A'

    datas = datas.sort_values()
    data_ultima_coleta = datas[-1].strftime('%Y-%m-%d')
    dias_desde_ultima_coleta = (data_previsao - pd.to_datetime(data_ultima_coleta)).days

    return None, None, None, len(datas), None, data_ultima_coleta, dias_desde_ultima_coleta

# Definir a data de previsão
data_previsao = datetime(2024, 10, 1)

# Carregar a planilha Excel
df = pd.read_excel('C:/Users/cs164112/Desktop/Byos Forecast/RECOLECCION.xlsx')

if 'Local' not in df.columns or 'Data Coleta' not in df.columns or 'Volume (L)' not in df.columns or 'Data Sugestão Cliente' not in df.columns:
    raise ValueError("As colunas 'Local', 'Data Coleta', 'Volume (L)' e 'Data Sugestão Cliente' são obrigatórias no arquivo Excel.")

# Calcular a capacidade máxima de cada cliente como o maior volume coletado
df['Capacidade Máxima (L)'] = df.groupby('Local')['Volume (L)'].transform('max')

# Calcular o volume total coletado de cada cliente
df['Volume Total Cliente'] = df.groupby('Local')['Volume (L)'].transform('sum')

# Calcular o volume da primeira coleta de cada cliente
df['Volume Primeira Coleta'] = df.groupby('Local')['Volume (L)'].transform('first')

# Calcular o volume -1St
df['Volume -1St'] = df['Volume Total Cliente'] - df['Volume Primeira Coleta']

# Calcular os dias entre a primeira e a última coleta de cada cliente
df['Dias Entre Coletas'] = (df.groupby('Local')['Data Coleta'].transform('max') - df.groupby('Local')['Data Coleta'].transform('min')).dt.days

# Calcular DAUCOP (média diária de geração de óleo usado)
df['DAUCOP'] = (df['Volume -1St'] / df['Dias Entre Coletas']).round(2)

previsoes = []

# Agrupar os dados por local e prever as próximas coletas até 31 de dezembro de 2024
for local in df['Local'].unique():
    dados_local = df[df['Local'] == local]
    datas_coletas = dados_local['Data Coleta'].tolist()
    capacidade_maxima = dados_local['Capacidade Máxima (L)'].iloc[0]  # Assumindo que a capacidade máxima é constante para cada local
    volume_total_cliente = dados_local['Volume Total Cliente'].iloc[0]
    volume_primeira_coleta = dados_local['Volume Primeira Coleta'].iloc[0]
    volume_menos_primeira = volume_total_cliente - volume_primeira_coleta
    dias_entre_coletas = dados_local['Dias Entre Coletas'].iloc[0]
    daucop = dados_local['DAUCOP'].iloc[0]
    
    # Verificar se 'Data Sugestão Cliente' não está vazia
    if not dados_local['Data Sugestão Cliente'].empty:
        data_sugestao_cliente_texto = dados_local['Data Sugestão Cliente'].iloc[-1].strftime('%Y-%m-%d') if not pd.isna(dados_local['Data Sugestão Cliente'].iloc[-1]) else 'N/A'
    else:
        data_sugestao_cliente_texto = "N/A"  # Ou outro valor padrão adequado

    _, _, _, numero_de_coletas, _, data_ultima_coleta, dias_desde_ultima_coleta = prever_proxima_coleta(local, datas_coletas, data_previsao)
    
    # Calcular "Dias para a Próxima Coleta vs Última"
    dias_para_proxima_coleta_vs_ultima = round((capacidade_maxima / daucop), 2) if daucop > 0 else 0
    data_proxima_coleta = pd.to_datetime(data_ultima_coleta) + timedelta(days=int(dias_para_proxima_coleta_vs_ultima))
    
    while data_proxima_coleta <= datetime(2024, 12, 31):
        if dias_para_proxima_coleta_vs_ultima <= 0:
            print(f"Erro: 'Days Until Next Collection' é inválido para o local {local}.")
            break  # Sai do loop para evitar looping infinito

        previsoes.append((
            local,  # Client
            data_sugestao_cliente_texto,  # Suggested Collection Date
            volume_total_cliente,  # Total Volume Collected
            volume_menos_primeira,  # Volume -1St
            dias_entre_coletas,  # Days Between Collections
            daucop,  # Daily Average Used Oil Collection (DAUCOP)
            capacidade_maxima,  # Maximum Storage Volume (L)
            dias_para_proxima_coleta_vs_ultima,  # Days Until Next Collection
            data_ultima_coleta,  # Last Collection Date
            data_proxima_coleta.strftime('%Y-%m-%d'),  # Next Collection Date
            numero_de_coletas,  # Number of Previous Collections
            dias_desde_ultima_coleta  # Days Since Last Collection
        ))
        data_ultima_coleta = data_proxima_coleta.strftime('%Y-%m-%d')
        data_proxima_coleta += timedelta(days=int(dias_para_proxima_coleta_vs_ultima))
        numero_de_coletas += 1


# Criar o DataFrame com as previsões
df_previsoes = pd.DataFrame(previsoes, columns=['Client', 'Suggested Collection Date', 'Total Volume Collected', 'Volume -1St', 'Days Between Collections', 'Daily Average Used Oil Collection (DAUCOP)', 'Maximum Storage Volume (L)', 'Days Until Next Collection', 'Last Collection Date', 'Next Collection Date', 'Number of Previous Collections', 'Days Since Last Collection'])

# Salvar as previsões em um arquivo Excel
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
caminho_area_de_trabalho = os.path.join('C:/Users/cs164112/Desktop/Byos Forecast', f'collection_forecasts_{timestamp}.xlsx')
df_previsoes.to_excel(caminho_area_de_trabalho, index=False)

# Abrir o Excel automaticamente e configurar a planilha
oApp = win32.Dispatch('Excel.Application')
oApp.Visible = True  # Tornar o Excel visível ao abrir o arquivo
workbook = oApp.Workbooks.Open(caminho_area_de_trabalho)

sheet = workbook.Sheets(1)

# Ajustar a largura das colunas automaticamente
sheet.Columns.AutoFit()

# Ajustar a largura da coluna A para 28
sheet.Columns('A').ColumnWidth = 28

# Aplicar filtros na linha de cabeçalho
sheet.Rows(1).AutoFilter()

# Congelar a primeira linha
sheet.Application.ActiveWindow.SplitRow = 1
sheet.Application.ActiveWindow.FreezePanes = True

# Centralizar o texto das colunas B até J
sheet.Range("B:J").HorizontalAlignment = -4108  # -4108 corresponde ao valor de alinhamento central no Excel

# Selecionar a célula A1 para exibição
sheet.Range("A1").Select()

# Salvar e manter o Excel aberto para visualização
workbook.Save()
oApp.DisplayAlerts = False  # Desativar alertas

print(f"Arquivo Excel com as previsões salvo na área de trabalho: {caminho_area_de_trabalho}")

