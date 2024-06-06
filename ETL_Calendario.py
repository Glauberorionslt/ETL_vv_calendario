import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def criar_tabela_calendario():
    # Definir a data inicial e a data final
    data_inicial = datetime(2020, 1, 1)
    data_final = datetime.now()

    # Gerar as datas no intervalo desejado
    datas = pd.date_range(start=data_inicial, end=data_final, freq='D')

    # Criar DataFrame com as datas e informações adicionais
    df_calendario = pd.DataFrame({
        'Data': datas,
        'Nome do Mês': datas.month_name(),
        'Mês (Inteiro)': datas.month,
        'Ano': datas.year
    })

    # Caminho do arquivo de saída
    output_directory = r"C:\Users\2160011883.VIA\OneDrive - Via Varejo S.A\Bases - DB_MANUTENCAO"
    output_file = os.path.join(output_directory, "d_calendario.xlsx")

    # Verificar se o arquivo de saída já existe e remover se necessário
    if os.path.exists(output_file):
        os.remove(output_file)
        print(f"Arquivo existente '{output_file}' removido.")

    # Criar um novo Workbook do openpyxl
    wb = Workbook()
    ws = wb.active

    # Copiar os dados do DataFrame para a planilha do Workbook
    for row in dataframe_to_rows(df_calendario, index=False, header=True):
        ws.append(row)

    # Salvar o arquivo Excel
    wb.save(output_file)
    print(f"Tabela de calendário salva com sucesso em '{output_file}'")

# Chamar a função para criar a tabela de calendário
criar_tabela_calendario()
