import pandas as pd
import win32com.client as win32
from datetime import datetime

# Lê a aba "Viagens"
df_viagens = pd.read_excel(
    "RupturaExpedição.xlsm",
    sheet_name='Viagens'
)

# Filtra linhas com valor na coluna 'Viagem'
df_viagens = df_viagens[df_viagens['Viagem'].notna()]

# Converte colunas numéricas para inteiro
if df_viagens['Viagem'].dtype == 'float':
    df_viagens['Viagem'] = df_viagens['Viagem'].astype(int)

if df_viagens['Turno'].dtype == 'float':
    df_viagens['Turno'] = df_viagens['Turno'].fillna(0).astype(int)

colunas_totais = ['DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
for col in colunas_totais:
    if col in df_viagens.columns and df_viagens[col].dtype == 'float':
        df_viagens[col] = df_viagens[col].fillna(0).astype(int)

# Reordena as colunas
ordem_colunas = ['Data', 'Viagem', 'Placa', 'Status', 'Cluster', 'Turno',
                 'DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
df_viagens = df_viagens[ordem_colunas]

# Cria linha de totais
linha_total = {col: df_viagens[col].sum() if col in colunas_totais else '' for col in df_viagens.columns}
linha_total['Cluster'] = 'TOTAL'

# Adiciona a linha de total
df_viagens = pd.concat([df_viagens, pd.DataFrame([linha_total])], ignore_index=True)

# Estilo visual com ajustes solicitados
estilo = """
<style>
    table {
        border-collapse: collapse;
        width: 100%;
        text-align: center;
        font-family: Arial, sans-serif;
        table-layout: auto;
    }
    thead th {
        background-color: #d3d3d3;
        border: 1px solid black;
        padding: 6px;
        white-space: nowrap;
    }
    tbody td {
        border: 1px solid black;
        padding: 6px;
        white-space: nowrap;
    }
    tbody tr:nth-child(even) {
        background-color: #f9f9f9;
    }
    tbody tr:nth-child(odd) {
        background-color: #ffffff;
    }
    tbody tr:last-child {
        background-color: #d3d3d3;
        font-weight: bold;
    }
</style>
"""

# Transforma em HTML
html_tabela = df_viagens.to_html(index=False, border=0, justify='center')

# Corpo do e-mail
corpo_email = estilo + f"""
<p>Segue relação de viagens:</p>
{html_tabela}
"""

# Envia via Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "guilherme.dcastro@americanas.io"
mail.CC = ""
mail.Subject = f"Acompanhamento Expedição - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
mail.HTMLBody = corpo_email
mail.Send()

print("E-mail enviado com sucesso!")
