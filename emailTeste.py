import pandas as pd
import win32com.client as win32
from datetime import datetime

# Lê a aba "Viagens"
df_viagens = pd.read_excel(
    "RupturaExpedição.xlsm",
    sheet_name='Viagens'
)

# Filtra linhas com valor na coluna 'Placa'
df_viagens = df_viagens[df_viagens['Placa'].notna()]

# Padroniza a coluna 'Placa' (opcional, mas útil)
df_viagens['Placa'] = df_viagens['Placa'].astype(str).str.upper().str.strip()

# Converte colunas numéricas para inteiro (se aplicável)
if 'Viagem' in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens['Viagem']):
    df_viagens['Viagem'] = df_viagens['Viagem'].astype(int)

if 'Turno' in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens['Turno']):
    df_viagens['Turno'] = df_viagens['Turno'].fillna(0).astype(int)

colunas_totais = ['DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
for col in colunas_totais:
    if col in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens[col]):
        df_viagens[col] = df_viagens[col].fillna(0).astype(int)

# Reordena as colunas na ordem desejada
ordem_colunas = ['Data', 'Viagem', 'Placa', 'Status', 'Cluster', 'Turno',
                 'DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
df_viagens = df_viagens[[col for col in ordem_colunas if col in df_viagens.columns]]

# Cria uma linha de totais com preenchimento adequado
linha_total = []
for col in df_viagens.columns:
    if col in colunas_totais:
        total = df_viagens[col].sum()
        linha_total.append(total)
    elif col == 'Cluster':
        linha_total.append('TOTAL')
    else:
        linha_total.append('')

# Adiciona a linha ao DataFrame
df_total = pd.DataFrame([linha_total], columns=df_viagens.columns)
df_viagens = pd.concat([df_viagens, df_total], ignore_index=True)

# Estilo visual da tabela HTML
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

# Converte o DataFrame em tabela HTML
html_tabela = df_viagens.to_html(index=False, border=0, justify='center')

# Corpo do e-mail
corpo_email = estilo + f"""
<p>Segue relação de viagens (filtradas por Placa):</p>
{html_tabela}
"""

# Envia via Outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "lucas.dimoura@americanas.io;lynconl.souza@americanas.io;ricafreitas@americanas.io;mailla.amorim@americanas.io;raphael.brandao@americanas.io;expedicao.cdmg@americanas.io;guilherme.dcastro@americanas.io;hajykaar.souza@americanas.io;luiz.asilva@americanas.io;braulio.terra@americanas.io"
mail.CC = ""
mail.Subject = f"Acompanhamento Expedição - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
mail.HTMLBody = corpo_email
mail.Send()

print("E-mail enviado com sucesso!")
