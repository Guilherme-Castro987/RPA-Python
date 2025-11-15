import pandas as pd
import win32com.client as win32
from datetime import datetime

# Caminho do arquivo
arquivo = "RupturaExpedição.xlsm"

# ==============================
# 🔹 Leitura da aba "Viagens"
# ==============================
df_viagens = pd.read_excel(arquivo, sheet_name='Viagens')

# Filtra linhas com valor na coluna 'Placa'
df_viagens = df_viagens[df_viagens['Placa'].notna()]

# Padroniza 'Placa'
df_viagens['Placa'] = df_viagens['Placa'].astype(str).str.upper().str.strip()

# Converte colunas numéricas
if 'Viagem' in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens['Viagem']):
    df_viagens['Viagem'] = df_viagens['Viagem'].astype("Int64")

if 'Turno' in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens['Turno']):
    df_viagens['Turno'] = df_viagens['Turno'].fillna(0).astype(int)

colunas_totais = ['DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
for col in colunas_totais:
    if col in df_viagens.columns and pd.api.types.is_float_dtype(df_viagens[col]):
        df_viagens[col] = df_viagens[col].fillna(0).astype(int)

# Reordena colunas
ordem_colunas = ['Data', 'Viagem', 'Placa', 'Status', 'Cluster', 'Turno',
                 'DropPoint', 'Eclusa', 'Aguardando SEC', 'SEC', 'Rup.Pend']
df_viagens = df_viagens[[col for col in ordem_colunas if col in df_viagens.columns]]

# Linha de total
linha_total = []
for col in df_viagens.columns:
    if col in colunas_totais:
        total = df_viagens[col].sum()
        linha_total.append(total)
    elif col == 'Cluster':
        linha_total.append('TOTAL')
    else:
        linha_total.append('')
df_total = pd.DataFrame([linha_total], columns=df_viagens.columns)
df_viagens = pd.concat([df_viagens, df_total], ignore_index=True)

# ==============================
# 🔹 Leitura da aba "Consolidado"
# ==============================
df_consolidado = pd.read_excel(arquivo, sheet_name='Consolidado', usecols="A:O")

# Remove linhas completamente vazias
df_consolidado = df_consolidado.dropna(how='all')

# Remove valores NaN (substitui por vazio para não aparecer no e-mail)
df_consolidado = df_consolidado.fillna("")

# 🔹 Converte colunas numéricas específicas para inteiro
colunas_int = ["Total und's", "Entrada", "Viagem", "Turno", "Ruptura"]

for col in colunas_int:
    if col in df_consolidado.columns:
        # Tenta converter; se falhar, ignora erro
        df_consolidado[col] = (
            pd.to_numeric(df_consolidado[col], errors='coerce')  # força números
            .fillna(0)  # substitui NaN por 0
            .astype(int)  # converte pra inteiro
        )

# ==============================
# 🔹 Estilo HTML compartilhado
# ==============================
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

# ==============================
# 🔹 Monta as duas tabelas HTML
# ==============================
html_viagens = df_viagens.to_html(index=False, border=0, justify='center')
html_consolidado = df_consolidado.to_html(index=False, border=0, justify='center')

# ==============================
# 🔹 Corpo do e-mail
# ==============================
corpo_email = estilo + f"""
<p>Segue acompanhamento da expedição:</p>

<h3>🚚 Tabela de Viagens</h3>
{html_viagens}

<br><br>

<h3>📊 Tabela Consolidado</h3>
{html_consolidado}
"""

# ==============================
# 🔹 Envia via Outlook
# ==============================
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "guilherme.dcastro@americanas.io"
mail.Subject = f"Acompanhamento Expedição - {datetime.now().strftime('%d/%m/%Y %H:%M')}"
mail.HTMLBody = corpo_email
mail.Send()

print("✅ E-mail enviado com sucesso!")
