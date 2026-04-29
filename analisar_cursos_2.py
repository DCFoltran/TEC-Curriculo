import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from datetime import datetime
# --- Carregamento do Arquivo CSV ---
#caminho_do_arquivo = 'relatorio_consulta_publica_avancada_curso_23_05_2025_01_22_05.csv' 

root = tk.Tk()
root.withdraw()
caminho_do_arquivo = filedialog.askopenfilename(
    title="Selecione uma planilha para análise...",
    initialdir=".", # Diretorio de início
    filetypes=(("Arquios CSV", "*.csv"), ("Arquios XLS", "*.xls"),("Todos os arquivos", "*.*"))
)
if caminho_do_arquivo:
    print(f"Arquivo selecionado: {caminho_do_arquivo}")
else:
    print("Nenhum arquivo selecionado.")
    quit()

file_name = Path(caminho_do_arquivo).stem + datetime.today().strftime('__%d_%m_%Y_%H_%M_%S') + ".csv"
#file_name = "relatorio_consulta_publica_avancada_curso_25_08_2025_10_43_33__05_02_2026_15_52_19.csv"

# Carregar o dataset
try:
    df = pd.read_csv(file_name, sep=';', encoding='utf-8')
except:
    try:
        df = pd.read_csv(file_name, sep=';', encoding='latin1')
    except:
        df = pd.read_csv(file_name, sep=',', encoding='utf-8')

# Limpeza e Mapeamentos
if 'UF' in df.columns:
    df['UF'] = df['UF'].astype(str).str.strip()

region_map = {
    'AC': 'Norte', 'AP': 'Norte', 'AM': 'Norte', 'PA': 'Norte', 'RO': 'Norte', 'RR': 'Norte', 'TO': 'Norte',
    'AL': 'Nordeste', 'BA': 'Nordeste', 'CE': 'Nordeste', 'MA': 'Nordeste', 'PB': 'Nordeste', 'PE': 'Nordeste', 'PI': 'Nordeste', 'RN': 'Nordeste', 'SE': 'Nordeste',
    'DF': 'Centro-Oeste', 'GO': 'Centro-Oeste', 'MT': 'Centro-Oeste', 'MS': 'Centro-Oeste',
    'ES': 'Sudeste', 'MG': 'Sudeste', 'RJ': 'Sudeste', 'SP': 'Sudeste',
    'PR': 'Sul', 'RS': 'Sul', 'SC': 'Sul'
}
if 'UF' in df.columns:
    df['Região'] = df['UF'].map(region_map)

# Função auxiliar para adicionar totais
def add_total_row(df_in, label_col, sum_col=None):
    df_out = df_in.copy()
    if sum_col:
        total = df_out[sum_col].sum()
        df_out['%'] = (df_out[sum_col] / total) * 100
        df_out['%'] = df_out['%'].apply(lambda x: "{:.2f}".format(x).replace('.', ','))
        new_row = {label_col: 'Total', sum_col: total, '%': '100,00'}
    else:
        # Para crosstabs simples
        pass 
    return pd.concat([df_out, pd.DataFrame([new_row])], ignore_index=True)

# --- Gerar Tabelas ---

# Tabela 1: Regiões (Cursos)
tab1 = df['Região'].value_counts().reset_index()
tab1.columns = ['Regiões Administrativas', 'Quantidade']
tab1 = add_total_row(tab1, 'Regiões Administrativas', 'Quantidade')

# Tabela 2: UFs (Cursos)
tab2 = df['UF'].value_counts().reset_index()
tab2.columns = ['UF', 'Quantidade']
tab2 = add_total_row(tab2, 'UF', 'Quantidade')

# Tabela 3: Categoria Adm (Cursos)
tab3 = df['Categoria Administrativa'].value_counts().reset_index()
tab3.columns = ['Categoria Administrativa', 'Quantidade']
tab3 = add_total_row(tab3, 'Categoria Administrativa', 'Quantidade')

# Tabela 4: Vagas por Categoria
tab4 = df.groupby('Categoria Administrativa')['Qt. Vagas Autorizadas'].sum().reset_index()
tab4.columns = ['Categoria Administrativa', 'Quantidade']
tab4 = tab4.sort_values(by='Quantidade', ascending=False)
tab4 = add_total_row(tab4, 'Categoria Administrativa', 'Quantidade')

# Tabela 5: Vagas por Região
tab5 = df.groupby('Região')['Qt. Vagas Autorizadas'].sum().reset_index()
tab5.columns = ['Regiões Administrativas', 'Quantidade']
tab5 = tab5.sort_values(by='Quantidade', ascending=False)
tab5 = add_total_row(tab5, 'Regiões Administrativas', 'Quantidade')

# Tabela 7: UF x Categoria (Cursos)
tab7 = pd.crosstab(df['UF'], df['Categoria Administrativa'], margins=True, margins_name='Total Geral').reset_index()
tab7.columns.name = None

# Tabela 8: UF x Categoria (Vagas)
tab8 = pd.pivot_table(df, values='Qt. Vagas Autorizadas', index='UF', columns='Categoria Administrativa', aggfunc='sum', fill_value=0, margins=True, margins_name='Total').reset_index()
tab8.columns.name = None

# Tabela 9: ENADE x Categoria
df['Valor ENADE_Label'] = pd.to_numeric(df['Valor ENADE'], errors='coerce').fillna('SC')
tab9 = pd.crosstab(df['Categoria Administrativa'], df['Valor ENADE_Label'], margins=True, margins_name='Total Geral').reset_index()
tab9.columns.name = None

# Tabela 10: ENADE x UF
tab10 = pd.crosstab(df['UF'], df['Valor ENADE_Label'], margins=True, margins_name='Total').reset_index()
tab10.columns.name = None

# Dados Figura 7
dados_fig7 = df['Carga Horária'].value_counts().sort_index().reset_index()
dados_fig7.columns = ['Carga Horária', 'Quantidade']

# --- Gerar Gráfico e Salvar Imagem ---
plt.figure(figsize=(10, 6))
bars = plt.bar(dados_fig7['Carga Horária'].astype(str), dados_fig7['Quantidade'], color='skyblue', edgecolor='black')
plt.title('Distribuição dos Cursos por Carga Horária')
plt.xlabel('Carga Horária (horas)')
plt.ylabel('Quantidade de Cursos')
for bar in bars:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', va='bottom')
plt.tight_layout()
img_filename = 'figura_7_export.png'
plt.savefig(img_filename)
plt.close()

# --- Exportar para Excel ---
output_file = 'Analise_Dados_Cursos.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    tab1.to_excel(writer, sheet_name='Tabela 1 - Regiões', index=False)
    tab2.to_excel(writer, sheet_name='Tabela 2 - UF', index=False)
    tab3.to_excel(writer, sheet_name='Tabela 3 - Cat Adm', index=False)
    tab4.to_excel(writer, sheet_name='Tabela 4 - Vagas Cat', index=False)
    tab5.to_excel(writer, sheet_name='Tabela 5 - Vagas Reg', index=False)
    tab7.to_excel(writer, sheet_name='Tabela 7 - UF x Cat', index=False)
    tab8.to_excel(writer, sheet_name='Tabela 8 - Vagas UF x Cat', index=False)
    tab9.to_excel(writer, sheet_name='Tabela 9 - ENADE Cat', index=False)
    tab10.to_excel(writer, sheet_name='Tabela 10 - ENADE UF', index=False)
    dados_fig7.to_excel(writer, sheet_name='Figura - Carga Horária', index=False)
    
    # Inserir imagem na aba Figura
    worksheet = writer.sheets['Figura - Carga Horária']
    worksheet.insert_image('D2', img_filename)

print(f"Arquivo '{output_file}' gerado com sucesso.")