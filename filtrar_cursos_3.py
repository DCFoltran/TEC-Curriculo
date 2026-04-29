import pandas as pd
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
# --- Carregamento do Arquivo CSV ---
#caminho_do_arquivo = 'relatorio_consulta_publica_avancada_curso_23_05_2025_01_22_05.csv' 

root = tk.Tk()
root.withdraw()
caminho_do_arquivo = filedialog.askopenfilename(
    title="Selecione uma planilha",
    initialdir=".", # Diretorio de início
    filetypes=(("Arquios CSV", "*.csv"), ("Arquios XLS", "*.xls"),("Todos os arquivos", "*.*"))
)
if caminho_do_arquivo:
    print(f"Arquivo selecionado: {caminho_do_arquivo}")
else:
    print("Nenhum arquivo selecionado.")
    quit()

gravar_arquivo = Path(caminho_do_arquivo).stem + datetime.today().strftime('__%d_%m_%Y_%H_%M_%S') + ".csv"
print(gravar_arquivo)
try:
    # Carrega o arquivo CSV.
    # header=5 significa que a 6ª linha do arquivo contém os nomes das colunas (0-indexed).
    # sep=';' indica que as colunas são separadas por ponto e vírgula.
    df = pd.read_csv(caminho_do_arquivo, header=0, sep=';', encoding='utf-8', dtype={'Situação do Curso':object})
    # Limpa os nomes das colunas (remove espaços extras)
    df.columns = [" ".join(str(col).strip().split()) for col in df.columns]
    # Remove colunas que são inteiramente vazias ou completamente "Unnamed"
    df.dropna(axis=1, how='all', inplace=True)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    print("Arquivo CSV carregado com sucesso!\n=========================================")
    print( df.columns )
#    print( df.dtypes )  # vericado para ajuste erro .str 
    print("=========================================")

    # --- Verificação dos Nomes das Colunas Esperadas (para segurança) ---
    colunas_necessarias = [
        'Nome da IES', 'Sigla da IES', 'Categoria Administrativa', 'Organização Acadêmica', 
        'Código do Curso', 'Nome do Curso', 'Qt. Vagas Autorizadas', 'Carga Horária',
        'Município', 'UF', 'Início Funcionamento', 'Data Ato de Criação',
        'Modalidade', 'Grau', 'Situação do Curso'
    ]
    colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
    if colunas_faltantes:
        print(f"AVISO: As seguintes colunas necessárias não foram encontradas no arquivo: {colunas_faltantes}")
        print("Por favor, verifique os nomes das colunas no seu arquivo e ajuste o script se necessário.")
    else:
        # --- Aplicação dos Filtros ---
        # É importante que os valores ("Presencial", "Licenciatura", "Em atividade")
        # correspondam exatamente ao que está no seu arquivo, incluindo maiúsculas/minúsculas e acentos.
        # O .str.strip() ajuda a remover espaços em branco antes/depois dos valores nas colunas.
        print("Filtrando ...")
        try:
            cursos_filtrados = df[
                (df['Modalidade'].str.contains("Presencial", regex=False) )
                & (df['Situação do Curso'].str.contains("atividade", case=False) )
                & (df['Grau'].str.strip() == "Licenciatura")
                & (df['Nome do Curso'].str.contains("Educação Especial",case=False))
                & (df['Início Funcionamento'].str.strip() != "Não iniciado")
            ]
        except Exception as e:
            print(f"ERRO: {e}")
            quit()
            
        print( cursos_filtrados )

        # --- Seleção das Colunas para a Lista Final ---
        colunas_para_exibir = [
            'Nome da IES', 'Sigla da IES', 'Categoria Administrativa', 'Organização Acadêmica', 
            'Código do Curso', 'Nome do Curso', 'Qt. Vagas Autorizadas', 'Carga Horária',
            'Município', 'UF', 'Início Funcionamento', 'Data Ato de Criação',
            'Modalidade', 'Grau', 'Situação do Curso',  
        ]
        #lista_final = cursos_filtrados[colunas_para_exibir]
        lista_final = cursos_filtrados
        # --- Exibição da Lista ---
        if not lista_final.empty:
            print("--- Lista de Cursos Filtrados ---")
            print(lista_final.to_string(index=False)) # .to_string() para ver todas as linhas
            lista_final.to_csv(gravar_arquivo, index=False, sep=';', encoding='utf-8-sig')
        else:
            print("Nenhum curso encontrado com os critérios especificados.")

except FileNotFoundError:
    print(f"ERRO: Arquivo não encontrado em '{caminho_do_arquivo}'. Verifique o caminho.")
except Exception as e:
    print(f"Ocorreu um erro: {e}")