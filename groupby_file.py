import os
import pandas as pd

import re

def limpar_nome(nome):
    # troca caracteres proibidos por "_"
    nome = re.sub(r'[<>:"/\\|?*]', '_', nome)
    # troca espaços e vírgulas por "_" opcionalmente
    #nome = re.sub(r'[\s,]+', '_', nome)
    # corta para evitar path longo demais
    return nome[:150]

pasta = "label_group_xlsx"
saida = "arquivos_por_documento"
os.makedirs(saida, exist_ok=True)

# dicionário: nome_arquivo -> lista de registros
coletor = {}

for arq in os.listdir(pasta):
    if not arq.endswith(".xlsx"):
        continue
    
    df = pd.read_excel(os.path.join(pasta, arq))
    
    if "Filename" not in df.columns:
        continue
    
    # demais colunas = categorias
    categorias = [c for c in df.columns if c != "Filename"]
    
    for _, row in df.iterrows():
        nome = str(row["Filename"])
        if nome not in coletor:
            coletor[nome] = []
        
        registro = {"categoria_origem": arq}
        for c in categorias:
            registro[c] = row.get(c, None)
        
        coletor[nome].append(registro)

# gerar um XLSX por filename
for nome, registros in coletor.items():
    df_out = pd.DataFrame(registros)
    nome_limpo = limpar_nome(str(nome))
    caminho_saida = os.path.join(saida, f"{nome_limpo}.xlsx")
    df_out.to_excel(caminho_saida, index=False)