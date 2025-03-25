import streamlit as st
import pandas as pd
import os
import shutil
from datetime import datetime
import openpyxl

def dividir_planilha(df, linhas_por_arquivo):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    pasta_saida = f"arquivos_divididos_{timestamp}"
    os.makedirs(pasta_saida, exist_ok=True)

    df.columns = df.columns.str.replace("\.\d+", "", regex=True)

    total_partes = (len(df) // linhas_por_arquivo) + (1 if len(df) % linhas_por_arquivo else 0)
    arquivos_gerados = []

    for i in range(total_partes):
        inicio = i * linhas_por_arquivo
        fim = inicio + linhas_por_arquivo
        df_parte = df.iloc[inicio:fim]

        arquivo_saida = os.path.join(pasta_saida, f"parte_{i+1}.xlsx")
        df_parte.to_excel(arquivo_saida, index=False)
        arquivos_gerados.append(arquivo_saida)

    return arquivos_gerados, pasta_saida

def criar_zip(pasta_saida):
    zip_filename = f"{pasta_saida}.zip"
    shutil.make_archive(pasta_saida, 'zip', pasta_saida)
    return zip_filename

st.title("Divisor de Planilhas Excel")

arquivo = st.file_uploader("Carregue um arquivo Excel", type=["xlsx"])

st.write("")
linhas_por_arquivo = st.number_input("Número de linhas por arquivo", min_value=1, value=50)

iniciar = st.button("Iniciar Processamento")

if iniciar and arquivo:
    df = pd.read_excel(arquivo, dtype=str)

    st.write("Processando, aguarde...")
    arquivos_gerados, pasta_saida = dividir_planilha(df, linhas_por_arquivo)
    zip_file = criar_zip(pasta_saida)

    st.success(f"{len(arquivos_gerados)} arquivos gerados e compactados!")

    with open(zip_file, "rb") as f:
        st.download_button(
            label="Baixar Arquivos ZIP",
            data=f,
            file_name=os.path.basename(zip_file),
            mime="application/zip"
        )

    shutil.rmtree(pasta_saida)
    os.remove(zip_file)
