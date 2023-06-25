import streamlit as st
import pandas as pd
import numpy as np
import io

# Função para identificar problemas na planilha
def identificar_problemas(df):
    problemas = []
    
    # Verificar valores numéricos em colunas de nomes
    for coluna in df.columns:
        if df[coluna].dtype == 'object':
            valores_numericos = df[coluna].str.isnumeric()
            if valores_numericos.any():
                linhas = df[valores_numericos][coluna].index.tolist()
                problemas.append(f"Valores numéricos encontrados na coluna '{coluna}' nas linhas {linhas}.")
    
    # Verificar valores em branco
    valores_em_branco = df.isnull().any(axis=1)
    if valores_em_branco.any():
        linhas = df[valores_em_branco].index.tolist()
        problemas.append(f"Valores em branco encontrados nas linhas {linhas}.")
    
    # Verificar linhas duplicadas
    duplicatas = df.duplicated()
    if duplicatas.any():
        linhas = df[duplicatas].index.tolist()
        problemas.append(f"Linhas duplicadas encontradas nas linhas {linhas}.")
    
    # Verificar valores negativos
    for coluna in df.columns:
        if df[coluna].dtype in ['int64', 'float64']:
            valores_negativos = df[coluna] < 0
            if valores_negativos.any():
                linhas = df[valores_negativos].index.tolist()
                problemas.append(f"Valores negativos encontrados na coluna '{coluna}' nas linhas {linhas}.")
    
    # Verificar nomes próprios iniciados com letra minúscula
    for coluna in df.columns:
        if df[coluna].dtype == 'object':
            nomes_minusculos = df[coluna].str.contains(r'\b[a-z]\w*\b', na=False)
            if nomes_minusculos.any():
                linhas = df.loc[nomes_minusculos, coluna].index.tolist()
                problemas.append(f"Nomes próprios iniciados com letra minúscula encontrados na coluna '{coluna}' nas linhas {linhas}.")
    
    return problemas

# Função para corrigir problemas na planilha
def corrigir_problemas(df, problemas_corrigir):
    # Remover linhas duplicadas
    if 'Linhas duplicadas' in problemas_corrigir:
        df = df.drop_duplicates()
    
    # Preencher valores em branco com uma string vazia ou permitir que o usuário digite valores
    if 'Valores em branco' in problemas_corrigir:
        for i, linha in df.iterrows():
            if linha.isnull().all():
                df.drop(i, inplace=True)
            elif linha.isnull().any():
                st.write(f"Linha {i+1}")
                st.write(linha)
                resposta_preencher = st.radio("Deseja preencher os valores em branco desta linha?", options=["Sim", "Não"])
                if resposta_preencher == "Sim":
                    for coluna in df.columns:
                        if pd.isnull(linha[coluna]):
                            novo_valor = st.text_input(f"Digite o novo valor para a coluna '{coluna}'", value="")
                            df.at[i, coluna] = novo_valor
    
    # Converter valores numéricos em colunas de nomes para string vazia ou permitir que o usuário digite valores
    if 'Valores numéricos' in problemas_corrigir:
        for i, linha in df.iterrows():
            for coluna in df.columns:
                if pd.to_numeric(linha[coluna], errors='coerce') and not pd.isnull(linha[coluna]):
                    st.write(f"Linha {i+1} - Coluna '{coluna}'")
                    st.write(linha[coluna])
                    resposta_preencher = st.radio("Deseja corrigir este valor numérico?", options=["Sim", "Não"])
                    if resposta_preencher == "Sim":
                        novo_valor = st.text_input(f"Digite o novo valor para a coluna '{coluna}'", value="")
                        df.at[i, coluna] = novo_valor
    
    # Converter valores negativos para zero ou permitir que o usuário digite valores
    if 'Valores negativos' in problemas_corrigir:
        for i, linha in df.iterrows():
            for coluna in df.columns:
                if pd.to_numeric(linha[coluna], errors='coerce') and linha[coluna] < 0:
                    st.write(f"Linha {i+1} - Coluna '{coluna}'")
                    st.write(linha[coluna])
                    resposta_preencher = st.radio("Deseja corrigir este valor negativo?", options=["Sim", "Não"])
                    if resposta_preencher == "Sim":
                        novo_valor = st.text_input(f"Digite o novo valor para a coluna '{coluna}'", value="")
                        df.at[i, coluna] = novo_valor
    
    # Corrigir nomes próprios iniciados com letra minúscula
    if 'Nomes próprios iniciados com letra minúscula' in problemas_corrigir:
        for coluna in df.columns:
            if df[coluna].dtype == 'object':
                nomes_minusculos = df[coluna].str.contains(r'\b[a-z]\w*\b', na=False)
                df.loc[nomes_minusculos, coluna] = df.loc[nomes_minusculos, coluna].str.capitalize()
    
    return df

# Configurações do Streamlit
st.set_page_config(layout="wide")
st.title("Identificação e Correção de Problemas em Planilhas Excel")

# Upload do arquivo Excel
st.sidebar.header("Upload do arquivo Excel")
uploaded_file = st.sidebar.file_uploader("Selecione um arquivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Leitura do arquivo Excel
    df = pd.read_excel(uploaded_file)
    
    # Exibição dos dados
    st.header("Dados do arquivo Excel")
    st.dataframe(df)
    
    # Identificação dos problemas
    problemas = identificar_problemas(df)
    
    # Exibição dos problemas identificados
    if problemas:
        st.header("Problemas identificados")
        for problema in problemas:
            st.write(problema)
        
        # Seleção dos problemas a serem corrigidos
        problemas_corrigir = st.multiselect("Selecione os problemas a serem corrigidos", problemas)
        
        if problemas_corrigir:
            # Botão de confirmação
            if st.button("Confirmar"):
                # Correção dos problemas selecionados
                df_corrigido = corrigir_problemas(df, problemas_corrigir)
                
                # Download do arquivo Excel corrigido
                st.header("Arquivo Excel corrigido")
                st.dataframe(df_corrigido)
                
                # Salvar alterações
                st.write("Salvar alterações:")
                st.download_button(
                    label="Download",
                    data=df_corrigido.to_excel,
                    file_name="planilha_corrigida.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.write("Nenhum problema identificado no arquivo Excel.")
