import streamlit as st
import pandas as pd

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
            nomes_minusculos = df[coluna].str.contains(r'\b[a-z]\w*\b')
            if nomes_minusculos.any():
                linhas = df[nomes_minusculos][coluna].index.tolist()
                problemas.append(f"Nomes próprios iniciados com letra minúscula encontrados na coluna '{coluna}' nas linhas {linhas}.")
    
    return problemas

# Função para corrigir problemas na planilha
def corrigir_problemas(df, problemas_corrigir):
    # Remover linhas duplicadas
    if 'Linhas duplicadas' in problemas_corrigir:
        df = df.drop_duplicates()
    
    # Preencher valores em branco com uma string vazia
    if 'Valores em branco' in problemas_corrigir:
        df = df.dropna()
    
    # Converter valores numéricos em colunas de nomes para string vazia
    if 'Valores numéricos' in problemas_corrigir:
        for coluna in df.columns:
            if df[coluna].dtype == 'object':
                valores_numericos = df[coluna].str.isnumeric()
                df.loc[valores_numericos, coluna] = ''
    
    # Converter valores negativos para zero
    if 'Valores negativos' in problemas_corrigir:
        for coluna in df.columns:
            if df[coluna].dtype in ['int64', 'float64']:
                valores_negativos = df[coluna] < 0
                df.loc[valores_negativos, coluna] = 0
                
    # Corrigir nomes próprios iniciados com letra minúscula
    if 'Nomes próprios iniciados com letra minúscula' in problemas_corrigir:
        for coluna in df.columns:
            if df[coluna].dtype == 'object':
                nomes_minusculos = df[coluna].str.contains(r'\b[a-z]\w*\b')
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
            # Correção dos problemas selecionados
            df_corrigido = corrigir_problemas(df, problemas_corrigir)
            
            # Download do arquivo Excel corrigido
            st.header("Arquivo Excel corrigido")
            st.dataframe(df_corrigido)
            st.download_button("Baixar arquivo Excel corrigido", df_corrigido.to_excel, file_name="planilha_corrigida.xlsx", label="Clique aqui para baixar")
    else:
        st.write("Nenhum problema identificado.")
