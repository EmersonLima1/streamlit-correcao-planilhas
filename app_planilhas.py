import pandas as pd
import streamlit as st
from io import BytesIO

def remover_linhas_duplicadas(df):
    df.drop_duplicates(inplace=True)
    return df

def remover_linhas_em_branco(df):
    df.dropna(inplace=True)
    return df

def converter_para_data_completa(df, nome_coluna):
    df[nome_coluna] = pd.to_datetime(df[nome_coluna])
    return df

def criar_coluna_dia(df, nome_coluna):
    df[nome_coluna + "_dia"] = df[nome_coluna].dt.day
    return df

def criar_coluna_mes(df, nome_coluna):
    df[nome_coluna + "_mes"] = df[nome_coluna].dt.month
    return df

def criar_coluna_ano(df, nome_coluna):
    df[nome_coluna + "_ano"] = df[nome_coluna].dt.year
    return df

def capitalizar_primeira_letra(df, nome_coluna):
    df[nome_coluna] = df[nome_coluna].str.title()
    return df

def main():
    st.title("Correções em Arquivos Excel")

    uploaded_file = st.file_uploader("Faça o upload de um arquivo Excel", type=["xls", "xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        st.subheader("DataFrame Original")
        st.write(df)

        st.subheader("Deseja realizar quais correções?")
        correcoes = st.multiselect(
            "Selecione as correções desejadas:",
            ["Remover linhas duplicadas", "Remover linhas em branco"]
        )

        if "Remover linhas duplicadas" in correcoes:
            df = remover_linhas_duplicadas(df)

        if "Remover linhas em branco" in correcoes:
            df = remover_linhas_em_branco(df)

        st.subheader("DataFrame Após Remoção de Linhas")
        st.write(df)

        nomes_colunas = list(df.columns)
        coluna_selecionada = st.selectbox("Selecione a coluna para correção", nomes_colunas)

        if st.button("Aplicar correções"):
            if coluna_selecionada:
                tipo_coluna_selecionada = df[coluna_selecionada].dtype

                st.subheader(f"Correções para a coluna '{coluna_selecionada}'")

                if tipo_coluna_selecionada == "datetime64[ns]":
                    correcoes_data = st.multiselect(
                        "Selecione as correções desejadas:",
                        ["Converter para data completa", "Criar coluna de dia", "Criar coluna de mês", "Criar coluna de ano"]
                    )

                    if "Converter para data completa" in correcoes_data:
                        df = converter_para_data_completa(df, coluna_selecionada)

                    if "Criar coluna de dia" in correcoes_data:
                        df = criar_coluna_dia(df, coluna_selecionada)

                    if "Criar coluna de mês" in correcoes_data:
                        df = criar_coluna_mes(df, coluna_selecionada)

                    if "Criar coluna de ano" in correcoes_data:
                        df = criar_coluna_ano(df, coluna_selecionada)

                elif tipo_coluna_selecionada == "object":
                    correcoes_texto = st.multiselect(
                        "Selecione as correções desejadas:",
                        ["Converter primeira letra de cada palavra para maiúscula"]
                    )

                    if "Converter primeira letra de cada palavra para maiúscula" in correcoes_texto:
                        df = capitalizar_primeira_letra(df, coluna_selecionada)

                st.subheader("DataFrame Corrigido")
                st.write(df)

                # Download do arquivo corrigido
                st.subheader("Baixar arquivo corrigido")
                excel_file = BytesIO()
                with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                excel_file.seek(0)
                st.download_button(
                    label="Baixar arquivo corrigido",
                    data=excel_file,
                    file_name="dataframe_corrigido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
