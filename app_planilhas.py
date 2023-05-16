import pandas as pd
import streamlit as st
from io import BytesIO

def remove_duplicated_rows(df):
    df.drop_duplicates(inplace=True)
    return df

def remove_blank_rows(df):
    df.dropna(inplace=True)
    return df

def convert_to_complete_date(df, column_name):
    df[column_name] = pd.to_datetime(df[column_name])
    return df


def create_day_column(df, column_name):
    df[column_name + "_day"] = df[column_name].dt.day
    return df

def create_month_column(df, column_name):
    df[column_name + "_month"] = df[column_name].dt.month
    return df

def create_year_column(df, column_name):
    df[column_name + "_year"] = df[column_name].dt.year
    return df

def capitalize_first_letter(df, column_name):
    df[column_name] = df[column_name].str.title()
    return df

def main():
    st.title("Correções em Arquivos Excel")

    uploaded_file = st.file_uploader("Faça o upload de um arquivo Excel", type=["xls", "xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        st.subheader("DataFrame Original")
        st.write(df)

        st.subheader("Deseja realizar quais correções?")
        corrections = st.multiselect(
            "Selecione as correções desejadas:",
            ["Remover linhas duplicadas", "Remover linhas em branco"]
        )

        if "Remover linhas duplicadas" in corrections:
            df = remove_duplicated_rows(df)

        if "Remover linhas em branco" in corrections:
            df = remove_blank_rows(df)

        st.subheader("DataFrame Após Remoção de Linhas")
        st.write(df)

        column_names = list(df.columns)
        selected_column = st.selectbox("Selecione a coluna para correção", column_names)

        if selected_column:
            selected_column_type = df[selected_column].dtype

            st.subheader(f"Correções para a coluna '{selected_column}'")

            if selected_column_type == "datetime64[ns]":
                date_corrections = st.multiselect(
                    "Selecione as correções desejadas:",
                    ["Converter para data completa", "Criar coluna de dia", "Criar coluna de mês", "Criar coluna de ano"]
                )

                if "Converter para data completa" in date_corrections:
                    df = convert_to_complete_date(df, selected_column)

                if "Criar coluna de dia" in date_corrections:
                    df = create_day_column(df, selected_column)

                if "Criar coluna de mês" in date_corrections:
                    df = create_month_column(df, selected_column)

                if "Criar coluna de ano" in date_corrections:
                    df = create_year_column(df, selected_column)

            elif selected_column_type == "object":
                text_corrections = st.multiselect(
                    "Selecione as correções desejadas:",
                    ["Converter primeira letra de cada palavra para maiúscula"]
                )

                if "Converter primeira letra de cada palavra para maiúscula" in text_corrections:
                    df = capitalize_first_letter(df, selected_column)

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
