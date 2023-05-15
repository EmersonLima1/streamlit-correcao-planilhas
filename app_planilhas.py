import pandas as pd
import streamlit as st
from io import BytesIO


def remove_duplicated_rows(df):
    df.drop_duplicates(inplace=True)
    return df


def remove_blank_rows(df):
    df.dropna(inplace=True)
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

        if st.button("Aplicar correções"):
            if "Remover linhas duplicadas" in corrections:
                df = remove_duplicated_rows(df)

            if "Remover linhas em branco" in corrections:
                df = remove_blank_rows(df)

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
