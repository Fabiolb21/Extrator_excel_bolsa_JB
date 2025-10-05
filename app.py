import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Processador de Etiquetas", layout="wide")

st.title("Processador de Planilhas")

# st.write("Este aplicativo processa um arquivo Excel para extrair e formatar dados conforme as especificações.")
st.markdown("""
Esta aplicação processa uma planilha Excel para extrair e transformar dados, gerando uma planilha única pronta para a criação de etiquetas.

**Instruções:**
1. Faça o upload do seu arquivo Excel (formato `.xlsm` ou `.xlsx`).
2. Aguarde o processamento dos dados.
3. Faça o download da planilha gerada.
""")

st.divider()

uploaded_file = st.file_uploader("Escolha um arquivo Excel (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name='RELATORIO')

        # 1. Manter colunas existentes
        df_final = pd.DataFrame()
        df_final["PLANO_PROD"] = df["PLANO_PROD"]
        df_final["OF_NUMERO"] = df["OF_NUMERO"]
        df_final["PROD_DESCRICAO"] = df["PRODUTO_DESCRICAO"]

        # 4. Criar coluna PRODUTO
        df_final["PRODUTO"] = df["PRODUTO_DESCRICAO"].apply(lambda x: x.split(" ")[0])

        # 5. Criar coluna DESCRICAO 1
        df_final["DESCRICAO 1"] = df["PRODUTO_DESCRICAO"].apply(lambda x: " ".join(x.split(" ")[1:]))

        # 6. Criar coluna DESCRICAO 2
        df_final["DESCRICAO 2"] = df["PRODUTO_DESCRICAO"].apply(lambda x: "/".join(x.split("/")[1:]) if "/" in x else "")

        # 7. Manter PROD_CODIGO
        df_final["PROD_CODIGO"] = df["PRODUTO_CODIGO"]

        # 8. Formatar PRECO_UNIT_PDV
        df_final["PRECO_UNIT_PDV"] = df["PRECO_UNIT_PDV"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        # 9. Manter e formatar GRADE_TAMANHO
        df_final["GRADE_TAMANHO"] = df["GRADE_TAMANHO"].fillna("")

        # 10. Manter CODIGO_BARRAS
        df_final["CODIGO_BARRAS"] = df["CODIGO_BARRAS"]

        # 11. Manter QTD
        df_final["QTD"] = df["QTD"]

        # 12. Manter UNID_MEDIDA
        df_final["UNID_MEDIDA"] = df["UNID_MEDIDA"]

        # 13. Criar PRODUTO_IMAGEM
        df_final["PRODUTO_IMAGEM"] = df["PRODUTO_DESCRICAO"].str[:9]

        # 14. Manter SEQUENCIA_PEDIDO
        df_final["SEQUENCIA_PEDIDO"] = df["SEQUENCIA_PEDIDO"]

        # 15. Manter ENTREGA_ID
        df_final["ENTREGA_ID"] = df["ENTREGA_ID"]

        # 16. Manter SITUACAO_ENTREGA
        df_final["SITUACAO_ENTREGA"] = df["SITUACAO_ENTREGA"]

        # 17. Manter PEDIDO
        df_final["PEDIDO"] = df["PEDIDO"]

        # 18. Criar CAMINHO (vazio)
        df_final["CAMINHO"] = ""

        # 19. Criar IMAGEM (vazio)
        df_final["IMAGEM"] = ""

        # Ordenar os dados
        df_final = df_final.sort_values(by=["PEDIDO", "PROD_CODIGO"], ascending=[True, True])

        # Inserir linhas em branco
        df_final_com_espacos = pd.DataFrame()
        pedidos = df_final["PEDIDO"].unique()
        for i, pedido in enumerate(pedidos):
            df_pedido = df_final[df_final["PEDIDO"] == pedido]
            df_final_com_espacos = pd.concat([df_final_com_espacos, df_pedido])
            if i < len(pedidos) - 1:
                df_final_com_espacos = pd.concat([df_final_com_espacos, pd.DataFrame([["" for _ in df_final.columns]], columns=df_final.columns)])

        # Formatar todas as colunas como texto
        df_final_com_espacos = df_final_com_espacos.astype(str)

        # Gerar o arquivo Excel .xls usando xlwt
        import xlwt
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('RELATORIO_PROCESSADO')

        # Escrever cabeçalhos
        for col_idx, col_name in enumerate(df_final_com_espacos.columns):
            worksheet.write(0, col_idx, str(col_name))

        # Escrever dados
        for row_idx, row in enumerate(df_final_com_espacos.values, start=1):
            for col_idx, value in enumerate(row):
                worksheet.write(row_idx, col_idx, str(value))

        # Salvar em BytesIO
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        
        st.success("Planilha processada com sucesso!")

        st.download_button(
            label="Baixar Planilha Processada (.xls)",
            data=output.getvalue(),
            file_name="relatorio_processado.xls",
            mime="application/vnd.ms-excel"
        )

        # Mostrar preview dos dados
        with st.expander("👁️ Visualizar Preview dos Dados Processados"):
                st.dataframe(df_final_com_espacos.head(100), use_container_width=True)

    except Exception as e:
        st.error(f"Ocorreu um erro: {e}")

