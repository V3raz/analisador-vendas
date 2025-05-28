#-------------------------------------------------------------------------------------------------------------------------
# IMPORTANTE!
# Para rodar o programa como app, abra o terminal e digite:
# python -m streamlit run app.py
#-------------------------------------------------------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Analisador de Vendas", layout="centered")

# Cabeçalho com imagem (opcional)
col1, col2, col3 = st.columns([1, 3, 1])
with col2:
    st.image("https://i.imgur.com/h2OZB8n.png", width=120)

st.title("📊 Analisador de Vendas com Python")
st.write("Faça upload da planilha de vendas (.xlsx) e receba uma análise automática com gráficos e relatório em Excel.")

# Upload da planilha
uploaded_file = st.file_uploader("📎 Faça upload da planilha de vendas", type=["xlsx"])

# Funções
def carregar_dados(file):
    df = pd.read_excel(file)
    if df.empty:
        st.error("A planilha está vazia!")
        st.stop()
    df["Data da Venda"] = pd.to_datetime(df["Data da Venda"])
    df["Mês"] = df["Data da Venda"].dt.to_period("M").astype(str)
    return df

def gerar_relatorios(df):
    por_produto_valor = df.groupby("Produto")["Valor da Venda"].sum().reset_index()
    por_mes = df.groupby("Mês")["Valor da Venda"].sum().reset_index()
    por_regiao = df.groupby("Região")["Valor da Venda"].sum().reset_index()
    return por_produto_valor, por_mes, por_regiao

def gerar_excel(produto, mes, regiao, resumo):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumo.to_excel(writer, sheet_name="Resumo", index=False)
        produto.to_excel(writer, sheet_name="Por Produto", index=False)
        mes.to_excel(writer, sheet_name="Por Mês", index=False)
        regiao.to_excel(writer, sheet_name="Por Região", index=False)

        # Formatação
        workbook = writer.book
        for sheet in writer.sheets:
            ws = writer.sheets[sheet]
            for column_cells in ws.columns:
                length = max(len(str(cell.value)) for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = length + 2
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = openpyxl.styles.Border(
                        left=openpyxl.styles.Side(style="thin"),
                        right=openpyxl.styles.Side(style="thin"),
                        top=openpyxl.styles.Side(style="thin"),
                        bottom=openpyxl.styles.Side(style="thin"),
                    )
    return output.getvalue()

# Processamento principal
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    colunas_esperadas = {"Data da Venda", "Produto", "Região", "Valor da Venda"}

    if not colunas_esperadas.issubset(set(df.columns)):
        st.error("❌ A planilha deve conter as colunas: Data da Venda, Produto, Região, Valor da Venda.")
    else:
        df = carregar_dados(uploaded_file)
        por_produto_valor, por_mes, por_regiao = gerar_relatorios(df)

        # Análises
        total_geral = df["Valor da Venda"].sum()
        produto_faturamento = por_produto_valor.sort_values("Valor da Venda", ascending=False).iloc[0]

        produto_qtd = df["Produto"].value_counts()
        produto_mais_vendido_qtd = produto_qtd.idxmax()
        qtd_vendas = produto_qtd.max()

        regiao_top = por_regiao.sort_values("Valor da Venda", ascending=False).iloc[0]

        # Resumo na tela
        st.subheader("✅ Resumo das Vendas")
        st.metric(label="💰 Total Geral Vendido", value=f"R$ {total_geral:,.2f}")
        st.success(f"💵 Produto com maior faturamento: {produto_faturamento['Produto']} (R$ {produto_faturamento['Valor da Venda']:,.2f})")
        st.success(f"📦 Produto mais vendido (quantidade): {produto_mais_vendido_qtd} ({qtd_vendas} vendas)")
        st.success(f"🌍 Região com maior faturamento: {regiao_top['Região']}")

        # Tabelas
        st.subheader("📈 Vendas por Produto (R$)")
        st.dataframe(por_produto_valor)

        st.subheader("📅 Vendas por Mês")
        st.dataframe(por_mes)

        st.subheader("🌍 Vendas por Região")
        st.dataframe(por_regiao)

        # Gráficos
        st.subheader("📊 Gráfico de Vendas por Produto")
        st.bar_chart(por_produto_valor.set_index("Produto"))

        st.subheader("📊 Gráfico de Vendas por Região")
        st.bar_chart(por_regiao.set_index("Região"))

        # Planilha "Por Produto" com quantidade
        produto_completo = por_produto_valor.copy()
        produto_completo["Quantidade de Vendas"] = df["Produto"].value_counts().reindex(produto_completo["Produto"]).values

        # Planilha "Resumo"
        resumo_df = pd.DataFrame({
            "Resumo": [
                f"Total Geral Vendido: R$ {total_geral:,.2f}",
                f"Produto com maior faturamento: {produto_faturamento['Produto']}",
                f"Produto mais vendido (quantidade): {produto_mais_vendido_qtd}",
                f"Região com maior faturamento: {regiao_top['Região']}"
            ]
        })

        # Botão de download
        nome_arquivo = f"relatorio_vendas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=gerar_excel(produto_completo, por_mes, por_regiao, resumo_df),
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
