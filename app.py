#-------------------------------------------------------------------------------------------------------------------------
# IMPORTANTE!
# Para rodar o programa como app, abra o terminal e digite:
# python -m streamlit run app.py
# Ou use o executável .exe gerado com PyInstaller.
#-------------------------------------------------------------------------------------------------------------------------

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime

# Configuração da página
st.set_page_config(page_title="Analisador de Vendas", layout="centered")

# Cabeçalho com imagem opcional (substitua se quiser)
col1, col2, col3 = st.columns([1, 3, 1])
with col2:
    st.image("https://i.imgur.com/h2OZB8n.png", width=120)  # Ou troque pela sua própria imagem

st.title("📊 Analisador de Vendas com Python")
st.write("Faça upload da planilha de vendas (.xlsx) e receba uma análise automática com gráficos e relatório Excel.")

# Upload da planilha
uploaded_file = st.file_uploader("📎 Faça upload da planilha de vendas", type=["xlsx"])

# Funções auxiliares
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

def gerar_excel(produto, mes, regiao):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        produto.to_excel(writer, sheet_name="Por Produto", index=False)
        mes.to_excel(writer, sheet_name="Por Mês", index=False)
        regiao.to_excel(writer, sheet_name="Por Região", index=False)
    return output.getvalue()

# Processamento
if uploaded_file:
    df = pd.read_excel(uploaded_file)
    colunas_esperadas = {"Data da Venda", "Produto", "Região", "Valor da Venda"}

    if not colunas_esperadas.issubset(set(df.columns)):
        st.error("❌ A planilha deve conter as colunas: Data da Venda, Produto, Região, Valor da Venda.")
    else:
        df = carregar_dados(uploaded_file)
        por_produto_valor, por_mes, por_regiao = gerar_relatorios(df)

        # Análises de produto
        produto_faturamento = por_produto_valor.sort_values("Valor da Venda", ascending=False).iloc[0]

        produto_qtd = df["Produto"].value_counts()
        produto_mais_vendido_qtd = produto_qtd.idxmax()
        qtd_vendas = produto_qtd.max()

        # Destaques principais
        total_geral = df["Valor da Venda"].sum()
        regiao_top = por_regiao.sort_values("Valor da Venda", ascending=False).iloc[0]

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

        # Botão de download
        nome_arquivo = f"relatorio_vendas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label="📥 Baixar Relatório Excel",
            data=gerar_excel(por_produto_valor, por_mes, por_regiao),
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
