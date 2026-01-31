import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import plotly.express as px

arquivo = r"C:\Users\tavin\Downloads\RELATORIO VEÍCULOS CARREGADOS E NÃO CARREGADOS.xlsx"

def analise_dinamica():
    dados = pd.read_excel(arquivo, sheet_name="DINAMICA", header=2)

    print("Colunas lidas: ", dados.columns)

    dados = dados.rename(columns={
        "Rótulos de Linha": "Operacao",
        "VALOR": "Valor",
        "QUANT.CARGA": "Quantidade"
    })

    if "Operacao" in dados.columns:
        dados = dados.dropna(subset=["Operacao"])
    else:
        print("Coluna 'Operacao' não encontrada. Verifique o nome exato na planilha Excel.")

    total = dados["Valor"].sum()
    dados["Impacto_%"] = (dados["Valor"] / total) * 100
    dados = dados.sort_values(by="Impacto_%", ascending=False)

    return dados, total

tabela, total_prejuizo = analise_dinamica()
tabela.to_excel("analise_dinamica.xlsx", index=False)
print("Arquivo 'Analise_dinamica.xlsx' gerado com sucesso!")

tabela_limpa = tabela.dropna(subset=["Operacao", "Valor"])
tabela_limpa = tabela_limpa[tabela_limpa["Valor"] > 0]


st.title("📊 Análise de Impacto Operacional")
st.metric("💰 Total de Prejuízo", f"R$ {total_prejuizo:,.2f}")

st.subheader("Tabela de Impacto Por Operação")
st.dataframe(tabela)

st.subheader("Gráfico de Impacto (%)")
st.bar_chart(tabela.set_index("Operacao")["Impacto_%"])

st.subheader("Distribuição do Prejuízo (Pizza)")
fig = px.pie(
    tabela,
    values="Valor",
    names="Operacao",
    title="Distribuição do Prejuízo",
    hole=0.3
)
fig.update_traces(textinfo="percent+label", textposition="inside")
st.plotly_chart(fig)
