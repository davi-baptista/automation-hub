import streamlit as st
import page.dados_gerais as dados_gerais, page.comparativo_faturamento as comparativo_faturamento, page.comparativo_funcionarios as comparativo_funcionarios, page.comparativo_notas as comparativo_notas

st.set_page_config(page_title="BI Customer Success", layout='wide')
# Define as páginas do aplicativo
        
paginas = {
    "Dados Gerais": dados_gerais,
    "Comparativo Faturamento": comparativo_faturamento,
    "Comparativo Funcionarios": comparativo_funcionarios,
    "Comparativo Notas": comparativo_notas
}

# Adiciona um seletor no sidebar para escolha da página
pagina_selecionada = st.sidebar.selectbox("Escolha uma página", list(paginas.keys()))

# Chama a função show() da página selecionada
paginas[pagina_selecionada].show()