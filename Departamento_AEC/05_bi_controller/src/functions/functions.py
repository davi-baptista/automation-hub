import streamlit as st

def aplicar_filtros(df, conditions):
    if conditions:
        condition = conditions.pop(0)
        for item in conditions:
            condition = condition & item
        df_filtered = df.loc[condition]
    else:
        df_filtered = df

    return df_filtered

def criar_sessoes_state():
    global selecionadas_empresas, selecionadas_grupos, selecionadas_anos, selecionadas_meses
    
    if 'grupos_selecionados' not in st.session_state:
        st.session_state['grupos_selecionados'] = ['Todos os Grupos']
        
    if 'empresas_selecionadas' not in st.session_state:
        st.session_state['empresas_selecionadas'] = ['Todas as Empresas']
        
    if 'anos_selecionados' not in st.session_state:
        st.session_state['anos_selecionados'] = ['Todos os Anos']
    
    if 'meses_selecionados' not in st.session_state:
        st.session_state['meses_selecionados'] = ['Todos os Meses']
    
def atualizar_selecoes_empresas():
    if 'empresas_selecionadas' not in st.session_state:
        st.session_state['empresas_selecionadas'] = ['Todas as Empresas']
        
    selecionadas_empresas = st.session_state.empresas_selecionadas
    if 'Todas as Empresas' in selecionadas_empresas:
        if len(selecionadas_empresas) > 1:  # Se outras empresas foram selecionadas junto com "Todas as Empresas"
            if selecionadas_empresas[-1] == 'Todas as Empresas':  # Se a última seleção foi "Todas as Empresas"
                st.session_state.empresas_selecionadas = ['Todas as Empresas']  # Mantém só "Todas as Empresas"
            else:
                # Remove "Todas as Empresas" da seleção, mantendo as demais
                st.session_state.empresas_selecionadas = [emp for emp in selecionadas_empresas if emp != 'Todas as Empresas']
                
def atualizar_selecoes_grupos():
    if 'grupos_selecionados' not in st.session_state:
        st.session_state['grupos_selecionados'] = ['Todos os Grupos']
        
    selecionadas_grupos = st.session_state.grupos_selecionados
    if 'Todos os Grupos' in selecionadas_grupos:
        if len(selecionadas_grupos) > 1:
            if selecionadas_grupos[-1] == 'Todos os Grupos':
                st.session_state.grupos_selecionados = ['Todos os Grupos']
            else:
                st.session_state.grupos_selecionados = [grp for grp in selecionadas_grupos if grp != 'Todos os Grupos']
    
def atualizar_selecoes_anos():
    if 'anos_selecionados' not in st.session_state:
        st.session_state['anos_selecionados'] = ['Todos os Anos']
        
    selecionadas_anos = st.session_state.anos_selecionados
    if 'Todos os Anos' in selecionadas_anos:
        if len(selecionadas_anos) > 1:
            if selecionadas_anos[-1] == 'Todos os Anos':
                st.session_state.anos_selecionados = ['Todos os Anos']
            else:
                st.session_state.anos_selecionados = [emp for emp in selecionadas_anos if emp != 'Todos os Anos']
    
def atualizar_selecoes_meses():
    if 'meses_selecionados' not in st.session_state:
        st.session_state['meses_selecionados'] = ['Todos os Meses']
        
    selecionadas_meses = st.session_state.meses_selecionados
    if 'Todos os Meses' in selecionadas_meses:
        if len(selecionadas_meses) > 1:
            if selecionadas_meses[-1] == 'Todos os Meses':
                st.session_state.meses_selecionados = ['Todos os Meses']
            else:
                st.session_state.meses_selecionados = [emp for emp in selecionadas_meses if emp != 'Todos os Meses']