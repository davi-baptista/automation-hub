import streamlit as st
import plotly.express as px
import locale

from loader.data_loader import carregar_dados
from functions.functions import (
    aplicar_filtros,
    atualizar_selecoes_empresas,
    atualizar_selecoes_grupos,
)

def show():
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    
    df = carregar_dados(status=1)
    
    todos_grupos = ['Todos os Grupos'] + sorted(df['GRUPO'].unique().tolist())
    grupos = st.sidebar.multiselect('Grupos', todos_grupos, default=['Todos os Grupos'], key='grupos_selecionados', on_change=atualizar_selecoes_grupos)
    
    if 'Todos os Grupos' not in grupos:
        df_alterado = df[df['GRUPO'].isin(grupos)]
    else:
        df_alterado = df.copy()
    
    todas_empresas = ['Todas as Empresas'] + df_alterado['chave_unica'].unique().tolist()
    empresas = st.sidebar.multiselect('Empresa', todas_empresas, default=['Todas as Empresas'], key='empresas_selecionadas', on_change=atualizar_selecoes_empresas)
    
    df_alterado['ANO'] = df_alterado['MES_NOME'].apply(lambda x: x.split(' ')[1])
    todos_anos = sorted(df_alterado['ANO'].unique().tolist())
    
    conditions = []
    if 'Todas as Empresas' not in empresas:
        conditions.append(df_alterado['chave_unica'].isin(empresas))
        
    df_filtered = aplicar_filtros(df_alterado, conditions)
    
    st.title('Comparativo dos Funcionarios')
    
    if df_filtered.empty:
        st.warning('Escolha uma empresa ou grupo de empresas para fazer o comparativo entre elas.')
    
    else:
        col1, col2, col3 = st.columns(3)
        start_month, end_month = col1.select_slider(
            'Selecione a faixa de meses para analise horizontal',
            options=df_filtered['MES_NOME'].unique(),
            value=(df_filtered.iloc[0]['MES_NOME'], df_filtered.iloc[-1]['MES_NOME']))
        
        primeiro_mes = df_filtered[df_filtered['MES_NOME'] == start_month]['QTDEFUNCIONARIOS'].sum()
        ultimo_mes = df_filtered[df_filtered['MES_NOME'] == end_month]['QTDEFUNCIONARIOS'].sum()
        
        var_percentual = ((ultimo_mes - primeiro_mes) / primeiro_mes) * 100 if primeiro_mes != 0 else 0

        delta =  f'{var_percentual:.2f}%'
        
        col3.metric('Analise Horizontal', delta, delta, label_visibility="collapsed")
        st.markdown("<hr>", unsafe_allow_html=True)
        
        for ano in reversed(todos_anos):
            st.title(f'Funcionarios ({ano})')
            metric, graphic = st.columns(2)
            df_ano_especifico = df_filtered[df_filtered['ANO'] == ano]
            
            func_total_ano = df_ano_especifico.groupby(['MES', 'MES_NOME'])[['QTDEFUNCIONARIOS']].sum().reset_index()
            func_total_ano = func_total_ano.sort_values('MES', ascending=True)

            with metric:
                ultimo_mes = func_total_ano.iloc[-1]['QTDEFUNCIONARIOS']
                faturamento_total_ano = func_total_ano['QTDEFUNCIONARIOS'].sum()
                
                st.metric(f'Quantidade de Funcionarios no Final do Ano', f'{locale.format_string("%0.0f", ultimo_mes, grouping=True)}')
                
                media_total_ano = faturamento_total_ano / len(func_total_ano['MES'].unique())
                st.metric(f'Media ao mês', f'{locale.format_string("%0.2f", media_total_ano, grouping=True)}')
                
            with graphic:
                fig_qtd_funcionarios = px.line(
                    func_total_ano,
                    x='MES_NOME',
                    y='QTDEFUNCIONARIOS',
                    text='QTDEFUNCIONARIOS',
                    title='Total de Funcionários por mês',
                    color_discrete_sequence=['#FF6245']
                )
                fig_qtd_funcionarios.update_traces(
                    hovertemplate='Mês: %{x}<br>Total de Funcionários: %{y}',
                    mode='lines+markers',  # Adiciona linha e marcadores
                    marker=dict(size=10),  # Ajusta o tamanho dos marcadores para torná-los mais visíveis
                )
                fig_qtd_funcionarios.update_layout(
                    xaxis_title='Mês', 
                    yaxis_title='Funcionários',
                    height=500
                )
                # Adicionando anotações de texto para cada ponto
                for x, y in zip(func_total_ano['MES_NOME'], func_total_ano['QTDEFUNCIONARIOS']):
                    fig_qtd_funcionarios.add_annotation(
                        x=x, 
                        y=y,
                        text=locale.format_string('%d', y, grouping=True),
                        showarrow=False,
                        xanchor='center',  # Centraliza o texto horizontalmente em relação ao ponto
                        yanchor='bottom',  # Ancora o texto abaixo do ponto (você pode usar 'top' para ancorar acima)
                        yshift=5
                    )
                
                st.plotly_chart(fig_qtd_funcionarios, use_container_width=True)
                
            st.markdown("<hr>", unsafe_allow_html=True)