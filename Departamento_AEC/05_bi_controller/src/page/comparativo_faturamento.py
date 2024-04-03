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
    
    st.title('Comparativo do Faturamento')
    
    if df_filtered.empty:
        st.warning('Escolha uma empresa ou grupo de empresas para fazer o comparativo entre elas.')
    
    else:
        col1, col2, col3 = st.columns(3)
        start_month, end_month = col1.select_slider(
            'Selecione a faixa de meses para analise horizontal',
            options=df_filtered['MES_NOME'].unique(),
            value=(df_filtered.iloc[0]['MES_NOME'], df_filtered.iloc[-1]['MES_NOME']))
        
        primeiro_mes = df_filtered[df_filtered['MES_NOME'] == start_month]['RECEITABRUTA'].sum()
        ultimo_mes = df_filtered[df_filtered['MES_NOME'] == end_month]['RECEITABRUTA'].sum()
        
        var_percentual = ((ultimo_mes - primeiro_mes) / primeiro_mes) * 100 if primeiro_mes != 0 else 0

        delta =  f'{var_percentual:.2f}%'
        
        col3.metric('Analise Horizontal', delta, delta, label_visibility="collapsed")
        st.markdown("<hr>", unsafe_allow_html=True)
        
        for ano in reversed(todos_anos):
            st.title(f'Faturamento ({ano})')
            metric, graphic = st.columns(2)
            df_ano_especifico = df_filtered[df_filtered['ANO'] == ano]
            receita_total_ano = df_ano_especifico.groupby(['MES', 'MES_NOME'])[['RECEITABRUTA']].sum().reset_index()
            receita_total_ano = receita_total_ano.sort_values('MES', ascending=True)

            with metric:
                faturamento_total_ano = receita_total_ano['RECEITABRUTA'].sum()
                
                st.metric(f'Total do ano', f'R$ {locale.format_string("%0.2f", faturamento_total_ano, grouping=True)}')
                
                media_total_ano = faturamento_total_ano / len(receita_total_ano['MES'].unique())
                st.metric(f'Media ao mês', f'R$ {locale.format_string("%0.2f", media_total_ano, grouping=True)}')
            with graphic:

                receita_total_ano['RECEITABRUTA_FORMATADA'] = receita_total_ano['RECEITABRUTA'].apply(lambda x: locale.format_string('%.2f', x, grouping=True))
                
                # Reordene os dados do DataFrame de receita total para que os meses sejam exibidos em ordem reversa
                receita_total_ano = receita_total_ano.sort_values('MES', ascending=False)
                
                fig_receita = px.bar(
                    receita_total_ano,
                    y='MES_NOME',
                    x='RECEITABRUTA',
                    text='RECEITABRUTA_FORMATADA',
                    title='Total de Receita Bruta por mês',
                    color_discrete_sequence=['#FF6245']
                )
                fig_receita.update_traces(
                    hovertemplate='Mês: %{y}<br>Receita Bruta: R$ %{x:,.2f}',
                    texttemplate='R$ %{text}',
                    textfont=dict(
                        color='white'
                    ),
                    textposition='inside'
                )
                fig_receita.update_layout(
                    yaxis_title='Mês',
                    xaxis_title='Receita Bruta',
                    height=500
                )
                
                st.plotly_chart(fig_receita, use_container_width=True)
                
            st.markdown("<hr>", unsafe_allow_html=True)