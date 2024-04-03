import streamlit as st
import plotly.express as px
import locale
import pandas as pd

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
    
    checkbox_entrada = st.sidebar.checkbox("Notas de Entrada", True, key='checkbox1')
    checkbox_saida = st.sidebar.checkbox("Notas de Saída", True, key='checkbox2')
    checkbox_tomado = st.sidebar.checkbox("Serviço Tomado", True, key='checkbox3')
    checkbox_prestado = st.sidebar.checkbox("Serviço Prestado", True, key='checkbox4')
    
    tipos_de_notas_selecionados = []
    if checkbox_entrada:
        tipos_de_notas_selecionados.append('NOTASENTRADAS')
    if checkbox_saida:
        tipos_de_notas_selecionados.append('NOTASSAIDAS')
    if checkbox_tomado:
        tipos_de_notas_selecionados.append('NOTASSERVTOMADOS')
    if checkbox_prestado:
        tipos_de_notas_selecionados.append('NOTASSERVPRESTADOS')
    
    if not tipos_de_notas_selecionados:
        tipos_de_notas_selecionados.append('NOTASENTRADAS')
        tipos_de_notas_selecionados.append('NOTASSAIDAS')
        tipos_de_notas_selecionados.append('NOTASSERVTOMADOS')
        tipos_de_notas_selecionados.append('NOTASSERVPRESTADOS')
        
    df_alterado['ANO'] = df_alterado['MES_NOME'].apply(lambda x: x.split(' ')[1])
    todos_anos = sorted(df_alterado['ANO'].unique().tolist())
    
    conditions = []
    if 'Todas as Empresas' not in empresas:
        conditions.append(df_alterado['chave_unica'].isin(empresas))
        
    df_filtered = aplicar_filtros(df_alterado, conditions)
    
    st.title('Comparativo das Notas')
    
    if df_filtered.empty:
        st.warning('Escolha uma empresa ou grupo de empresas para fazer o comparativo entre elas.')
    
    else:
        col1, col2, col3 = st.columns(3)
        start_month, end_month = col1.select_slider(
            'Selecione a faixa de meses para analise horizontal',
            options=df_filtered['MES_NOME'].unique(),
            value=(df_filtered.iloc[0]['MES_NOME'], df_filtered.iloc[-1]['MES_NOME']))
        
        primeiro_mes = df_filtered.loc[df_filtered['MES_NOME'] == start_month, tipos_de_notas_selecionados].sum().sum()
        ultimo_mes = df_filtered.loc[df_filtered['MES_NOME'] == end_month, tipos_de_notas_selecionados].sum().sum()
        
        var_percentual = ((ultimo_mes - primeiro_mes) / primeiro_mes) * 100 if primeiro_mes != 0 else 0

        delta =  f'{var_percentual:.2f}%'
        
        col3.metric('Analise Horizontal', delta, delta, label_visibility="collapsed")
        st.markdown("<hr>", unsafe_allow_html=True)
        
        for ano in reversed(todos_anos):
            st.title(f'Quantidade de Notas ({ano})')
            metric, graphic = st.columns(2)
            df_ano_especifico = df_filtered[df_filtered['ANO'] == ano]
            
            notas_total_ano = df_ano_especifico.groupby(['MES', 'MES_NOME'])[['NOTASENTRADAS', 'NOTASSAIDAS', 'NOTASSERVTOMADOS', 'NOTASSERVPRESTADOS']].sum().reset_index()
            notas_total_ano = notas_total_ano.sort_values('MES', ascending=True)
            
            soma_total = notas_total_ano[['NOTASENTRADAS', 'NOTASSAIDAS', 'NOTASSERVTOMADOS', 'NOTASSERVPRESTADOS']].sum().sum()
            
            with metric:
                somas, medias = metric.columns(2)
                
                with somas:
                    st.metric(f'Total notas do ano', f'{locale.format_string("%0.0f", soma_total, grouping=True)}')
                    if checkbox_entrada:
                        soma_entrada = notas_total_ano['NOTASENTRADAS'].sum().sum()
                        st.metric(f'Total notas de entrada do ano', f'{locale.format_string("%0.0f", soma_entrada, grouping=True)}')
                    if checkbox_saida:
                        soma_saida = notas_total_ano['NOTASSAIDAS'].sum().sum()
                        st.metric(f'Total notas de saida do ano', f'{locale.format_string("%0.0f", soma_saida, grouping=True)}')
                    if checkbox_tomado:
                        soma_tomado = notas_total_ano['NOTASSERVTOMADOS'].sum().sum()
                        st.metric(f'Total notas serviços tomados do ano', f'{locale.format_string("%0.0f", soma_tomado, grouping=True)}')
                    if checkbox_prestado:
                        soma_prestado = notas_total_ano['NOTASSERVPRESTADOS'].sum().sum()
                        st.metric(f'Total notas serviços prestados do ano', f'{locale.format_string("%0.0f", soma_prestado, grouping=True)}')
                
                with medias:
                    qtd_meses = len(notas_total_ano['MES'].unique())
                    st.metric(f'Media ao mês', f'{locale.format_string("%0.2f", soma_total / qtd_meses, grouping=True)}')
                    if checkbox_entrada:
                        st.metric(f'Media ao mês', f'{locale.format_string("%0.2f", soma_entrada / qtd_meses, grouping=True)}')
                    if checkbox_saida:
                        st.metric(f'Media ao mês', f'{locale.format_string("%0.2f", soma_saida / qtd_meses, grouping=True)}')
                    if checkbox_tomado:
                        st.metric(f'Media ao mês', f'{locale.format_string("%0.2f", soma_tomado / qtd_meses, grouping=True)}')
                    if checkbox_prestado:
                        st.metric(f'Media ao mês', f'{locale.format_string("%0.2f", soma_prestado / qtd_meses, grouping=True)}')
                
            with graphic:
                df_total_por_categoria = pd.melt(notas_total_ano, id_vars=['MES_NOME'], value_vars=tipos_de_notas_selecionados, var_name='Categoria', value_name='Valor')
                
                color_map = {
                    'Notas Saída': '#D96704',
                    'Serviços Prestados': '#F2D377',
                    'Notas Entrada': '#F2836B',
                    'Serviços Tomados': '#F24C3D'
                }
                novos_rotulos = {
                    'NOTASENTRADAS': 'Notas Entrada',
                    'NOTASSAIDAS': 'Notas Saída',
                    'NOTASSERVTOMADOS': 'Serviços Tomados',
                    'NOTASSERVPRESTADOS': 'Serviços Prestados'
                }
                df_total_por_categoria['Categoria'] = df_total_por_categoria['Categoria'].map(novos_rotulos)
                
                textos_para_barras = []
                qtd_meses = len(df_total_por_categoria['MES_NOME'].unique())
                qtd_categorias = len(df_total_por_categoria['Categoria'].unique())
                
                textos_para_barras += [''] * (qtd_meses*(qtd_categorias-1))
    
                for mes in df_total_por_categoria['MES_NOME'].unique():
                    total = df_total_por_categoria[df_total_por_categoria['MES_NOME']==mes]['Valor'].sum()
                    textos_para_barras += [f"{total:,.0f}".replace(',', '.')]
        
                fig_receita = px.bar(
                    df_total_por_categoria,
                    y='Valor',
                    x='MES_NOME',
                    text=textos_para_barras,
                    color='Categoria',
                    title='Total de Notas por mês',
                    color_discrete_map=color_map
                )
                fig_receita.update_traces(
                    hovertemplate='Mês: %{x}<br>Quantidade de notas: %{y}',
                    textposition='outside'
                )
                fig_receita.update_layout(
                    yaxis_title='Mês',
                    xaxis_title='Notas Totais'
                )
                
                st.plotly_chart(fig_receita, use_container_width=True)
                
            st.markdown("<hr>", unsafe_allow_html=True)