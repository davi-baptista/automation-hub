import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import locale

from loader.data_loader import carregar_dados
from functions.functions import (
    aplicar_filtros,
    criar_sessoes_state,
    atualizar_selecoes_empresas,
    atualizar_selecoes_grupos,
    atualizar_selecoes_anos,
    atualizar_selecoes_meses,
)

def show():
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    st.title('Dados Gerais')
    df = carregar_dados(status=1)
    criar_sessoes_state()
        
    todos_grupos = ['Todos os Grupos'] + sorted(df['GRUPO'].unique().tolist())
    grupos = st.sidebar.multiselect(
        'Grupos',
        todos_grupos,
        default=['Todos os Grupos'],
        key='grupos_selecionados',
        on_change=atualizar_selecoes_grupos
    )
    
    if 'Todos os Grupos' not in grupos:
        df_alterado = df[df['GRUPO'].isin(grupos)]
    else:
        df_alterado = df.copy()

    todas_empresas = ['Todas as Empresas'] + sorted(df_alterado['chave_unica'].unique().tolist())
    empresas = st.sidebar.multiselect(
        'Empresa',
        todas_empresas,
        default=['Todas as Empresas'],
        key='empresas_selecionadas',
        on_change=atualizar_selecoes_empresas
    )
        
    df_alterado['ANO'] = df_alterado['MES_NOME'].apply(lambda x: x.split(' ')[1])
    todos_anos = ['Todos os Anos'] + sorted(df_alterado['ANO'].unique().tolist())
    anos = st.sidebar.multiselect(
        'Ano',
        todos_anos,
        default=['Todos os Anos'],
        key='anos_selecionados',
        on_change=atualizar_selecoes_anos
    )
        
    df_alterado['MES_STR'] = df_alterado['MES_NOME'].apply(lambda x: x.split(' ')[0])
    todos_meses = [
        'Todos os Meses', 'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
    ]
    meses = st.sidebar.multiselect(
        'Mês',
        todos_meses,
        default=['Todos os Meses'],
        key='meses_selecionados',
        on_change=atualizar_selecoes_meses
    )

    tipos_de_notas_selecionados = create_checkboxes()

    conditions = []
    if 'Todas as Empresas' not in empresas:
        conditions.append(df_alterado['chave_unica'].isin(empresas))
    if 'Todos os Anos' not in anos:
        conditions.append(df_alterado['ANO'].isin(anos))
    if 'Todos os Meses' not in meses:
        conditions.append(df_alterado['MES_STR'].isin(meses))
        
    df_filtered = aplicar_filtros(df_alterado, conditions)

    if df_filtered.empty:
        st.warning('Nenhum dado encontrado com os filtros aplicados.')
    else:
        with st.container():
            (
                media_honorarios,
                media_lancamentos,
                delta_honorarios,
                delta_lancamentos
            ) = calculate_averages(df_filtered)
            display_metrics(
                media_honorarios,
                media_lancamentos,
                delta_honorarios,
                delta_lancamentos
            )
            
        st.markdown("<hr>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2, gap='large')
        col3, col4 = st.columns(2, gap='large')
        col5, col6 = st.columns(2, gap='large')
        col7, col8 = st.columns(2, gap='large')


        func_total = df_filtered.groupby(['MES', 'MES_NOME'])[['QTDEFUNCIONARIOS']].sum().reset_index()
        
        fig_qtd_funcionarios = px.line(
            func_total,
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
        for x, y in zip(func_total['MES_NOME'], func_total['QTDEFUNCIONARIOS']):
            fig_qtd_funcionarios.add_annotation(
                x=x, 
                y=y,
                text=locale.format_string('%d', y, grouping=True),
                showarrow=False,
                xanchor='center',  # Centraliza o texto horizontalmente em relação ao ponto
                yanchor='bottom',  # Ancora o texto abaixo do ponto (você pode usar 'top' para ancorar acima)
                yshift=5
            )
            
        col1.plotly_chart(fig_qtd_funcionarios, use_container_width=True)


        grouped_total = df_filtered.groupby(['MES', 'MES_NOME'])[['ADMISSOES', 'DEMISSOES']].sum().reset_index()
        df_melted = pd.melt(grouped_total, id_vars=['MES_NOME'], value_vars=['ADMISSOES', 'DEMISSOES'], var_name='Categoria', value_name='Valor')
        
        df_melted['ValorFormatado'] = df_melted['Valor'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
        
        fig_grouped = px.bar(
            df_melted,
            x='MES_NOME',
            y='Valor',
            text='ValorFormatado',
            color='Categoria',
            barmode='group',
            title='Admissões e Demissões totais por Mês',
            color_discrete_sequence=['#D96704', '#F2D377']
        )
        fig_grouped.update_traces(
            hovertemplate='Mês: %{x}<br>Categoria: %{label}<br>Valor: %{y}',
            texttemplate='%{y:,.0f}',
            textposition='outside'
        )
        fig_grouped.update_layout(
            xaxis_title='Mês',
            yaxis_title='Admissões e Demissões',
            height=500
        )
        
        col2.plotly_chart(fig_grouped, use_container_width=True)
        

        notas_total = df_filtered.groupby(['MES', 'MES_NOME'])[tipos_de_notas_selecionados].sum().reset_index()
        notas_total['Total_Notas'] = notas_total[tipos_de_notas_selecionados].sum(axis=1)
        
        notas_total['ValorFormatado'] = notas_total['Total_Notas'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))

        # Criar gráfico de barras para o total de notas por mês
        fig_notas_total = px.bar(
            notas_total,
            x='MES_NOME',
            y='Total_Notas',
            text='ValorFormatado',
            title='Quantidade Total de Notas por mês',
            color_discrete_sequence=['#FF6245']
        )
        fig_notas_total.update_traces(
            hovertemplate='Mês: %{x}<br>Total de Notas: %{y}',
            textposition='outside'
        )
        fig_notas_total.update_layout(
            xaxis_title='Mês',
            yaxis_title='Notas',
            height=500
        )
        
        col3.plotly_chart(fig_notas_total, use_container_width=True)

        
        total_por_categoria = notas_total[tipos_de_notas_selecionados].sum()
        
        df_total_por_categoria = pd.DataFrame({
            'Categoria': total_por_categoria.index,
            'Valor': total_por_categoria.values
        })
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
        fig_qtd_notas = px.pie(
            df_total_por_categoria,
            values='Valor',
            names='Categoria',
            title='Tipos de Notas',
            color='Categoria',
            color_discrete_map=color_map
        )
        fig_qtd_notas.update_traces(
            hovertemplate='Categoria: %{label}<br>Valor: %{value} (%{percent})',
            texttemplate='%{label}<br>%{percent}'
        )
        fig_qtd_notas.update_layout(
            height=500
        )
        
        col4.plotly_chart(fig_qtd_notas, use_container_width=True)


        lanc_total = df_filtered.groupby(['MES', 'MES_NOME'])[['REGISTROSCONTABEIS']].sum().reset_index()
        
        fig_lancamentos = px.line(
            lanc_total,
            x='MES_NOME',
            y='REGISTROSCONTABEIS',
            title='Total de Lançamentos por mês',
            color_discrete_sequence=['#F2D377']
            )
        fig_lancamentos.update_traces(
            hovertemplate='Mês: %{x}<br>Total de Lançamentos: %{y}',
            mode='lines+markers',  # Adiciona linha e marcadores
            marker=dict(size=10),  # Ajusta o tamanho dos marcadores para torná-los mais visíveis
        )
        fig_lancamentos.update_layout(
            xaxis_title='Mês', 
            yaxis_title='Lançamentos',
            height=500
        )
        # Adicionando anotações de texto para cada ponto
        for x, y in zip(lanc_total['MES_NOME'], lanc_total['REGISTROSCONTABEIS']):
            fig_lancamentos.add_annotation(
                x=x,
                y=y,
                text=locale.format_string('%d', y, grouping=True),
                showarrow=False,
                xanchor='center',  # Centraliza o texto horizontalmente em relação ao ponto
                yanchor='bottom',  # Ancora o texto abaixo do ponto (você pode usar 'top' para ancorar acima)
                yshift=5
            )
            
        col5.plotly_chart(fig_lancamentos, use_container_width=True)


        receita_total = df_filtered.groupby(['MES', 'MES_NOME'])[['RECEITABRUTA']].sum().reset_index()
        # Reordene os dados do DataFrame de receita total para que os meses sejam exibidos em ordem reversa
        receita_total = receita_total.sort_values('MES', ascending=False)
        
        receita_total['RECEITABRUTA_FORMATADA'] = receita_total['RECEITABRUTA'].apply(lambda x: locale.format_string('%.2f', x, grouping=True))
        
        fig_receita = px.bar(
            receita_total,
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
        
        col6.plotly_chart(fig_receita, use_container_width=True)


        hono_total = df_filtered.groupby(['MES', 'MES_NOME'])[['HONORARIOS']].sum().reset_index()
        hono_total['HONORARIOS_FORMATADO'] = hono_total['HONORARIOS'].apply(lambda x: locale.format_string('%.2f', x, grouping=True))

        hono_pagos_total = get_unique_fees(df_filtered)
        pagos_total = hono_pagos_total[['HONORARIOS_PAGOS']].sum().sum()
        
        lista_pagos_total = [pagos_total] * len(hono_total['MES'].unique())
        
        fig_honorarios = px.bar(
            hono_total,
            x='MES_NOME',
            y='HONORARIOS',
            text='HONORARIOS_FORMATADO',
            title='Total de Honorarios por mês',
            color_discrete_sequence=['#FF6245']
        )
        fig_honorarios.update_traces(
            hovertemplate='Mês: %{x}<br>Total de Honorarios: R$ %{y:,.2f}',
            textposition='outside'
        )
        fig_honorarios.update_layout(
            xaxis_title='Mês', 
            yaxis_title='Honorarios',
            height=500
        )
        # Adiciona uma linha para os honorários pagos
        fig_honorarios.add_trace(
            go.Scatter(x=hono_total['MES_NOME'], y=lista_pagos_total, mode='lines',
                    name=f'Honorário Pago<br> R$ {locale.format_string('%.2f', pagos_total, grouping=True)}', line=dict(color='#F2D377'))
        )
            
        col7.plotly_chart(fig_honorarios, use_container_width=True)
        
        col1.markdown("<hr>", unsafe_allow_html=True)
        col2.markdown("<hr>", unsafe_allow_html=True)
        col3.markdown("<hr>", unsafe_allow_html=True)
        col4.markdown("<hr>", unsafe_allow_html=True)
        col5.markdown("<hr>", unsafe_allow_html=True)
        col6.markdown("<hr>", unsafe_allow_html=True)
        col7.markdown("<hr>", unsafe_allow_html=True)
        
        
def create_checkboxes():
     # Seleção de tipos de notas
    selecao_checkbox1 = st.sidebar.checkbox("Notas de Saída", True, key='checkbox1')
    selecao_checkbox2 = st.sidebar.checkbox("Serviço Prestado", True, key='checkbox2')
    selecao_checkbox3 = st.sidebar.checkbox("Notas de Entrada", True, key='checkbox3')
    selecao_checkbox4 = st.sidebar.checkbox("Serviço Tomado", True, key='checkbox4')

    tipos_de_notas_selecionados = []
    if selecao_checkbox1:
        tipos_de_notas_selecionados.append('NOTASSAIDAS')
    if selecao_checkbox2:
        tipos_de_notas_selecionados.append('NOTASSERVPRESTADOS')
    if selecao_checkbox3:
        tipos_de_notas_selecionados.append('NOTASENTRADAS')
    if selecao_checkbox4:
        tipos_de_notas_selecionados.append('NOTASSERVTOMADOS')
        
    return tipos_de_notas_selecionados

    
def calculate_averages(df_filtered):
    total_meses = len(df_filtered['MES'].unique())
    if total_meses == 0:
        return 0, 0, 0, 0, None, None, None, None, None, None
    
    grouped_data = df_filtered.groupby('MES')['HONORARIOS'].sum().reset_index()
    grouped_data = grouped_data.sort_values('MES')
    media_honorarios = grouped_data['HONORARIOS'].sum() / total_meses
    first_month_honorarios = grouped_data.iloc[0]['HONORARIOS']
    last_month_honorarios = grouped_data.iloc[-1]['HONORARIOS']
    delta_honorarios = round(((last_month_honorarios - first_month_honorarios) / first_month_honorarios) * 100, 2) if first_month_honorarios != 0 else 0

    df_cnpj_filtered = df_filtered.drop_duplicates(subset=['CNPJ'], keep='first')
    hono_pago_total = df_cnpj_filtered['HONORARIOS_PAGOS'].sum().sum()
    delta_honorarios_pagos = round(((hono_pago_total - media_honorarios) / media_honorarios) * 100, 2) if media_honorarios != 0 else 0
    
    return media_honorarios, hono_pago_total, delta_honorarios, delta_honorarios_pagos


def display_metrics(media_honorarios, hono_pago_total, delta_honorarios, delta_honorarios_pagos):
    mediacol1, mediacol2, mediacol3 = st.columns(3)
    with mediacol1:
        st.metric('Média dos Honorários', f"R$ {locale.format_string("%0.2f", media_honorarios, grouping=True)}", f'{delta_honorarios}%')
    with mediacol2:
        st.metric('Último Honorário Pago', locale.format_string("%0.2f", hono_pago_total, grouping=True))
    with mediacol3:
        st.metric('Diferença dos Honorários', locale.format_string("%0.2f", hono_pago_total - media_honorarios, grouping=True), f'{delta_honorarios_pagos}%')
        
def get_unique_fees(df):
    # Primeiro, garantir que o DataFrame está ordenado por CNPJ (e por data, se aplicável)
    df = df.sort_values(by=['CNPJ', 'MES'], ascending=True)

    # Use cumcount() para contar as ocorrências de cada CNPJ
    df['cnpj_ocorrencia'] = df.groupby('CNPJ').cumcount()

    # Aplique uma condição para zerar 'HONORARIOS_PAGOS' exceto para a primeira ocorrência de cada CNPJ
    df['HONORARIOS_PAGOS'] = df.apply(lambda x: x['HONORARIOS_PAGOS'] if x['cnpj_ocorrencia'] == 0 else 0, axis=1)

    # Remover a coluna auxiliar 'cnpj_ocorrencia', pois ela já cumpriu seu propósito
    df = df.drop(columns=['cnpj_ocorrencia'])
    return df