import base64
import io
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from fpdf import FPDF


# Função para carregar dados
def load_data(uploaded_file):
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            return df
        except Exception as e:
            st.error(f"Erro ao carregar o arquivo: {e}")
            return None
    return None

# Função para processar contas a receber (planilha original)
def process_receivables(df_receivables):
    # Remover caracteres não numéricos (como 'R$', '.', ',', etc.)
    df_receivables['Valor'] = df_receivables['Valor'].replace({'R\$': '', '\.': '', ',': '.'}, regex=True)
    
    # Converter a coluna para float
    df_receivables['Valor'] = df_receivables['Valor'].astype(float)
    
    # Converter as colunas de data para datetime
    if 'Pagamento' in df_receivables.columns:
        df_receivables['Pagamento'] = pd.to_datetime(df_receivables['Pagamento'], format='%d/%m/%Y')
    if 'Data' in df_receivables.columns:
        df_receivables['Data'] = pd.to_datetime(df_receivables['Data'], format='%d/%m/%Y')
    
    # Converter a coluna 'Unidade' para string
    df_receivables['Unidade'] = df_receivables['Unidade'].astype(str)
    
    # Renomear a coluna 'Valor' para 'Recebimentos'
    df_receivables = df_receivables.rename(columns={'Valor': 'Recebimentos'})
    
    return df_receivables

# Função para processar contas a pagar (planilha original)
def process_payables(df_payables):
    # Remover caracteres não numéricos (como 'R$', '.', ',', etc.)
    df_payables['Valor'] = df_payables['Valor'].replace({'R\$': '', '\.': '', ',': '.'}, regex=True)
    
    # Converter a coluna para float
    df_payables['Valor'] = df_payables['Valor'].astype(float)
    
    # Converter as colunas de data para datetime
    if 'Pagamento' in df_payables.columns:
        df_payables['Pagamento'] = pd.to_datetime(df_payables['Pagamento'], format='%d/%m/%Y')
    if 'Data' in df_payables.columns:
        df_payables['Data'] = pd.to_datetime(df_payables['Data'], format='%d/%m/%Y')
    
    # Converter a coluna 'Unidade' para string
    df_payables['Unidade'] = df_payables['Unidade'].astype(str)
    
    # Renomear a coluna 'Valor' para 'Pagamentos'
    df_payables = df_payables.rename(columns={'Valor': 'Pagamentos'})
    
    return df_payables

# Função para processar o Relatório Caixa - Contas a Receber (nova planilha)
def process_cash_report(df_cash_report):
    # Verificar se a planilha é do tipo "RELATÓRIO CAIXA - CONTAS A RECEBER"
    if 'Entrada' in df_cash_report.columns and 'Data' in df_cash_report.columns:
        # Renomear a coluna 'Entrada' para 'Recebimentos'
        df_cash_report = df_cash_report.rename(columns={'Entrada': 'Recebimentos'})
        
        # Remover pontos de milhares e substituir vírgula por ponto
        df_cash_report['Recebimentos'] = (
            df_cash_report['Recebimentos']
            .str.replace('.', '', regex=False)  # Remove pontos de milhares
            .str.replace(',', '.')  # Substitui vírgula por ponto
            .astype(float)  # Converte para float
        )
        
        # Verificar se todos os valores são numéricos após a substituição
        if not df_cash_report['Recebimentos'].apply(lambda x: isinstance(x, (int, float))).all():
            st.error("A coluna 'Recebimentos' contém valores não numéricos.")
            return None
        
        # Converter a coluna 'Data' para datetime
        df_cash_report['Data'] = pd.to_datetime(df_cash_report['Data'], format='%d/%m/%Y')
        
        # Adicionar uma coluna 'Conta Analítica' se não existir
        if 'Conta Analítica' not in df_cash_report.columns:
            df_cash_report['Conta Analítica'] = 'Outros'
        
        # Adicionar uma coluna 'Unidade' se não existir
        if 'Unidade' not in df_cash_report.columns:
            df_cash_report['Unidade'] = 'UNIDADE SOMBRIO'  # Ou outra unidade padrão
        
        return df_cash_report
    else:
        st.error("A planilha de Relatório Caixa - Contas a Receber não está no formato esperado.")
        return None

# Adicionar campos de entrada para os saldos das contas bancárias no sidebar
st.sidebar.header("Saldos Bancários Iniciais")
saldo_sombrio = st.sidebar.number_input("Saldo Inicial - Conta Sombrio", value=0.0)
saldo_fumaca = st.sidebar.number_input("Saldo Inicial - Conta Fumaça", value=0.0)

# Função para calcular o fluxo de caixa com saldos iniciais
def calculate_cash_flow(df_receivables, df_payables, df_cash_report, start_date, end_date, regime, saldo_sombrio, saldo_fumaca, unidade_selecionada):
    date_range = pd.date_range(start=start_date, end=end_date, freq='D')
    cash_flow = pd.DataFrame(index=date_range, columns=['Recebimentos', 'Pagamentos', 'Saldo'])
    cash_flow = cash_flow.fillna(0)

    # Definir a coluna de data com base no regime
    coluna_data = 'Pagamento' if regime == 'Caixa' else 'Data'

    # Soma os recebimentos por dia (das duas planilhas de recebimentos)
    for index, row in df_receivables.iterrows():
        if row[coluna_data] in cash_flow.index:
            cash_flow.loc[row[coluna_data], 'Recebimentos'] += row['Recebimentos']

    # Soma os recebimentos da terceira planilha (Relatório Caixa - Contas a Receber)
    for index, row in df_cash_report.iterrows():
        if row['Data'] in cash_flow.index:
            cash_flow.loc[row['Data'], 'Recebimentos'] += row['Recebimentos']

    # Soma os pagamentos por dia
    for index, row in df_payables.iterrows():
        if row[coluna_data] in cash_flow.index:
            cash_flow.loc[row[coluna_data], 'Pagamentos'] += row['Pagamentos']

    # Calcula o saldo diário
    cash_flow['Saldo'] = cash_flow['Recebimentos'] - cash_flow['Pagamentos']

    # Adicionar saldo inicial com base na unidade selecionada
    if unidade_selecionada == "Todas as Unidades":
        saldo_inicial = saldo_sombrio + saldo_fumaca
    elif unidade_selecionada == "UNIDADE SOMBRIO":
        saldo_inicial = saldo_sombrio
    elif unidade_selecionada == "UNIDADE MORRO DA FUMAÇA":
        saldo_inicial = saldo_fumaca
    else:
        saldo_inicial = 0

    # Calcular o saldo acumulado
    cash_flow['Saldo Acumulado'] = cash_flow['Saldo'].cumsum() + saldo_inicial

    return cash_flow



# Função para gerar relatório em PDF
def generate_pdf(cash_flow, recebimentos_por_conta, pagamentos_por_conta, unidade_selecionada, regime):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Título do Relatório
    pdf.cell(200, 10, txt=f"Relatório Financeiro - Unidade: {unidade_selecionada} - Regime: {regime}", ln=True, align="C")
    pdf.ln(10)
    
    # Adicionar tabela de fluxo de caixa
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt="Fluxo de Caixa Diário", ln=True)
    pdf.ln(5)
    
    # Cabeçalho da tabela
    pdf.cell(40, 10, txt="Data", border=1)
    pdf.cell(40, 10, txt="Recebimentos", border=1)
    pdf.cell(40, 10, txt="Pagamentos", border=1)
    pdf.cell(40, 10, txt="Saldo Acumulado", border=1)
    pdf.ln()
    
    # Dados da tabela
    for index, row in cash_flow.iterrows():
        pdf.cell(40, 10, txt=index.strftime('%d/%m/%Y'), border=1)
        pdf.cell(40, 10, txt=f"R$ {row['Recebimentos']:,.2f}", border=1)
        pdf.cell(40, 10, txt=f"R$ {row['Pagamentos']:,.2f}", border=1)
        pdf.cell(40, 10, txt=f"R$ {row['Saldo Acumulado']:,.2f}", border=1)
        pdf.ln()
    
    # Adicionar gráficos de recebimentos e pagamentos por conta analítica
    pdf.ln(10)
    pdf.cell(200, 10, txt="Recebimentos por Conta Analítica", ln=True)
    pdf.ln(5)
    
    for index, row in recebimentos_por_conta.iterrows():
        pdf.cell(100, 10, txt=f"{row['Conta Analítica']}: R$ {row['Recebimentos']:,.2f}", ln=True)
    
    pdf.ln(10)
    pdf.cell(200, 10, txt="Pagamentos por Conta Analítica", ln=True)
    pdf.ln(5)
    
    for index, row in pagamentos_por_conta.iterrows():
        pdf.cell(100, 10, txt=f"{row['Conta Analítica']}: R$ {row['Pagamentos']:,.2f}", ln=True)
    
    # Salvar o PDF
    pdf_output = pdf.output(dest='S')  # Retorna um bytearray
    return pdf_output

# Função para criar um link de download
def create_download_link(val, filename, file_type):
    b64 = base64.b64encode(val)
    return f'<a href="data:application/{file_type};base64,{b64.decode()}" download="{filename}">Baixar {filename}</a>'

# Função principal
def main():
    st.title("Dashboard Financeiro - Fortneer")

    # Upload das planilhas
    st.sidebar.header("Upload das Planilhas")
    uploaded_file_receivables_quitadas = st.sidebar.file_uploader("Carregar Contas a Receber Quitadas", type=["xlsx"])
    uploaded_file_receivables_pendentes = st.sidebar.file_uploader("Carregar Contas a Receber Pendentes", type=["xlsx"])
    uploaded_file_payables_quitadas = st.sidebar.file_uploader("Carregar Contas a Pagar Quitadas", type=["xlsx"])
    uploaded_file_payables_pendentes = st.sidebar.file_uploader("Carregar Contas a Pagar Pendentes", type=["xlsx"])
    uploaded_file_cash_report = st.sidebar.file_uploader("Carregar Relatório Caixa - Contas a Receber", type=["xlsx"])

    if (uploaded_file_receivables_quitadas and uploaded_file_receivables_pendentes and 
        uploaded_file_payables_quitadas and uploaded_file_payables_pendentes and uploaded_file_cash_report):
        
        df_receivables_quitadas = load_data(uploaded_file_receivables_quitadas)
        df_receivables_pendentes = load_data(uploaded_file_receivables_pendentes)
        df_payables_quitadas = load_data(uploaded_file_payables_quitadas)
        df_payables_pendentes = load_data(uploaded_file_payables_pendentes)
        df_cash_report = load_data(uploaded_file_cash_report)

        if (df_receivables_quitadas is not None and df_receivables_pendentes is not None and 
            df_payables_quitadas is not None and df_payables_pendentes is not None and df_cash_report is not None):
            
            # Processar as planilhas de contas a receber
            df_receivables_quitadas = process_receivables(df_receivables_quitadas)
            df_receivables_pendentes = process_receivables(df_receivables_pendentes)
            
            # Agregar as duas planilhas de contas a receber
            df_receivables = pd.concat([df_receivables_quitadas, df_receivables_pendentes], ignore_index=True)
            
            # Processar as planilhas de contas a pagar
            df_payables_quitadas = process_payables(df_payables_quitadas)
            df_payables_pendentes = process_payables(df_payables_pendentes)
            
            # Agregar as duas planilhas de contas a pagar
            df_payables = pd.concat([df_payables_quitadas, df_payables_pendentes], ignore_index=True)
            
            # Processar a planilha de relatório de caixa
            df_cash_report = process_cash_report(df_cash_report)

            # Exibir os dados processados das planilhas de contas a receber e a pagar
            st.subheader("Dados Processados - Contas a Receber")
            st.write(df_receivables)

            st.subheader("Dados Processados - Contas a Pagar")
            st.write(df_payables)

            # Filtro por Unidade
            st.sidebar.header("Filtros")
            unidades_recebimentos = df_receivables['Unidade'].unique()
            unidades_pagamentos = df_payables['Unidade'].unique()
            unidades_cash_report = df_cash_report['Unidade'].unique()
            
            # Garantir que todas as unidades sejam strings e ordená-las
            unidades_recebimentos = [str(u) for u in unidades_recebimentos]
            unidades_pagamentos = [str(u) for u in unidades_pagamentos]
            unidades_cash_report = [str(u) for u in unidades_cash_report]
            todas_unidades = sorted(set(unidades_recebimentos + unidades_pagamentos + unidades_cash_report))
            
            # Adicionar a opção "Todas as Unidades"
            todas_unidades = ["Todas as Unidades"] + todas_unidades
            
            unidade_selecionada = st.sidebar.selectbox(
                "Selecione a Unidade",
                options=todas_unidades
            )

            # Filtrar dados por unidade selecionada
            if unidade_selecionada == "Todas as Unidades":
                df_receivables_filtrado = df_receivables
                df_payables_filtrado = df_payables
                df_cash_report_filtrado = df_cash_report
            else:
                df_receivables_filtrado = df_receivables[df_receivables['Unidade'] == unidade_selecionada]
                df_payables_filtrado = df_payables[df_payables['Unidade'] == unidade_selecionada]
                df_cash_report_filtrado = df_cash_report[df_cash_report['Unidade'] == unidade_selecionada]

            # Selecionar o regime (Competência ou Caixa)
            regime = st.sidebar.selectbox(
                "Selecione o Regime",
                options=["Caixa", "Competência"]
            )

            # Verificar se a coluna correta existe no DataFrame
            coluna_regime = 'Pagamento' if regime == 'Caixa' else 'Data'
            if coluna_regime not in df_receivables_filtrado.columns or coluna_regime not in df_payables_filtrado.columns:
                st.error(f"A coluna '{coluna_regime}' não foi encontrada nos dados. Verifique o arquivo carregado.")
                return

            # Selecionar o período de análise
            st.sidebar.header("Período de Análise")
            start_date = st.sidebar.date_input("Data Inicial", datetime.today())
            end_date = st.sidebar.date_input("Data Final", datetime.today() + timedelta(days=30))

            # Converter start_date e end_date para datetime64[ns]
            start_date = pd.to_datetime(start_date)
            end_date = pd.to_datetime(end_date)

            # Calcular o fluxo de caixa
            cash_flow = calculate_cash_flow(df_receivables_filtrado, df_payables_filtrado, df_cash_report_filtrado, start_date, end_date, regime, saldo_sombrio, saldo_fumaca, unidade_selecionada)

            # Exibir o fluxo de caixa
            st.subheader(f"Fluxo de Caixa Diário - Unidade: {unidade_selecionada} - Regime: {regime}")
            
            # Selecionar o tipo de gráfico
            tipo_grafico = st.selectbox(
                "Selecione o Tipo de Gráfico",
                options=["Linha", "Barras", "Área"]
            )

            # Criar o gráfico conforme a seleção
            if tipo_grafico == "Linha":
                fig_fluxo = px.line(cash_flow, x=cash_flow.index, y=['Recebimentos', 'Pagamentos', 'Saldo Acumulado'],
                                    title=f"Fluxo de Caixa Diário - Unidade: {unidade_selecionada} - Regime: {regime}",
                                    labels={"value": "Valor (R$)", "variable": "Tipo", "index": "Data"})
            elif tipo_grafico == "Barras":
                fig_fluxo = px.bar(cash_flow, x=cash_flow.index, y=['Recebimentos', 'Pagamentos', 'Saldo Acumulado'],
                                   title=f"Fluxo de Caixa Diário - Unidade: {unidade_selecionada} - Regime: {regime}",
                                   labels={"value": "Valor (R$)", "variable": "Tipo", "index": "Data"})
            elif tipo_grafico == "Área":
                fig_fluxo = px.area(cash_flow, x=cash_flow.index, y=['Recebimentos', 'Pagamentos', 'Saldo Acumulado'],
                                    title=f"Fluxo de Caixa Diário - Unidade: {unidade_selecionada} - Regime: {regime}",
                                    labels={"value": "Valor (R$)", "variable": "Tipo", "index": "Data"})

            # Exibir o gráfico
            st.plotly_chart(fig_fluxo)

            # Adicionar um filtro de dia específico
            dia_especifico = st.date_input(
                "Selecione um dia para visualizar os valores",
                min_value=start_date,
                max_value=end_date,
                value=start_date
            )

            # Converter o dia_especifico para o mesmo formato do índice do cash_flow
            dia_especifico = pd.to_datetime(dia_especifico)

            # Filtrar os dados para o dia específico
            if dia_especifico in cash_flow.index:
                dados_dia = cash_flow.loc[cash_flow.index == dia_especifico]
                st.write(f"Valores para o dia {dia_especifico.strftime('%d/%m/%Y')}:")
                st.write(dados_dia)
            else:
                st.write(f"Não há dados disponíveis para o dia {dia_especifico.strftime('%d/%m/%Y')}.")

            # Análise de Contas Analíticas
            st.subheader(f"Análise de Contas Analíticas - Unidade: {unidade_selecionada} - Regime: {regime}")
            
            # Gráfico de Pizza - Recebimentos por Conta Analítica
            recebimentos_por_conta = df_receivables_filtrado.groupby('Conta Analítica')['Recebimentos'].sum().reset_index()
            fig_pizza_recebimentos = px.pie(recebimentos_por_conta, values='Recebimentos', names='Conta Analítica',
                                            title=f"Distribuição dos Recebimentos por Conta Analítica - Unidade: {unidade_selecionada}",
                                            hover_data=['Recebimentos'])
            st.plotly_chart(fig_pizza_recebimentos)

            # Gráfico de Pizza - Pagamentos por Conta Analítica
            pagamentos_por_conta = df_payables_filtrado.groupby('Conta Analítica')['Pagamentos'].sum().reset_index()
            fig_pizza_pagamentos = px.pie(pagamentos_por_conta, values='Pagamentos', names='Conta Analítica',
                                          title=f"Distribuição dos Pagamentos por Conta Analítica - Unidade: {unidade_selecionada}",
                                          hover_data=['Pagamentos'])
            st.plotly_chart(fig_pizza_pagamentos)

            # Indicadores Financeiros
            st.subheader(f"Indicadores Financeiros - Unidade: {unidade_selecionada} - Regime: {regime}")

            # Filtrar os dados com base no regime
            if regime == 'Caixa':
                coluna_regime = 'Pagamento'
            else:
                coluna_regime = 'Data'

            # Calcular totais com base no regime
            total_recebimentos = df_receivables_filtrado[df_receivables_filtrado[coluna_regime].between(start_date, end_date)]['Recebimentos'].sum()
            total_pagamentos = df_payables_filtrado[df_payables_filtrado[coluna_regime].between(start_date, end_date)]['Pagamentos'].sum()
            saldo_liquido = total_recebimentos - total_pagamentos

            # Exibir indicadores em colunas
            col1, col2, col3 = st.columns(3)
            col1.metric("Total de Recebimentos", f"R$ {total_recebimentos:,.2f}")
            col2.metric("Total de Pagamentos", f"R$ {total_pagamentos:,.2f}")
            col3.metric("Saldo Líquido", f"R$ {saldo_liquido:,.2f}")

            # Média Diária de Recebimentos e Pagamentos
            dias_periodo = (end_date - start_date).days + 1
            media_recebimentos = total_recebimentos / dias_periodo
            media_pagamentos = total_pagamentos / dias_periodo

            col4, col5 = st.columns(2)
            col4.metric("Média Diária de Recebimentos", f"R$ {media_recebimentos:,.2f}")
            col5.metric("Média Diária de Pagamentos", f"R$ {media_pagamentos:,.2f}")

            # Maior Recebimento e Pagamento
            if not df_receivables_filtrado.empty:
                df_receivables_filtrado_periodo = df_receivables_filtrado[df_receivables_filtrado[coluna_regime].between(start_date, end_date)]
                if not df_receivables_filtrado_periodo.empty:
                    maior_recebimento = df_receivables_filtrado_periodo.loc[df_receivables_filtrado_periodo['Recebimentos'].idxmax()]
                    st.write(f"**Maior Recebimento:** {maior_recebimento['Conta Analítica']} - R$ {maior_recebimento['Recebimentos']:,.2f}")
                else:
                    st.write("**Maior Recebimento:** Nenhum recebimento no período selecionado.")
            else:
                st.write("**Maior Recebimento:** Nenhum dado disponível para a unidade selecionada.")

            if not df_payables_filtrado.empty:
                df_payables_filtrado_periodo = df_payables_filtrado[df_payables_filtrado[coluna_regime].between(start_date, end_date)]
                if not df_payables_filtrado_periodo.empty:
                    maior_pagamento = df_payables_filtrado_periodo.loc[df_payables_filtrado_periodo['Pagamentos'].idxmax()]
                    st.write(f"**Maior Pagamento:** {maior_pagamento['Conta Analítica']} - R$ {maior_pagamento['Pagamentos']:,.2f}")
                else:
                    st.write("**Maior Pagamento:** Nenhum pagamento no período selecionado.")
            else:
                st.write("**Maior Pagamento:** Nenhum dado disponível para a unidade selecionada.")

            # Análise Temporal
            st.subheader(f"Análise Temporal - Unidade: {unidade_selecionada} - Regime: {regime}")
            fig_tendencia = px.line(cash_flow, x=cash_flow.index, y='Saldo Acumulado',
                                    title=f"Tendência do Saldo Acumulado - Unidade: {unidade_selecionada}",
                                    labels={"value": "Valor (R$)", "index": "Data"})
            st.plotly_chart(fig_tendencia)

            # Análise por Unidade (se "Todas as Unidades" for selecionada)
            if unidade_selecionada == "Todas as Unidades":
                st.subheader("Análise por Unidade")
    
                # Filtrar os dados com base no regime
                if regime == 'Caixa':
                    coluna_regime = 'Pagamento'
                else:
                    coluna_regime = 'Data'

                # Calcular recebimentos e pagamentos por unidade com base no regime
                recebimentos_por_unidade = df_receivables[df_receivables[coluna_regime].between(start_date, end_date)].groupby('Unidade')['Recebimentos'].sum().reset_index()
                pagamentos_por_unidade = df_payables[df_payables[coluna_regime].between(start_date, end_date)].groupby('Unidade')['Pagamentos'].sum().reset_index()

                # Renomear as colunas para evitar duplicação
                recebimentos_por_unidade = recebimentos_por_unidade.rename(columns={'Recebimentos': 'Recebimentos'})
                pagamentos_por_unidade = pagamentos_por_unidade.rename(columns={'Pagamentos': 'Pagamentos'})

                # Concatenar os DataFrames
                tabela_resumida_unidades = pd.merge(recebimentos_por_unidade, pagamentos_por_unidade, on='Unidade', how='outer')

                # Preencher valores NaN com 0 (caso haja unidades sem recebimentos ou pagamentos)
                tabela_resumida_unidades = tabela_resumida_unidades.fillna(0)

                # Gráfico de Barras - Recebimentos e Pagamentos por Unidade
                fig_unidades = px.bar(tabela_resumida_unidades, x='Unidade', y=['Recebimentos', 'Pagamentos'],
                              title="Recebimentos e Pagamentos por Unidade",
                              labels={"value": "Valor (R$)", "variable": "Tipo", "Unidade": "Unidade"})
                st.plotly_chart(fig_unidades)

            # Botão para gerar relatório em PDF
            if st.button("Gerar Relatório em PDF"):
                pdf_output = generate_pdf(cash_flow, recebimentos_por_conta, pagamentos_por_conta, unidade_selecionada, regime)
                st.markdown(create_download_link(pdf_output, "relatorio_financeiro.pdf", "pdf"), unsafe_allow_html=True)

            # Botão para gerar relatório em Excel
            if st.button("Gerar Relatório em Excel"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    cash_flow.to_excel(writer, sheet_name='Fluxo de Caixa')
                    recebimentos_por_conta.to_excel(writer, sheet_name='Recebimentos por Conta')
                    pagamentos_por_conta.to_excel(writer, sheet_name='Pagamentos por Conta')
                output.seek(0)
                st.markdown(create_download_link(output.getvalue(), "relatorio_financeiro.xlsx", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
