import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import io
import plotly.express as px
import os

st.set_page_config(page_title='Dash', page_icon='https://i.imgur.com/mOEfCM8.png', layout='wide')

st.image('https://seeklogo.com/images/G/gertec-logo-D1C911377C-seeklogo.com.png?v=637843433630000000', width=200)
st.header('', divider='gray')

if 'lpn' not in st.session_state:
    lpn_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/LPN.xlsx' #xlsx
    base_sharepoint_url = 'https://gertecsao.sharepoint.com/sites/Expedio/' #link
    prevendas_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/' #link

    #Autenticação no Sharepoint
    auth = AuthenticationContext(base_sharepoint_url)
    auth.acquire_token_for_user(st.secrets.sharepoint_credenciais.user,
                                st.secrets.sharepoint_credenciais.password)
    ctx = ClientContext(prevendas_sharepoint_url, auth)
    web = ctx.web
    ctx.execute_query()

    def df_sharep(file_url, tipo='parquet', sheet='', header=0):
        """Gera um DataFrame através de um arquivo do Sharepoint."""
        file_response = File.open_binary(ctx, file_url)
        bytes_file_obj = io.BytesIO()
        bytes_file_obj.write(file_response.content)
        bytes_file_obj.seek(0)
        if tipo == 'parquet':
            return pd.read_parquet(bytes_file_obj)
        elif tipo == 'htm':
            return pd.read_csv(bytes_file_obj)
        elif tipo == 'excel':
            if sheet != '':
                return pd.read_excel(bytes_file_obj, sheet, dtype='str', header=header)
            else:
                return pd.read_excel(bytes_file_obj, dtype='str', header=header)
            
    st.session_state['lpn'] = df_sharep(lpn_sharepoint_url, 'excel')

#Exigência de credênciais.
if 'senha' not in st.session_state:
    st.switch_page('dashscreen.py')

elif st.session_state['senha'] == st.secrets.visual_pass or st.session_state['senha'] == st.secrets.editor_pass:

    #Links necessários para o programa.
    base_sharepoint_url = 'https://gertecsao.sharepoint.com/sites/Expedio/' #link
    prevendas_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/' #link

    sd1_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/SD1.xlsx' #xlsx
    lpn_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/LPN.xlsx' #parquet
    hist_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/historico.parquet' #parquet


    #Autenticação no Sharepoint
    auth = AuthenticationContext(base_sharepoint_url)
    auth.acquire_token_for_user(st.secrets.sharepoint_credenciais.user,
                                st.secrets.sharepoint_credenciais.password)
    ctx = ClientContext(prevendas_sharepoint_url, auth)
    web = ctx.web
    ctx.execute_query()


    def df_sharep(file_url, tipo='parquet', sheet='', header=0):
        """Gera um DataFrame através de um arquivo do Sharepoint."""
        file_response = File.open_binary(ctx, file_url)
        bytes_file_obj = io.BytesIO()
        bytes_file_obj.write(file_response.content)
        bytes_file_obj.seek(0)
        if tipo == 'parquet':
            return pd.read_parquet(bytes_file_obj)
        elif tipo == 'htm':
            return pd.read_csv(bytes_file_obj)
        elif tipo == 'excel':
            if sheet != '':
                return pd.read_excel(bytes_file_obj, sheet, dtype='str', header=header)
            else:
                return pd.read_excel(bytes_file_obj, dtype='str', header=header)

        
    def imprimir_danfe(nf, data_emissão):
        """Abre um ChromeDriver e busca a Nota Fiscal através do MasterTax."""


    def extrair_seriais(nr_caixa):
        """Abre um ChromeDriver e busca no Gerfloor os seriais inseridos na caixa."""


    def cadastrar_seriais(seriais):
        """Cadastra os seriais no HIST com base em dataframe informado."""


    def imprimir_endereço(pn, vaga, armazem, caixa, qtd, obs):
        """Gera uma identificação de endereço."""


    def create_df_entradas():
        sd1 = df_sharep(sd1_sharepoint_url, 'excel', header=2)
        lpn = df_sharep(lpn_sharepoint_url, 'excel')
        sd1.rename(columns={
            'Produto     ':'PN',
            'Filial      ':'FILIAL',
            'Quantidade  ':'QUANTIDADE',
            'Tipo Entrada':'TES',
            'Cod. Fiscal ':'CFOP',
            'Documento   ':'NF ENTRADA',
            'DT Emissao  ':'DT EMISSAO',
            'DT Digitacao':'DT DIGITACAO', 
            'Serie       ':'SERIE',
            'Armazem     ':'ARMAZEM',
            'Docto. Orig.':'NF ORIGINAL',
            'Serie Orig. ':'SERIE SAIDA'
        }, inplace=True)
        sd1 = sd1.join(lpn.set_index('PN'), on='PN', how='inner')
        sd1 = sd1[['FILIAL', 'PN', 'QUANTIDADE', 'TES', 'CFOP', 'NF ENTRADA', 'SERIE', 'DT EMISSAO', 'DT DIGITACAO',
                'ARMAZEM', 'NF ORIGINAL', 'SERIE SAIDA']]
        sd1['NF ENTRADA'] = sd1['NF ENTRADA'] + '/' + sd1['SERIE']
        sd1['NF ORIGINAL'] = sd1['NF ORIGINAL'] + '/' + sd1['SERIE SAIDA']
        sd1.drop(columns=['SERIE', 'SERIE SAIDA'], inplace=True)

        sd1.reset_index(drop=True, inplace=True)
        return sd1


    def create_df_pendencias():
        hist = df_sharep(hist_sharepoint_url)
        lpn = st.session_state['lpn']
        hist = hist.join(lpn.set_index('PN'), on='PN', how='left')
        hist = hist[hist['DESC FLUXO'].isin(['AG. ENTRADA', 'AG. FÍSICO'])]
        return hist


    def create_df_ag_entrada():
        hist = df_sharep(hist_sharepoint_url)
        lpn = st.session_state['lpn']
        hist = hist.join(lpn.set_index('PN'), on='PN', how='left')
        hist = hist[hist['DESC FLUXO'] == 'AG. ENTRADA']
        hist = hist.groupby(['FILIAL', 'ARMAZEM', 'ENDEREÇO', 'CAIXA', 'PN', 'DESCRICAO'])[['NS']].count().reset_index()
        hist.rename(columns={'NS':'QUANTIDADE'}, inplace=True)
        return hist


    def create_fig_hist_ag_entrada():
        df = st.session_state['ag_entrada_detalhado']
        df['DATA'] = df['DATA RECEBIMENTO'].str[:7].str.replace("-","/")
        df = df.groupby(['DATA'])[['NS']].count().reset_index()

        fig = px.bar(
            df,
            x='DATA',
            y='NS',
            text='NS'
        )

        fig.update_traces(textposition='outside',
                          orientation='v',
                          textfont_size=14)
        
        fig.update_layout(yaxis_title=None,
                          xaxis_title=None,
                          yaxis_visible=False)

        return fig


    if 'historico_entradas' not in st.session_state:
        st.session_state['historico_entradas'] = create_df_entradas()

    tabs_ae, tabs_af = st.tabs(['Aguardando Saldo', 'Aguardando Físico'])

    if 'df_pendencias_entradas' not in st.session_state:
        st.session_state['df_pendencias_entradas'] = create_df_pendencias()


    tabs_ae.title('Aguardando Entrada')

    if 'df_ag_entrada' not in st.session_state:
        st.session_state['df_ag_entrada'] = create_df_ag_entrada()

    r1_col1, r1_col2 = tabs_ae.columns(2)
    r1_col1.write('Resumo de equipamentos agurdando entrada.')
    df_ag_entrada = r1_col1.dataframe(st.session_state['df_ag_entrada'][['FILIAL',
                                                                 'ARMAZEM',
                                                                 'ENDEREÇO',
                                                                 'CAIXA',
                                                                 'PN',
                                                                 'DESCRICAO',
                                                                 'QUANTIDADE']], hide_index=True, on_select='rerun',
                                                                 use_container_width=True)

    ag_entrada_detalhado = st.session_state['df_pendencias_entradas']
    ag_entrada_simples = st.session_state['df_ag_entrada']
    ag_entrada_detalhado['CONCATENADO'] = ag_entrada_detalhado['FILIAL'] + ag_entrada_detalhado['PN'] + ag_entrada_detalhado['ENDEREÇO'] + ag_entrada_detalhado['CAIXA'] 
    ag_entrada_simples['CONCATENADO'] = ag_entrada_simples['FILIAL'] + ag_entrada_simples['PN'] + ag_entrada_simples['ENDEREÇO'] + ag_entrada_simples['CAIXA']
    if df_ag_entrada.selection.rows:
        filtro = list(ag_entrada_simples.iloc[df_ag_entrada.selection.rows]['CONCATENADO'])
        ag_entrada_detalhado = ag_entrada_detalhado[ag_entrada_detalhado['CONCATENADO'].isin(filtro)]
        st.session_state['ag_entrada_detalhado'] = ag_entrada_detalhado

    r2_col1, r2_col2 = tabs_ae.columns(2)
    if 'ag_entrada_detalhado' in st.session_state and df_ag_entrada.selection.rows:
        r1_col2.write('Distribuição de equipamentos recebido por mês.')
        r1_col2.plotly_chart(create_fig_hist_ag_entrada())

        r2_col1.write('Lista detalhada de equipamentos recebidos.')
        df_ag_entrada_detalhado = r2_col1.dataframe(st.session_state['ag_entrada_detalhado'][['NS',
                                                          'PN',
                                                          'CAIXA',
                                                          'ENDEREÇO',
                                                          'FILIAL',
                                                          'NF ENTRADA',
                                                          'DATA RECEBIMENTO',
                                                          'OBS']],
                        hide_index=True,
                        use_container_width=True,
                        on_select='rerun',
                        column_config={'DATA RECEBIMENTO': st.column_config.DateColumn('DT RECEBIMENTO', format="DD/MM/YYYY")})
        
    hist_entradas_ag_entrada = st.session_state['historico_entradas']
    
    if 'df_ag_entrada_detalhado' in locals():
        if df_ag_entrada_detalhado.selection.rows:
            filtro = list(ag_entrada_detalhado.iloc[df_ag_entrada_detalhado.selection.rows]['PN'])
            hist_entradas_ag_entrada = hist_entradas_ag_entrada[hist_entradas_ag_entrada['PN'].isin(filtro)]
            hist_entradas_ag_entrada['DT EMISSAO'] = pd.to_datetime(hist_entradas_ag_entrada['DT EMISSAO'])
            hist_entradas_ag_entrada['DT DIGITACAO'] = pd.to_datetime(hist_entradas_ag_entrada['DT DIGITACAO'])
            st.session_state['hist_entradas_ag_entrada'] = hist_entradas_ag_entrada
        
        if 'hist_entradas_ag_entrada' in st.session_state and df_ag_entrada_detalhado.selection.rows:
            ag_entrada_detalhado['CONCATENADO2'] = ag_entrada_detalhado['FILIAL'] + ag_entrada_detalhado['PN'] + ag_entrada_detalhado['NF ENTRADA']
            ag_entrada_detalhado['CONCATENADO3'] = ag_entrada_detalhado['PN'] + ag_entrada_detalhado['NF ENTRADA']

            def color_coding(row):
                if row['FILIAL'] + row['PN'] + row['NF ENTRADA'] in ag_entrada_detalhado.iloc[df_ag_entrada_detalhado.selection.rows]['CONCATENADO2']:
                    return ['background-color:green'] * len(row)
                elif row['PN'] + row['NF ENTRADA'] in ag_entrada_detalhado.iloc[df_ag_entrada_detalhado.selection.rows]['CONCATENADO3']:
                    return ['background-color:yellow'] * len(row)
                else:
                    return [''] * len(row)

            r2_col2.write('Histórico de entrada de saldo.')
            df_entradas_ag_entrada = r2_col2.dataframe(hist_entradas_ag_entrada[['FILIAL',
                    'PN',
                    'QUANTIDADE',
                    'TES',
                    'CFOP',
                    'NF ENTRADA',
                    'DT EMISSAO',
                    'DT DIGITACAO',
                    'ARMAZEM',
                    'NF ORIGINAL']].sort_values(['DT DIGITACAO'], ascending=False).style.format({'DT EMISSAO':'{:%d/%m/%Y}',
                                                                                                 'DT DIGITACAO':'{:%d/%m/%Y}'}).apply(color_coding, axis=1),
                hide_index=True, use_container_width=True,
                column_config={'DT EMISSAO': st.column_config.DateColumn('DT EMISSÃO', format='DD/MM/YYYY'),
                               'DT DIGITACAO': st.column_config.DateColumn('DT DIGITAÇÃO', format='DD/MM/YYYY')})


