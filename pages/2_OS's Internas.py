import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import io
import plotly.express as px


"""Exigência de credênciais."""
if 'senha' not in st.session_state:
    with st.form('login', clear_on_submit=True):
        senha = st.text_input(label='Senha', type='password')
        submitted = st.form_submit_button("Login")
        if submitted:
            if senha == st.secrets.visual_pass or senha == st.secrets.editor_pass:
                st.session_state['senha'] = senha
                st.rerun()
            else:
                st.warning('Senha inválida!', icon="⚠️")

elif st.session_state['senha'] == st.secrets.visual_pass or st.session_state['senha'] == st.secrets.editor_pass:

    """Links necessários para o programa."""
    base_sharepoint_url = 'https://gertecsao.sharepoint.com/sites/Expedio/' #link
    prevendas_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/' #link

    rbt_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/RBT.xlsx' #xlsx
    lpn_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/Utility/LPN.parquet' #parquet
    osi_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/OSI.parquet' #parquet
    p911_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/Utility/P911.parquet' #parquet
    hist_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/HIST.parquet' #parquet


    """Autenticação no Sharepoint"""
    auth = AuthenticationContext(base_sharepoint_url)
    auth.acquire_token_for_user(st.secrets.sharepoint_credenciais.user,
                                st.secrets.sharepoint_credenciais.password)
    ctx = ClientContext(prevendas_sharepoint_url, auth)
    web = ctx.web
    ctx.execute_query()


    def df_sharep(file_url, tipo='parquet', sheet=''):
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
                return pd.read_excel(bytes_file_obj, sheet, dtype='str')
            else:
                return pd.read_excel(bytes_file_obj, dtype='str')
            
    
    def importaçao_os(seriais):
        """Gera a importação de uma OS através de informações do HIST."""


    def cadastrar_os(seriais, numos):
        """Cadastra uma OS no HIST a partir de um DataFrame informado."""
