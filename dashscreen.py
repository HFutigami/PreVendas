import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import io
from time import sleep
from datetime import datetime, timedelta

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
else:
    ''
  # nf = 109187
  # user = '41695910893'
  # senha = '123Mudar'  

  # st.write(datetime.now())
  # servico = Service(ChromeDriverManager().install())
  # st.write(datetime.now())
  # chrome_options = Options()
  # # chrome_options.add_argument("--headless=new")
  # navegador = webdriver.Chrome(service=servico, options=chrome_options)

  # navegador.get('https://gertec.mastertax.app/login')
  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/nz-layout/app-login/div/div/div[2]/form/nz-form-item[1]/nz-form-control/div/div/nz-input-group/input'))).send_keys('41695910893')
  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/nz-layout/app-login/div/div/div[2]/form/nz-form-item[2]/nz-form-control/div/div/nz-input-group/input'))).send_keys('123Mudar')
  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/nz-layout/app-login/div/div/div[2]/form/nz-form-item[2]/nz-form-control/div/div/nz-input-group/input'))).send_keys(Keys.RETURN)

  # sleep(3)

  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/nz-layout/app-header/div/app-megamenu/div'))).click()
  
  # sleep(2)

  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[2]/div/nz-modal-container/div/div/div[2]/div/div[1]/div[1]/ul/li/a'))).click()

  # WebDriverWait(navegador, 200).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div[3]/table/tbody/tr[5]')))

  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form[1]/div/div/div[2]/div[2]/input'))).send_keys(nf)
  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form[1]/div/div/div[2]/div[2]/input'))).send_keys(Keys.RETURN)

  # WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/form[2]/div/div/div/div[2]/div/table/tbody/tr[1]/td[10]/a[1]'))).click()
