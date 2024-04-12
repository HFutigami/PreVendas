from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from time import sleep
import pandas as pd
import numpy as np
from datetime import datetime
import flet as ft
import threading
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import io
import pyarrow as pa

sharepoint_base_url = 'https://gertecsao.sharepoint.com/sites/PowerBi-Estoque/'
folder_in_sharepoint = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Planilhas/'
sharepoint_user = 'gertec.visualizador@gertec.com.br'
sharepoint_password = 'VY&ks28@AM2!hs1'

saldo_exp_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Bases%20de%20Dados/saldo_exp.parquet'

auth = AuthenticationContext(sharepoint_base_url)
auth.acquire_token_for_user(sharepoint_user, sharepoint_password)
ctx = ClientContext(sharepoint_base_url, auth)
web = ctx.web
ctx.execute_query()

def df_sharep(file_url, header=0, format='parquet'):
    """Gera um DataFrame a partir de um diretório do SharePoint."""
    file_response = File.open_binary(ctx, file_url)
    bytes_file_obj = io.BytesIO()
    bytes_file_obj.write(file_response.content)
    bytes_file_obj.seek(0)
    if format == 'parquet':
        return pd.read_parquet(bytes_file_obj)
    elif format == 'csv':
        return pd.read_csv(bytes_file_obj, header=header, sep=';')
    elif format == 'excel':
        return pd.read_excel(bytes_file_obj, header=header, dtype='str')

def login_intranet(navegador):
    """"""
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div[1]/div[1]/form/input[1]'))).send_keys("lab")
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div[1]/div[1]/form/input[2]'))).send_keys("gertec")
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div[1]/div[1]/form/input[3]'))).click()
    
def gerar_etiqueta(navegador, codigo, hw, op, ns, fab):
    """"""
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[1]/input'))).send_keys(codigo)
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[6]/input'))).send_keys(hw)
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[3]/input'))).send_keys(op)
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[5]/input'))).send_keys(ns)
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[4]/input'))).send_keys(1)
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[3]/img'))).click()
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[7]/select'))).click()
    if fab == "BA":
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[7]/select/option[1]'))).click()
    else:
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[7]/select/option[2]'))).click()
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[4]/form/p[8]/select'))).click()
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[4]/form/p[8]/select/option[2]'))).click()
    WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[4]/form/input[2]'))).click()
    
    navegador.execute_script('window.print();')
    
def abrir_saldo():
    try:
        saldo = df_sharep(saldo_exp_url)
        saldo['Val.NF Entr.'] = saldo['Val.NF Entr.'].astype(np.float64)
        saldo['Dt Entrada'] = pd.to_datetime(saldo['Dt Entrada'])
        saldo['Dt Recebimento'] = pd.to_datetime(saldo['Dt Recebimento'])
    except:
        saldo = abrir_saldo()
    return saldo


def main(page: ft.Page):
    page.title = "Etiquetas"
    page.window_width = 500
    page.window_height = 300
    page.window_resizable = False

    ### Configurações e login do ChromeDriver no Gerfloor
    def chrome_drive_login():
        servico = Service(ChromeDriverManager().install())
        chrome_options = Options()
        # chrome_options.add_argument("--headless=new")
        chrome_options.add_argument('--kiosk-printing')
        chrome_options.add_experimental_option("detach", True)
        navegador = webdriver.Chrome(service=servico, options=chrome_options)
        return navegador


    ### Telas de mensagens
    ns_pendente = ft.AlertDialog(title=ft.Text('Ainda existem seriais sendo processadas!', size=20), bgcolor='RED')


    ### Funções
    def fechar_chrome(e):
        """
        Caso o usuário tente fechar o aplicativo, bloqueia se ainda tiver caixa sendo processada
        ou fecha tudo, incluindo o ChromeDriver.
        """
        if e.data == 'close':
            if len(seriais_lista.controls) != 0:
                page.dialog = ns_pendente
                ns_pendente.open = True
                page.update()
            else:
                page.window_destroy()
                navegador.quit()
        

    def alterar_lista():
        while True:
            if len(seriais_lista.controls) != 0:
                navegador.get('http://172.17.1.23:8082/Intranet/index.mtw')
                login_intranet(navegador)
                navegador.get('http://172.17.1.23:8082/Intranet/relatorios/producao/etiquetas/mecanica/novoPadraoNS/embalagem50x30.jsp')
                dados = seriais_lista.controls[0].value
                dados = dados.split("-")
                if dados[0][1:2] == "M":
                    dados[4] = "AM"
                elif dados[0][1:2] == "I":
                    dados[4] = "BA"
                gerar_etiqueta(navegador, dados[0], dados[1], dados[2], dados[3], dados[4])
                seriais_lista.controls.remove(seriais_lista.controls[0])
                page.update()

    def atualizar_seriais_lista(e):
        if len(input_seriais.value) == 6:
            caixa = input_seriais.value
            input_seriais.value = ''
            input_seriais.focus()
            saldo = abrir_saldo()
            saldo = saldo[saldo['Caixa'] == caixa]['Nr Serie']
            saldo = list(saldo)
            for i in saldo:
                seriais_lista.controls.append(ft.Text(input_cod.value+"-"+input_HW.value+"-"+input_OP.value+"-"+i[:15]+"-"+input_fab.value))
            page.update()
        else:
            dados = input_seriais.value
            try:
                lista_seriais = dados.split(",")
                if lista_seriais[0][0:1] == "@":
                    for j in range(1, len(lista_seriais)):
                        seriais_lista.controls.append(ft.Text(input_cod.value+"-"+input_HW.value+"-"+input_OP.value+"-"+lista_seriais[j][:15]+"-"+input_fab.value))
                        page.update()                
                else:
                    for j in range(len(lista_seriais)):
                        seriais_lista.controls.append(ft.Text(input_cod.value+"-"+input_HW.value+"-"+input_OP.value+"-"+lista_seriais[j][:15]+"-"+input_fab.value))
                        page.update()
            except:
                input_seriais.value = ''
                input_seriais.focus()
                seriais_lista.controls.append(ft.Text(input_cod.value+"-"+input_HW.value+"-"+input_OP.value+"-"+i[:15]+"-"+input_fab.value))
                page.update()


    ### Navegador
    navegador = chrome_drive_login()

    ### Visual do Flet
    
    # Evita que o aplicativo seja fechado
    page.window_prevent_close = True
    page.on_window_event = fechar_chrome

    # Visuais
    input_seriais = ft.TextField(label='Seriais / Diretório',
                                 on_submit=atualizar_seriais_lista)
    input_cod = ft.TextField(label='Código do Produto')
    input_OP = ft.TextField(label='OP')
    input_HW = ft.TextField(label='HW')
    input_fab = ft.RadioGroup(content=ft.Column([
        ft.Radio(value="BA", label="Ilhéus"),
        ft.Radio(value="AM", label="Manaus")
    ]))
    seriais_lista = ft.ListView(expand=1, spacing=10, padding=20, auto_scroll=True)
        
    page.add(input_seriais, input_cod, input_HW, input_OP, input_fab, seriais_lista)
    
    # Threading para processar as caixas em segundo plano
    thread = threading.Thread(target=alterar_lista)
    thread.daemon = True
    input_seriais.on_scroll = thread.start()

ft.app(target=main)

