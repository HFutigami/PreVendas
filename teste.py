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
import pathlib
import pyarrow as pa

sharepoint_base_url = 'https://gertecsao.sharepoint.com/sites/PowerBi-Estoque/'
folder_in_sharepoint = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Planilhas/'
sharepoint_user = 'gertec.visualizador@gertec.com.br'
sharepoint_password = 'VY&ks28@AM2!hs1'

protheus_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Bases%20de%20Dados/protheus.parquet'
rebatismo_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Bases%20de%20Dados/rebatismo.parquet'
saldo_exp_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Bases%20de%20Dados/saldo_exp.parquet'

base_dados_sharep_url = '/sites/PowerBi-Estoque/Documentos%20Compartilhados/General/Expedi%C3%A7%C3%A3o/Bases%20de%20Dados/'

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


def login_gerfloor(navegador):
    """
    Faz todo o processo de login no site do gerfloor.
    """
    try:
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-select-value-1"]/span'))).click()
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-option-4"]'))).click()
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-login/div/mat-card-content/form/mat-form-field[2]/div[1]/div/div[3]/input'))).send_keys("Jfutigami")
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-login/div/mat-card-content/form/mat-form-field[3]/div[1]/div/div[3]/input'))).send_keys("Gertec#77685454")
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-login/div/mat-card-actions/button'))).click()
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/gf-header/mat-toolbar/button'))).click()  
    except:
        sleep(1)
        login_gerfloor(navegador)


def mov_embalagem_estoque(navegador, caixa):
    """
    Movimenta a caixa do laboratório para o estoque.
    """
    try:
        navegador.get("https://psg.gertec.com.br/box-moving")
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-box-moving/mat-card/mat-card-content/form/mat-card/div/span[1]/mat-form-field/div[1]'))).click()
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div/mat-option[2]/span'))).click()
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-box-moving/mat-card/mat-card-content/form/mat-card/div/span[2]/mat-form-field/div[1]/div/div[2]/input'))).send_keys(caixa)
        WebDriverWait(navegador, 2).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-box-moving/mat-card/mat-card-content/form/mat-card/div/span[2]/mat-form-field/div[1]/div/div[2]/input'))).send_keys(Keys.RETURN)

        result = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/mat-snack-bar-container/div/div/div/div/gf-mat-snack-bar/div[2]'))).text
        if result == 'Falha ao salvar o registro --> Packing não encontrado':
            return False
    except:
        if 'login' in navegador.current_url:
            login_gerfloor(navegador)
        mov_embalagem_estoque(navegador, caixa)
    return True

def mov_estoque_expedicao(navegador, caixa):
    """
    Faz a movimentação do estoque para a expedição.
    """
    try:
        navegador.get("https://psg.gertec.com.br/box-moving")
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-box-moving/mat-card/mat-card-content/form/mat-card/div/span[1]/mat-form-field/div[1]'))).click()
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div/mat-option[4]/span'))).click()
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-box-moving/mat-card/mat-card-content/form/mat-card/div/span[2]/mat-form-field/div[1]/div/div[2]/input'))).send_keys(caixa)
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-box-moving/mat-card/mat-card-content/form/mat-card/div/span[2]/mat-form-field/div[1]/div/div[2]/input'))).send_keys(Keys.RETURN)
    
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/mat-snack-bar-container/div/div/div/div/gf-mat-snack-bar/div[2]'))).text
    except:
        if 'login' in navegador.current_url:
            login_gerfloor(navegador)
        mov_embalagem_estoque(navegador, caixa)


def coleta_infos_seriais(navegador, caixa):
    """
    Busca no gerfloor os seriais e o status dos equipamentos dentro da caixa
    e retorna uma list com todos os seriais e uma string com o status.
    """
    try:
        navegador.get("https://psg.gertec.com.br/packing/remove")
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-packing-remove/mat-card/mat-card/mat-card-content/form/mat-form-field/div[1]/div/div[2]/input'))).send_keys(caixa)
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-packing-remove/mat-card/mat-card/mat-card-content/form/mat-form-field/div[1]/div/div[2]/input'))).send_keys(Keys.RETURN)
        
        sleep(1)

        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-packing-remove/mat-card/mat-card/mat-paginator/div/div/div[1]/mat-form-field/div[1]/div/div[2]/mat-select'))).click()
        WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[2]/div/div/mat-option[3]'))).click()

        status_equip = WebDriverWait(navegador, 20).until(EC.visibility_of_element_located((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-packing-remove/mat-card/div[1]/span[2]'))).text
        cliente = WebDriverWait(navegador, 20).until(EC.visibility_of_element_located((By.XPATH, '/html/body/app-root/mat-sidenav-container/mat-sidenav-content/gf-packing-remove/mat-card/div[1]/span[1]'))).text
        chtml = navegador.page_source

        seriais_gerfloor = pd.read_html(chtml)
        seriais_gerfloor[0].iloc[:,3] = seriais_gerfloor[0].iloc[:,3].astype('str')
        
        lista_seriais = list(seriais_gerfloor[0].iloc[:,3])

        status_equip = status_equip.split(':')[1].strip()
        cliente = cliente.split(':')[1].strip()
    except:
        if 'login' in navegador.current_url:
            login_gerfloor(navegador)
        lista_seriais, status_equip, cliente = coleta_infos_seriais(navegador, caixa)

    return lista_seriais, status_equip, cliente




def abrir_saldo():
    try:
        saldo = df_sharep(saldo_exp_url)
        saldo['Val.NF Entr.'] = saldo['Val.NF Entr.'].astype(np.float64)
        saldo['Dt Entrada'] = pd.to_datetime(saldo['Dt Entrada'])
        saldo['Dt Recebimento'] = pd.to_datetime(saldo['Dt Recebimento'])
    except:
        saldo = abrir_saldo()
    return saldo


def atualizar_saldo(seriais, status, cliente, caixa):
    """
    
    """
    # Coletando informações apenas dos seriais bipados
    infos = base[base['Nr Serie'].isin(seriais)]
    infos.insert(9, 'Status', status)
    infos.insert(9, 'Caixa', caixa)
    infos.insert(9, 'Dt Recebimento', datetime.now())

    if len(infos['Nr Serie']) != len(seriais):
        for i in range(len(seriais)):
            if seriais[i] not in list(infos['Nr Serie']): seriais[i] = str(seriais[i]).zfill(16)
        infos2 = base[base['Nr Serie'].isin(seriais)]
        infos2 = pd.concat([infos, infos2])
        if len(infos2['Nr Serie']) != len(seriais):
            seriais_sem_dados = [[x] for x in seriais if x not in list(infos2['Nr Serie'])]
            df_sem_dados = pd.DataFrame(columns=['Nr Serie'], data=seriais_sem_dados)
            df_sem_dados['Status'] = status
            df_sem_dados['Client Final'] = cliente
            df_sem_dados['Caixa'] = caixa
            df_sem_dados['Dt Recebimento'] = datetime.now()
            
            infos = pd.concat([infos2, df_sem_dados], ignore_index=True)

            # Atualizando a base de Saldo
            saldo = abrir_saldo()
            saldo = pd.concat([saldo, infos], ignore_index=True)
            
            parquet_buffer = io.BytesIO()
            tabela = pa.Table.from_pandas(saldo)
            pa.parquet.write_table(tabela, parquet_buffer)

            # Salvando a base de Saldo
            File.save_binary(ctx, base_dados_sharep_url + 'saldo_exp.parquet', parquet_buffer.getvalue())
            return False
        else:
            infos2[infos2['Status'].isna()][['Status']] = status
            infos2[infos2['Caixa'].isna()][['Caixa']] = caixa
            infos2[infos2['Dt Recebimento'].isna()][['Dt Recebimento']] = datetime.now()
            infos = infos2
    
    # Atualizando a base de Saldo
    saldo = abrir_saldo()
    saldo = pd.concat([saldo, infos], ignore_index=True)
    
    parquet_buffer = io.BytesIO()
    tabela = pa.Table.from_pandas(saldo)
    pa.parquet.write_table(tabela, parquet_buffer)

    # Salvando a base de Saldo
    File.save_binary(ctx, base_dados_sharep_url + 'saldo_exp.parquet', parquet_buffer.getvalue())

    return True


### DataFrames

base = df_sharep(protheus_url)
saldo = abrir_saldo()


### Aplicativo Flet

def main(page: ft.Page):
    page.title = "Movimentações - Pré-Expedição"
    page.window_width = 500
    page.window_height = 300
    page.window_resizable = False
    lista_caixas = []

    ### Configurações e login do ChromeDriver no Gerfloor
    def chrome_drive_login():
        servico = Service(ChromeDriverManager().install())
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_experimental_option("detach", True)
        navegador = webdriver.Chrome(service=servico, options=chrome_options)
        return navegador


    ### Telas de mensagens
    def caixa_error(caixa):
        return ft.AlertDialog(title=ft.Text(f'Caixa {caixa} inválida!', size=20), bgcolor='RED')


    ### Funções
    def fechar_chrome(e):
        """
        Caso o usuário tente fechar o aplicativo, bloqueia se ainda tiver caixa sendo processada
        ou fecha tudo, incluindo o ChromeDriver.
        """
        if e.data == 'close':
            if len(caixa_list.controls) != 0:
                lista_caixas.clear()
                for i in range(len(caixa_list.controls)):
                    lista_caixas.append([str(caixa_list.controls[i]).split("'")[-2]])
                caixas_pendentes = pd.DataFrame(columns=['Caixa'], data=lista_caixas)
                caixas_pendentes.to_csv(str(pathlib.Path().resolve()) + f'\CaixasPendentes.csv', index=False)
            else:
                caixas_pendentes = pd.DataFrame(columns=['Caixa'])
                caixas_pendentes.to_csv(str(pathlib.Path().resolve()) + f'\CaixasPendentes.csv', index=False)
                page.window_destroy()
        

    def alterar_lista():
        """
        Processa cada caixa bipada uma a uma.
        """
        while len(lista_caixas) == 0:
            if len(caixa_list.controls) != 0:
                try:
                    navegador.get('https://psg.gertec.com.br')
                except:
                    navegador = chrome_drive_login()
                    navegador.get('https://psg.gertec.com.br')
                sleep(1)
                if 'login' in navegador.current_url:
                    login_gerfloor(navegador)
                else:
                    saldo = abrir_saldo()
                    caixa_info = caixa_list.controls[0].value
                    if len(saldo[saldo['Caixa'] == caixa_info]) == 0:
                        if mov_embalagem_estoque(navegador, caixa_info):
                            seriais_info, status_info, cliente_info = coleta_infos_seriais(navegador, caixa_info)
                            mov_estoque_expedicao(navegador, caixa_info)
                            att = atualizar_saldo(seriais_info, status_info, cliente_info, caixa_info)
                            if att is False:
                                resultado_list.controls.insert(0, ft.Text(str(caixa_list.controls[0]).split("'")[-2] + ' - SEM DADOS NA BASE', bgcolor=ft.colors.ORANGE_100))
                            else: resultado_list.controls.insert(0, ft.Text(str(caixa_list.controls[0]).split("'")[-2] + ' - MOVIMENTADA COM SUCESSO', color=ft.colors.WHITE, bgcolor=ft.colors.GREEN_500))
                        else: resultado_list.controls.insert(0, ft.Text(str(caixa_list.controls[0]).split("'")[-2] + ' - PACKING NÃO ENCONTRADO', color=ft.colors.WHITE, bgcolor=ft.colors.RED_500))
                    else: resultado_list.controls.insert(0, ft.Text(str(caixa_list.controls[0]).split("'")[-2] + ' - JÁ POSSUI REGISTRO', bgcolor=ft.colors.LIGHT_GREEN_100))
                    caixa_list.controls.remove(caixa_list.controls[0])
                    page.update()


    def caixa(e):
        """
        Adiciona a caixa bipada a uma list para que posssamos trata-las uma a uma.
        """
        if len(nr_caixa.value) != 6:
            cx_error = caixa_error(nr_caixa.value)
            page.dialog = cx_error
            cx_error.open = True
            nr_caixa.value = ''
            nr_caixa.focus()
            page.update()
        else:
            i = nr_caixa.value
            nr_caixa.value = ''
            nr_caixa.focus()
            caixa_list.controls.append(ft.Text(i))
            page.update()

    ### Visual do Flet
    
    # Evita que o aplicativo seja fechado
    page.window_prevent_close = True
    page.on_window_event = fechar_chrome

    # Visuais
    nr_caixa = ft.TextField(on_submit=caixa)
    resultado_list = ft.ListView(expand=1, spacing=10, padding=20)
    caixa_list = ft.ListView(expand=1, spacing=10, padding=20, auto_scroll=True)
    col1 = ft.Column(col=6, controls=[caixa_list])
    col2 = ft.Column(col=6, controls=[resultado_list])
    lin = ft.ResponsiveRow([col1, col2])
    page.add(nr_caixa, lin)

    cxs_pend = pd.read_csv(str(pathlib.Path().resolve()) + f'\CaixasPendentes.csv')
    for i in cxs_pend['Caixa']:
        caixa_list.controls.append(ft.Text(str(i)))
        page.update()

    # Threading para processar as caixas em segundo plano
    thread = threading.Thread(target=alterar_lista)
    thread.daemon = True
    caixa_list.on_scroll = thread.start()

ft.app(target=main)
