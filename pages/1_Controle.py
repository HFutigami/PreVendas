import streamlit as st
import pandas as pd
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from datetime import datetime
import io
import plotly.express as px
import webbrowser
import os
import barcode
from barcode.writer import ImageWriter

st.set_page_config(page_title='Dash', page_icon='https://i.imgur.com/mOEfCM8.png', layout='wide')

st.image('https://seeklogo.com/images/G/gertec-logo-D1C911377C-seeklogo.com.png?v=637843433630000000', width=200)
st.header('', divider='gray')

#Exigência de credênciais.
if 'senha' not in st.session_state:
    st.switch_page('dashscreen.py')

elif st.session_state['senha'] == st.secrets.visual_pass or st.session_state['senha'] == st.secrets.editor_pass:

    #Links necessários para o programa.
    base_sharepoint_url = 'https://gertecsao.sharepoint.com/sites/Expedio/' #link
    prevendas_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/' #link

    spe_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/SPE.xlsx' #xlsx
    lpn_sharepoint_url = '/sites/Expedio/Documentos%20Compartilhados/General/Bases/LPN.xlsx' #xlsx
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


    def imprimir_serial_lista(rows):
        """Gera uma etiqueta 50x30 para o terminal ou para a caixa."""
        df = st.session_state['saldo_estoque_detalhado'][['NS',
                                                          'PN',
                                                          'DESCRICAO']].reset_index()

        html_content = '''<style>
            div, img {
                font-family: system-ui;
                font-weight: bold;
                font-size: 10px;
            }
            .container {
                display: grid;
                grid-template-columns: 40% 60%;
            }
            </style>
            
            <div class="container">'''

        for i in rows:
            pn = df.iloc[i]['PN']
            ns = df.iloc[i]['NS']
            desc = df.iloc[i]['DESCRICAO']

            code = barcode.get_barcode('code128', f'{ns}', writer=ImageWriter())

            options = {
                'module_width': 0.5,
                'module_height': 7,
                'font_size': 1,
                'text_distance': -2,
                'quiet_zone': 5,
                'color': 'black',
                'background': 'white'
            }
            code.save(f'etiquetas/{ns}', options=options)

            html_content += '''
            <div>'''+ pn + '''</div>
            <div>'''+ desc + '''</div>
            <div> NS: ''' + ns + '''</div>
            <img src=''' + f'''"etiquetas/{ns}.png"''' + '''style="width: 100%"/>
            <div>ㅤ</div>
            <div>ㅤ</div>
            '''

        html_content += """</div>"""
        # Salve o conteúdo em um arquivo temporário
        temp_file = 'etiqueta.html'
        with open(temp_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # Abra o arquivo HTML no navegador padrão
        file_path = 'file://' + os.path.abspath(temp_file)
        webbrowser.open(file_path)


    def imprimir_serial_unitario(rows):
        """Gera uma etiqueta 50x30 para o terminal ou para a caixa."""
        df = st.session_state['saldo_estoque_detalhado'][['NS',
                                                          'PN',
                                                          'DESCRICAO']].reset_index()

        html_content = ''
        data = str(datetime.now().day).zfill(2) + '/' + str(datetime.now().month).zfill(2) + '/' + str(datetime.now().year)

        for i in rows:
            pn = df.iloc[i]['PN']
            ns = df.iloc[i]['NS']
            desc = df.iloc[i]['DESCRICAO']

            if pn[1:2] == 'M':
                cnpj = '03.654.119/0003-38'
            elif pn[1:2] == 'I':
                cnpj = '03.654.119/0001-76'
            else:
                cnpj = '03.654.119/0003-38'

            pn = str(pn).replace('M', '0').replace('I', '0')

            code = barcode.get_barcode('code128', f'{pn}', writer=ImageWriter())

            options = {
                'module_width': 0.5,
                'module_height': 5,
                'font_size': 1,
                'text_distance': -2,
                'quiet_zone': 5,
                'color': 'black',
                'background': 'white'
            }
            code.save(f'etiquetas/{pn}', options=options)

            code = barcode.get_barcode('code128', f'{ns}', writer=ImageWriter())

            options = {
                'module_width': 0.5,
                'module_height': 7,
                'font_size': 1,
                'text_distance': -2,
                'quiet_zone': 5,
                'color': 'black',
                'background': 'white'
            }
            code.save(f'etiquetas/{ns}', options=options)

            html_content += '''<style>
            span, div, img {
                font-family: system-ui;
                font-weight: bold;
            }
            img {
                width: 100%;
            }
            body {
                height: 100%;
                width: 100%;
            }
            div {
                text-align: center;
                font-size: 15px;
            }
            </style>
            <body>
            <span style="font-size: 20px;">'''+ desc + '''</span><br>
            <span style="font-size: 20px;">''' + '''Data: ''' + f'''{data}</span>
            <div>Cód. Gertec: ''' + pn + '''</div>
            <img src=''' + f'''"etiquetas/{pn}.png"''' + '''style="width: 100%"/>
            <div style="font-size: 20px;">NS: ''' + ns + '''</div>
            <img src=''' + f'''"etiquetas/{ns}.png"''' + '''style="width: 100%"/>
            <div>Indústria Brasileira<br>
            Fabricado por Gertec Brasil LTDA<br>
            CNPJ:''' + cnpj + '''<br>
            www.gertec.com.br</div>
            </body>'''

        # Salve o conteúdo em um arquivo temporário
        temp_file = 'etiqueta.html'
        with open(temp_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # Abra o arquivo HTML no navegador padrão
        file_path = 'file://' + os.path.abspath(temp_file)
        webbrowser.open(file_path)


    def imprimir_endereço(rows):
        """Gera uma identificação de endereço."""
        # Defina o conteúdo do documento HTML
        df = st.session_state['df_saldo'].groupby(['FILIAL',
                                                    'ARMAZEM',
                                                    'ENDEREÇO',
                                                    'CAIXA',
                                                    'PN',
                                                    'DESCRICAO'])['NS'].count().reset_index()

        html_content = ''
        for i in rows:
            pn = df.iloc[i]['PN']
            vaga = df.iloc[i]['ENDEREÇO']
            armazem = df.iloc[i]['ARMAZEM']
            caixa = df.iloc[i]['CAIXA']
            qtd = df.iloc[i]['NS']
            descriçao = df.iloc[i]['DESCRICAO']

            code = barcode.get_barcode('code128', f'{pn}', writer=ImageWriter())

            options = {
                'module_width': 0.5,
                'module_height': 10,
                'font_size': 8,
                'text_distance': 4,
                'quiet_zone': 5,
                'color': 'black',
                'background': 'white'
            }
            code.save(f'etiquetas\{pn}_barcode', options=options)

            html_content += '''
            <html>
            <style>
            table, th, td {
                border:1px solid black;
                width:100%;
                border-collapse: collapse;
                font-weight: bold;
                font-family: system-ui;
                font-size: 22px;
            }
            th {
                width:14.33%;
            }
            tr {
                height:35px;
            }
            </style>
            <body>

            <table>
            <tr>
                <th colspan="5" rowspan="2" style="font-size:40px">'''+str(pn)+'''</th>
                <th style="background-color:rgb(0, 0, 0); color:rgb(255, 255, 255)">'''+str(vaga)+'''</th>
                <th  style="background-color:rgb(0, 0, 0); color:rgb(254, 254, 254); font-size:30px" rowspan="2">'''+str(caixa)+'''</th>
            </tr>
            <tr>
                <th style="background-color:rgb(0, 0, 0); color:rgb(255, 255, 255)">'''+str(armazem)+'''</th>
            </tr>
            <tr>
                <th colspan="7" rowspan="2">'''+str(descriçao)+'''</th>
            </tr>
            <tr>
            </tr>
            <tr>
                <th>'''+str(qtd)+'''</th>
                <th> </th>
                <th> </th>
                <th> </th>
                <th> </th>
                <th> </th>
                <th> </th>
            </tr>
            <tr>
                <th> </th>
                <th> </th>
                <th> </th>
                <th> </th>
                <th> </th>
                <th> </th>
                <th> </th>
            </tr>
            <tr>
                <th rowspan="2">OBS:</th>
                <th colspan="6" rowspan="2"></th>
            </tr>
            <tr>
            </tr>
            </table>

            </body>
            </html>'''

        # Salve o conteúdo em um arquivo temporário
        temp_file = 'etiqueta.html'
        with open(temp_file, 'w', encoding='utf-8') as f:
            f.write(html_content)

        # Abra o arquivo HTML no navegador padrão
        file_path = 'file://' + os.path.abspath(temp_file)
        webbrowser.open(file_path)


    @st.experimental_dialog('Lista de Seriais', width='large')
    def gerar_etiquetas_df():
        df_seriais_etiquetagem = st.dataframe(st.session_state['saldo_estoque_detalhado'][['NS',
                                                          'PN',
                                                          'DESCRICAO',
                                                          'CAIXA',
                                                          'ENDEREÇO',
                                                          'ARMAZEM',
                                                          'FILIAL',
                                                          'NF ENTRADA',
                                                          'OS INTERNA',
                                                          'PV SAIDA',
                                                          'STATUS',
                                                          'OBS']].reset_index(drop=True),
                     on_select='rerun',
                     hide_index=True)

        btse_col1, btse_col2, _ = st.columns([2,2,7])
        if btse_col1.button('Unitário'):
            imprimir_serial_unitario(df_seriais_etiquetagem.selection.rows)
        if btse_col2.button('Lista'):
            imprimir_serial_lista(df_seriais_etiquetagem.selection.rows)


    def create_df_saldo():
        hist = df_sharep(hist_sharepoint_url)
        lpn = df_sharep(lpn_sharepoint_url, 'excel')
        lpn.set_index(['PN'], inplace=True)
        hist = hist[hist['DESC FLUXO'].isin(['ARMAZENADO', 'EXPEDIÇÃO', 'LABORATÓRIO'])]
        hist = hist.join(lpn, on=['PN'], how='inner')
        return hist
    
    
    def create_df_saldo_por_endereço():
        df = df_sharep(spe_sharepoint_url, 'excel', header=1)
        df = df[['Produto', 'Armazem', 'Endereco', 'Quantidade']]
        df = df[(df['Armazem'] != 'EP') | (df['Produto'] == '6I100020')]
        df.columns = ['PN', 'ARMAZEM', 'ENDEREÇO', 'QTD SISTÊMICA']
        lpn = df_sharep(lpn_sharepoint_url, 'excel')
        lpn.set_index(['PN'], inplace=True)
        df = df.join(lpn, on=['PN'], how='inner')

        saldo = st.session_state['df_saldo']
        saldo = saldo.groupby(['PN', 'DESCRICAO', 'ARMAZEM', 'ENDEREÇO'])[['NS']].count()
        saldo.rename(columns={'NS':'QTD FÍSICA'}, inplace=True)

        df = df.join(saldo, on=['PN', 'DESCRICAO', 'ARMAZEM', 'ENDEREÇO'], how='outer')
        df.loc[df['QTD SISTÊMICA'].isna(), 'QTD SISTÊMICA'] = 0
        df['QTD SISTÊMICA'] = df['QTD SISTÊMICA'].astype(int)
        df.loc[df['QTD FÍSICA'].isna(), 'QTD FÍSICA'] = 0
        df['DELTA'] = df['QTD FÍSICA'] - df['QTD SISTÊMICA']
        df = df[df['DELTA'] == 0]
        return df


    def create_fig_status_equipamentos():
        df = st.session_state['saldo_estoque_detalhado'][['NS',
                                                          'STATUS']].copy()
        
        df = df.groupby(['STATUS'])[['NS']].count().reset_index()

        fig = px.bar(df,
                     x='STATUS',
                     y='NS',
                     color='STATUS',
                     orientation='v',
                     text='NS',
                     color_discrete_map={'SEM AVALIAÇÃO':'#004173',
                                         'DESCARREGADO':'#008000',
                                         'SEM MANUAL':'#32CD32',
                                         'EMBALAGEM AVARIADA':'#FFD700',
                                         'DESATUALIZADO':'#FF8C00',
                                         'TAMPER':'#8B0000'},
                     category_orders={'STATUS':['SEM AVALIAÇÃO', 'DESCARREGADO', 'SEM MANUAL', 'EMBALAGEM AVARIADA', 'DESATUALIZADO', 'TAMPER']})
        
        fig.update_traces(textposition='inside',
                                        orientation='v')
      
        fig.update_layout(yaxis_title=None,
                                        xaxis_title=None,
                                        yaxis_visible=False)
        
        return fig


    tabs_se, tabs_sd = st.tabs(['Saldo em Estoque', 'Saldo Divergente'])

    if 'df_saldo' not in st.session_state:
        st.session_state['df_saldo'] = create_df_saldo()

    
    tabs_se.title('Saldo em Estoque')
    bt_col1, bt_col2, _ = tabs_se.columns([2, 2, 6])
    colse1, colse2 = tabs_se.columns([5,5])
    colse1.write('Saldo resumido de equipamentos em estoque.')
    saldo_estoque = colse1.dataframe(st.session_state['df_saldo'].groupby([
        'FILIAL',
        'ARMAZEM',
        'ENDEREÇO',
        'CAIXA',
        'PN',
        'DESCRICAO'])['NS'].count().reset_index(), hide_index=True, use_container_width=True,
                 column_config={
                    "FILIAL":st.column_config.TextColumn(label="FILIAL", disabled=True),
                    "ARMAZEM":st.column_config.TextColumn(label="ARMAZÉM", disabled=True),
                    "ENDEREÇO":st.column_config.TextColumn(label="ENDEREÇO", disabled=True),
                    "CAIXA":st.column_config.TextColumn(label="CAIXA", disabled=True),
                    "PN":st.column_config.TextColumn(label="PN", disabled=True),
                    "DESCRICAO":st.column_config.TextColumn(label="DESCRIÇÃO", disabled=True),
                    "NS":st.column_config.TextColumn(label="NS", disabled=True)
                 }, on_select='rerun')
    

    if saldo_estoque.selection.rows:
        bt_etiqueta_endereço = bt_col1.button("ETIQUETA DE ENDEREÇO", use_container_width=True)
        bt_etiqueta_seriais = bt_col2.button("ETIQUETA DE SERIAIS", use_container_width=True)

        if bt_etiqueta_endereço.bit_count() > 0:
            imprimir_endereço(saldo_estoque.selection.rows)

        if bt_etiqueta_seriais.bit_count() > 0:
            gerar_etiquetas_df()

    saldo_estoque_detalhado = st.session_state['df_saldo']
    saldo_estoque_simples = st.session_state['df_saldo'].groupby([
        'FILIAL',
        'ARMAZEM',
        'ENDEREÇO',
        'CAIXA',
        'PN',
        'DESCRICAO'])['NS'].count().reset_index()
    saldo_estoque_detalhado['CONCATENADO'] = saldo_estoque_detalhado['PN'] + saldo_estoque_detalhado['ARMAZEM'] + saldo_estoque_detalhado['ENDEREÇO'] + saldo_estoque_detalhado['CAIXA']
    saldo_estoque_simples['CONCATENADO'] = saldo_estoque_simples['PN'] + saldo_estoque_simples['ARMAZEM'] + saldo_estoque_simples['ENDEREÇO'] + saldo_estoque_simples['CAIXA']
    if saldo_estoque.selection.rows:
        filtro = list(saldo_estoque_simples.iloc[saldo_estoque.selection.rows]['CONCATENADO'])
        saldo_estoque_detalhado = saldo_estoque_detalhado[saldo_estoque_detalhado['CONCATENADO'].isin(filtro)]
        st.session_state['saldo_estoque_detalhado'] = saldo_estoque_detalhado

    if 'saldo_estoque_detalhado' in st.session_state and saldo_estoque.selection.rows:
        tabs_se.write'
        tabs_se.data_editor(st.session_state['saldo_estoque_detalhado'][['NS',
                                                          'PN',
                                                          'DESCRICAO',
                                                          'CAIXA',
                                                          'ENDEREÇO',
                                                          'ARMAZEM',
                                                          'FILIAL',
                                                          'NF ENTRADA',
                                                          'OS INTERNA',
                                                          'PV SAIDA',
                                                          'DESC FLUXO',
                                                          'STATUS',
                                                          'OBS']],
                        hide_index=True,
                        use_container_width=True,
                        column_config={
                            "FILIAL":st.column_config.TextColumn(label="FILIAL", disabled=True),
                            "ARMAZEM":st.column_config.TextColumn(label="ARMAZÉM", disabled=True),
                            "ENDEREÇO":st.column_config.TextColumn(label="ENDEREÇO", disabled=True),
                            "CAIXA":st.column_config.TextColumn(label="CAIXA", disabled=True),
                            "PN":st.column_config.TextColumn(label="PN", disabled=True),
                            "DESCRICAO":st.column_config.TextColumn(label="DESCRIÇÃO", disabled=True),
                            "NS":st.column_config.TextColumn(label="NS", disabled=True),
                            "NF ENTRADA":st.column_config.TextColumn(label="NF ENTRADA", disabled=True),
                            "OS INTERNA":st.column_config.TextColumn(label="OS INTERNA", disabled=True),
                            "PV SAIDA":st.column_config.TextColumn(label="PV SAIDA", disabled=True),
                            "DESC FLUXO":st.column_config.TextColumn(label="DESC FLUXO", disabled=True),
                            "STATUS":st.column_config.TextColumn(label="STATUS", disabled=True),
                            'OBS':st.column_config.TextColumn(width='large')
                        })
        colse2.write('Distribuição do status dos equipamentos.')
        colse2.plotly_chart(create_fig_status_equipamentos())

    tabs_sd.title('Saldo Divergente')

    col1, col2 = tabs_sd.columns([5,5])

    if 'df_saldo_por_endereço' not in st.session_state:
        st.session_state['df_saldo_por_endereço'] = create_df_saldo_por_endereço()

    saldo_divergente = col1.dataframe(st.session_state['df_saldo_por_endereço'][['PN',
                                                                      'DESCRICAO',
                                                                      'ARMAZEM',
                                                                      'ENDEREÇO',
                                                                      'QTD FÍSICA',
                                                                      'QTD SISTÊMICA',
                                                                      'DELTA']], hide_index=True, use_container_width=True, on_select='rerun')

    spe_detalhado = st.session_state['df_saldo'].copy()
    spe_simples = st.session_state['df_saldo_por_endereço'].copy()
    spe_detalhado['CONCATENADO'] = spe_detalhado['PN'] + spe_detalhado['ARMAZEM'] + spe_detalhado['ENDEREÇO']
    spe_simples['CONCATENADO'] = spe_simples['PN'] + spe_simples['ARMAZEM'] + spe_simples['ENDEREÇO']
    if saldo_divergente.selection.rows:
        filtro = list(spe_simples.iloc[saldo_divergente.selection.rows]['CONCATENADO'])
        spe_detalhado = spe_detalhado[spe_detalhado['CONCATENADO'].isin(filtro)]
        st.session_state['spe_detalhado'] = spe_detalhado

    if 'spe_detalhado' in st.session_state and saldo_divergente.selection.rows:
        col2.dataframe(st.session_state['spe_detalhado'][['NS',
                                                          'PN',
                                                          'CAIXA',
                                                          'ENDEREÇO',
                                                          'ARMAZEM',
                                                          'FILIAL',
                                                          'PV SAIDA',
                                                          'DESC FLUXO',
                                                          'OBS']],
                        hide_index=True,
                        use_container_width=True,
                        column_config={
                            'OBS':st.column_config.TextColumn(width='large')
                        })
