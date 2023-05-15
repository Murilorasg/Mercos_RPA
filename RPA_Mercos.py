import os, sys
import pandas as pd
import datetime
import dados
from time import sleep
from playwright.sync_api import sync_playwright

class Projeto():
    
# Recursos ------------------------------------------------------------------------------------------
    
    def __init__(self):
        
        self.dict_pedidos = {
            'id_pedido':str,
            'cnpj':str,
            'nome_cliente':str,
            'representada':str,
            'estado':str,
            'cod_produto':str,
            'descricao_produto':str,
            'qnt':int,
            'preco':float,
            'desconto':float,
            'cond_pagamento':str,
            'transportadora':str,
        }
        
        self.pedidos_geral = pd.DataFrame(columns=self.dict_pedidos.keys())
        
        Projeto.set_variaveis_caminhos(self)
        
    def set_variaveis_caminhos(self):
            
            # ----------------------
            
            # // VARIÁVEIS DATA   
            try:
                
                self.caminho_raiz = (sys.argv[0].replace('RPA_Mercos.py',""))
                self.caminho_pastas = (sys.argv[0].replace('RPA_Mercos.py',"")+'Execucoes')
                self.caminho_layout = (sys.argv[0].replace('RPA_Mercos.py',"")+'Layout')
                self.data_processamento = datetime.datetime.today()
                self.data_hora = self.data_processamento.strftime("%Y_%m_%d__%H_%M")
                self.pasta_arquivos = os.path.join(self.caminho_pastas,self.data_hora)
                
                if os.path.isdir(self.pasta_arquivos):
                    pass
                else:
                    os.mkdir(self.pasta_arquivos)
                    
                return True
            
            except Exception as e:
                
                return False

    
    def le_excel(self, caminho):
        
        df = pd.read_excel(caminho, header=0)

        return df
    
    def grava_excel(self, df, caminho):
        
        df.to_excel(caminho, index=False)
        
        return
    
    def cria_layout(self, df, caminho):
        
        df.to_excel(caminho, index=False)
        
        return
        
    def trata_dados_cnpj(self, cnpj):
        
        cnpj = cnpj.replace(" ","")
        cnpj = cnpj.replace("-","")
        cnpj = cnpj.replace(".","")
        cnpj = cnpj.replace("/","")
        
        return cnpj
    
    def trata_dados_estado(self, estado):
        
        estado = estado.replace(" ","")
        separa = estado.split(',')
        estado = separa[1]
        
        
        return estado
    
    def trata_dados_produtos(self, preco, desconto):
        
        preco = preco.replace("R","")
        preco = preco.replace("$","")
        preco = preco.replace(" " ,"")
        preco = preco.replace(",",".")
        preco = float(preco)
        

        desconto = desconto.replace(",",".")
        desconto = desconto.replace("'","")
        digit = 0
        for digit in range(len(desconto)):
            if desconto[digit].isnumeric():
                digit+1
            elif ((digit == 2) & (desconto[digit] == '.')):
                digit+1
            else:
                break
        desconto = desconto[0:digit]
        
        if desconto == "":
            desconto = 0
            
        desconto = float(desconto)
        
        dados = [preco,desconto]
        
        return dados
    
# Mercos -----------------------------------------------------------------------------------------
                
    def deve_logar_mercos(self,browser):
        
        context_mercos = browser.new_context(locale='pt-BR', timezone_id="America/Sao_Paulo")
        page_mercos = context_mercos.new_page()
        
        url = "https://app.mercos.com/login"
        username = dados.mercos_lg
        password = dados.mercos_ps
        
        page_mercos.goto(url)
        page_mercos.fill('//*[@id="id_usuario"]', username)
        page_mercos.fill('//*[@id="id_senha"]', password)
        page_mercos.click('//*[@id="botaoEfetuarLogin"]')
        
        return page_mercos
        
    def seleciona_pedidos_mercos(self, page_mercos):
        
        for id_pedido in self.df_pedidos_mercos.iterrows():
            
            page_mercos.click('//*[@id="aba_pedidos"]/span')
            id_pedido = str(id_pedido[1][0])
            page_mercos.fill('//*[@id="id_texto"]', id_pedido)
            page_mercos.click('//*[@id="form_pesquisa_normal"]/div[1]/button')
            page_mercos.click('//*[@id="js-div-global"]/div[2]/section/div[2]/div[1]/div[4]/div[2]/div[1]')
            cnpj = page_mercos.locator('//*[@id="selecionado_autocomplete_id_codigo_cliente"]/span/div/h5/small[2]').inner_text()
            cnpj = Projeto.trata_dados_cnpj(cnpj)
            estado = page_mercos.locator('//*[@id="selecionado_autocomplete_id_codigo_cliente"]/span/div/div[2]/div/span').inner_text()
            estado = Projeto.trata_dados_estado(estado)
            nome_cliente = page_mercos.locator('//*[@id="selecionado_autocomplete_id_codigo_cliente"]/span/div/h5/small[1]').inner_text()
            representada = page_mercos.locator('//*[@id="nome-representada-selecionada"]').inner_text()
            cond_pagamento = page_mercos.locator('//*[@id="informacoes_complementares"]/div/div/div[2]/div/div/div[2]').inner_text()
            transportadora = page_mercos.locator('//*[@id="informacoes_complementares"]/div/div/div[3]/div[2]/div/div[2]').inner_text()
            mostrar_todos = page_mercos.locator('//*[@id="listagem_item"]/div[2]/a').inner_text()
            if 'mostrar todos' in mostrar_todos.lower():
                page_mercos.click('//*[@id="listagem_item"]/div[2]/a')
            tr = page_mercos.query_selector_all("tr")
            for linha in tr:
                try:
                    classe = linha.get_attribute('class')
                    if classe != None:
                        if 'dados_item' in classe:
                            td = linha.query_selector_all("td")
                            cod_produto = td[1].inner_text()
                            descricao_produto = td[2].inner_text()
                            qnt = td[3].inner_text()
                            preco = td[6].inner_text()
                            desconto = td[5].inner_text()
                            
                            list_dados = Projeto.trata_dados_produtos(preco,desconto)
                            preco = list_dados[0]
                            desconto = list_dados[1]
                            
                            self.pedidos_geral.loc[len(self.pedidos_geral)] = (id_pedido, cnpj, nome_cliente, representada, estado, cod_produto, descricao_produto, qnt, preco, desconto, 
                                                                        cond_pagamento, transportadora)
                        else:
                            pass
                    else:
                        pass
                except:
                    pass       
         
        Projeto.grava_excel(self.pedidos_geral,self.pasta_arquivos+'/pedidos_geral.xlsx')
        Projeto.grava_excel(self.pedidos_geral,self.caminho_layout+'/layout.xlsx')
                
        return
    
# Opus ------------------------------------------------------------------------------------------

    def deve_logar_opus(self, browser):
        
        context_opus = browser.new_context(locale='pt-BR', timezone_id="America/Sao_Paulo")
        page_opus = context_opus.new_page()
        
        url = "http://opusled.ddns.mobi:84/"
        username = dados.opus_lg
        password = dados.opus_ps
        
        page_opus.goto(url)
        page_opus.fill('//*[@id="UserName"]', username)
        page_opus.fill('//*[@id="Password"]', password)
        page_opus.click('//*[@id="btnLogin"]')
        sleep(3)
        
        return page_opus
    
    def deve_digitar_pedidos_opus(self, page_opus):
        
        dict_result_pedidos_opus = {
            'pedido':str,
            'status':str,
        }
        
        self.df_result_pedidos_opus = pd.DataFrame(columns=dict_result_pedidos_opus.keys())
        
        if len(self.pedidos_geral) == 0:
            
            self.pedidos_geral = Projeto.le_excel(self.caminho_layout+'/layout.xlsx')
            
        pedidos_opus = self.pedidos_geral.loc[self.pedidos_geral['representada']=='OpusLed']
        # pedidos = list(map(int,pedidos_opus['id_pedido'].drop_duplicates()))  
        pedidos = list(pedidos_opus['id_pedido'].drop_duplicates())
        
        
        for pedido in pedidos:
            try:
                order = self.pedidos_geral.loc[self.pedidos_geral['id_pedido'] == pedido]
                order = order.reset_index()
                try:
                    page_opus.goto('http://opusled.ddns.mobi:84/')
                    page_opus.click('//*[@id="mainpage"]/div[2]/div[1]/nav/ul/li[4]/a')
                    page_opus.click('//*[@id="btnBuscarCliente"]')
                    page_opus.fill('//*[@id="Filter_PARC_CGC"]', str(order['cnpj'][0]))
                    page_opus.click('//*[@id="btnPesquisarCliente"]')
                except:
                    button = page_opus.query_selector_all('[data-popup-button]')[1].click()
                    page_opus.click('//*[@id="btnBuscarCliente"]')
                    page_opus.fill('//*[@id="Filter_PARC_CGC"]', str(order['cnpj'][0]))
                    page_opus.click('//*[@id="btnPesquisarCliente"]')
                    
                page_opus.click('//*[@id="btnPesquisarCliente"]')
                page_opus.click('//*[@id="listaLocRapCliente"]/li[2]/a')
                if order['estado'][0] == 'SãoPaulo':
                    page_opus.locator('#Pedido_UNID_COD').select_option('OPERA SOLUCOES')           
                else:
                    page_opus.locator('#Pedido_UNID_COD').select_option('OPUS SISTEMAS')
                sleep(3)
                page_opus.fill('//*[@id="Pedido_PEDS_EXT_CODCLI"]', '.')
                order['cond_pagamento'][0] = '28/35' #-----------------retirar
                page_opus.locator('#Pedido_COPG_COD_001').select_option((order['cond_pagamento'][0]+' DIAS')) 
                page_opus.click('//*[@id="identificacao-pedido"]/div[1]/a[2]')
                sleep(3)
                
                for row, item in order.iterrows():
                    
                    page_opus.fill('//*[@id="edtPROD_EXT_COD"]', item['cod_produto'])
                    page_opus.click('//*[@id="btnPesquisarProduto"]')
                    sleep(3)
                    page_opus.fill('//*[@id="Pedido_ItemAtual_ITPS_QTD_PED"]', str(item['qnt']))
                    page_opus.fill('//*[@id="Pedido_ItemAtual_ITPS_VLF_PRELIQ"]', str(item['preco']))
                    page_opus.fill('//*[@id="Pedido_ItemAtual_ITPS_EXT_CODCLI"]', '.')
                    page_opus.click('//*[@id="btnSalvar"]')
                    
                page_opus.click('//*[@id="itens-pedido"]/div[1]/a[2]')
                order['transportadora'][0]='TRANSVOAR' #-----------------retirar
                page_opus.locator('#Pedido_SERT_COD').select_option(order['transportadora'][0])
                page_opus.click('//*[@id="info-adicionais"]/div[1]/a[2]')
                # page_opus.click('//*[@id="revisao-pedido"]/div[2]/div[3]/a[2]')      
                
                self.df_result_pedidos_opus.loc[len(self.df_result_pedidos_opus)] = (pedido, 'OK')                
                
            except Exception as e:
                
                self.df_result_pedidos_opus.loc[len(self.df_result_pedidos_opus)] = (pedido, f'ERRO: {e}')
                pass
                
        Projeto.grava_excel(self.df_result_pedidos_opus, self.pasta_arquivos+'/result_pedidos_opus.xlsx')
            
        page_opus.close()            
            
# Execucao -------------------------------------------------------------------------------------- 
            
    def roteiro_mercos(self,browser):
        
        page = Projeto.deve_logar_mercos(browser)
        page = Projeto.seleciona_pedidos_mercos(page)
        return page
    
    def roteiro_opus(self,browser):

        page = Projeto.deve_logar_opus(browser)
        Projeto.deve_digitar_pedidos_opus(page)
        
    def roteiro_finaliza(self,page_mercos):
        
        if len(self.df_result_pedidos_opus) > 0:
            
            pedidos_ok = self.df_result_pedidos_opus.loc[self.df_result_pedidos_opus['status']=='OK']
            
            for row, pedido in pedidos_ok.iterrows():
                
                page_mercos.click('//*[@id="aba_pedidos"]/span')                
                page_mercos.fill('//*[@id="id_texto"]', pedido['pedido'][row])
                page_mercos.click('//*[@id="form_pesquisa_normal"]/div[1]/button')
                page_mercos.click('//*[@id="js-div-global"]/div[2]/section/div[2]/div[1]/div[4]/div[2]/div[1]')
                page_mercos.click('//*[@id="js-div-global"]/div[3]/section/div[3]/div[10]/div[3]/button[1]')
                sleep(3)
                
            page_mercos.close()
                            
    def execute(self):

        self.df_pedidos_mercos = Projeto.le_excel(self.caminho_raiz+'/pedidos.xlsx')
        if len(self.df_pedidos_mercos) > 0:
            with sync_playwright() as playwright:
                chromium = playwright.chromium
                browser = chromium.launch(headless=False)
                page_mercos = Projeto.roteiro_mercos(browser)
                Projeto.roteiro_opus(browser)
                Projeto.roteiro_finaliza(page_mercos)
        else:
            print('Nenhum pedido para ser transmitido')
            
        
Projeto = Projeto()
Projeto.execute()