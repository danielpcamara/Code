import pandas as pd
import numpy as np
import json
from datetime import datetime
import re
import sqlite3

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

from time import sleep

# Usuario  seq, nome, permissa, impostos, transferencia, email
# Empresa id, nome, cadastro, ativo, razao, login, inscf, elamins, telefone, usuarios, endereco, numero, complemento, bairro, cep, uf, municipio

class eContabilSite:

    def __init__(self, dbg=False):
        chrome_options = webdriver.ChromeOptions()
        settings = {
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2
        }

        with open('config.json') as file:
            files_folder = json.load(file)['files_folder']

        prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                'savefile.default_directory': files_folder}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument('--kiosk-printing')

        self.browser = webdriver.Chrome(options=chrome_options)
        self.main_window = self.browser.window_handles[0]
        self.login()
        self.session = self.browser.current_url.split("?")[1]
        self.cnx = sqlite3.connect('scraped.db')
        self.dbg = dbg

    @property
    def main_url(self):
        return "https://www.centraldecontabilidade.com.br/econtabil"
    
    def create_tables(self):
        db = sqlite3.connect('scraped.db')
        cursor=db.cursor()
        comando = """
            CREATE TABLE IF NOT EXISTS Andamento (
                id, 
                compet, 
                status, 
                PRIMARY KEY (id, compet)
            );
        """

        cursor.execute(comando)
        db.commit()
        db.close()
        pass
    
    def reprint(self, tx):
        tx = '\r' + tx + ' ' * 200
        tx = tx[:199]
        print(tx, end='')
        pass

    def update_andamento(self, code, compet, status):
        db = sqlite3.connect('scraped.db')
        cursor=db.cursor()
        comando = f"INSERT OR IGNORE INTO Andamento (id, compet, status) values ('{code}','{compet}','{status}')"
        cursor.execute(comando)
        comando = f"UPDATE Andamento SET status = '{status}' WHERE id = '{code}' and compet = '{compet}'"
        cursor.execute(comando)
        db.commit()
        db.close()
        pass

    def get_andamento_status(self, code, compet):
        con = sqlite3.connect('scraped.db')
        cursor = con.cursor()
        cursor.execute(f"SELECT status FROM Andamento where id = '{code}' and compet = '{compet}';")
        try:
            x = cursor.fetchall()[0][0]
        except:
            x = None
        return x

    def save_to_db (self, df, table_name, if_exists='fail'): #base
        df.to_sql(name=table_name, con=self.cnx, if_exists=if_exists)
        pass

    def drop_table(self, table_name):
        cursor=self.cnx.cursor()
        command = f'DROP TABLE IF EXISTS {table_name.upper()};'
        cursor.execute(command)
        self.cnx.commit
        pass
    
    def get_record_from_page(self, id, fields, dbg=None): #base
        # fields format: [[query(ex: '//*[@id="txtEmpresa"]'), value_to_get(ex: get_attribute, text, option, check), extra_parameter(use un attribute ex: value)]]
        if dbg == None:
            dbg = self.dbg
      
        to_add=[id]

        for field in fields:
            try:
                elm = self.browser.find_element(By.XPATH, field[0])

                if field[1] == 'get_attribute':
                    value = elm.get_attribute(field[2])
                elif field[1] == 'text':
                    value = elm.text
                elif field[1] == 'option':
                    value = Select(elm).first_selected_option.text
                elif field[1] == 'check':
                    value = elm.is_selected()
            except NoSuchElementException as e:
                print(f'Error field: {field[0]}; msg: {str(e)[slice(100)]}')
                value = np.nan

            to_add.append(value)            
            
        return to_add

    def get_data(self, elms_xpath, fields, column_names, dbg=None, type_data='df'): #base
        
        if dbg == None:
            dbg = self.dbg

        def get_value_from_element(elm_parent, field):
            
            try:
                elm = elm_parent.find_element(By.XPATH, field[0])
                if field[1] == 'get_attribute':
                    value = elm.get_attribute(field[2])
                elif field[1] == 'text':
                    value = elm.text
                elif field[1] == 'option':
                    value = Select(elm).first_selected_option.text
            except Exception as e:
                print(f'Error field: {field[0]}; msg: {str(e)}')
                value = np.nan
            return value
        
        pages_elem = self.browser.find_elements(By.XPATH, elms_xpath)

        df_list = []

        for elm in pages_elem:
            try:
                to_add=[]
                id = get_value_from_element(elm, fields[0])
                to_add.append(id)
                
                if id.replace('.','').strip().isdigit():
                    for field in fields[1:]:
                        next_field = get_value_from_element(elm, field)
                        to_add.append(next_field)

                    df_list.append(to_add)
                else:
                    raise
            except:
                pass
            if dbg:
                print(to_add)
        if type_data == 'df':
            df = pd.DataFrame(df_list, columns=column_names)
        else:
            df = df_list
        return df

    def login(self):
        with open('config.json') as file:
            credentials = json.load(file)

        self.browser.get(self.main_url + "/login.asp?log_=timeout&d=" + credentials['domain'])
        self.browser.find_element(By.ID, 'txtLogin').send_keys(credentials['user'])
        self.browser.find_element(By.NAME, 'txtSenha').send_keys(credentials['key'])
        self.browser.find_element(By.NAME, 'Submit').click()

        WebDriverWait(self.browser, 90).until(EC.presence_of_element_located((By.XPATH, '/html/body/table[2]/tbody/tr/td[2]/table/tbody/tr[2]/td/img')))
    
    def get_clients(self, dbg=None): # ok
        
        if dbg == None:
            dbg = self.dbg
        def load_page(self):
            self.browser.get(self.main_url + "/adm/clientes.asp?" + self.session + "&q=2&a=consultar")
            self.browser.execute_script("document.getElementById('txtCnpj_Pesq').value='PESQUISA';document.getElementById('cmdConsulta').click();")

        def get_pages_commands(self):
            pages_elem = self.browser.find_elements(By.XPATH, '//*[@id="tb_resultado"]/tbody/tr[2]/td/span')

            pages_to_check = []
            for elm in pages_elem:
                pages_to_check.append(elm.get_attribute("onclick"))

            return pages_to_check
        
        def get_companys_commands(self, pages_to_check):
            compays_to_check = []
            for page in pages_to_check:
                self.browser.execute_script(page)
                companys_elem = self.browser.find_elements(By.XPATH, '//*[@id="tb_resultado"]/tbody/tr')
                for elm in companys_elem:
                    compays_to_check.append(elm.get_attribute("onclick"))

            compays_to_check = [x for x in compays_to_check if x is not None]
            compays_to_check.pop(0) #remove 1st (new)
            return compays_to_check

        def feed_company_df(self, dfs, command, dbg):

            self.browser.execute_script(command)
            dfc = dfs[0]
            dfu = dfs[1]
            id = command.split("(")[1].replace(")", '')
            fields = [
                ['//*[@id="txtEmpresa"]', 'get_attribute', 'value'],
                ['//*[@id="tb_dados_empresa"]/tbody/tr[3]/td[5]', 'text', ''],
                ['//*[@id="txt_fAtiva"]', 'option', ''],
                ['//*[@id="txtRazaoSocial"]', 'get_attribute', 'value'],
                ['//*[@id="cboIc_Ent_Financeira"]', 'option', ''],
                ['//*[@id="txt_carteira"]', 'option', ''],
                ['//*[@id="txtLogin"]', 'get_attribute', 'value'],
                ['//*[@id="txtSenha"]', 'get_attribute', 'value'],
                ['//*[@id="txtCnpj"]', 'get_attribute', 'value'],
                ['//*[@id="txtCnae"]', 'get_attribute', 'value'],
                ['//*[@id="txt_nat_jurid"]', 'get_attribute', 'value'],
                ['//*[@id="txt_ramo"]', 'get_attribute', 'value'],
                ['//*[@id="txt_capital_social"]', 'get_attribute', 'value'],
                ['//*[@name="txt_dt_abertura"]', 'get_attribute', 'value'],
                ['//*[@id="txt_registro_num"]', 'get_attribute', 'value'],
                ['//*[@id="txt_registro_org"]', 'get_attribute', 'value'],
                ['//*[@id="txt_resp_nome"]', 'get_attribute', 'value'],
                ['//*[@id="txt_resp_cpf"]', 'get_attribute', 'value'],
                ['//*[@id="txt_resp_cod_qualif"]', 'option', ''],
                ['//*[@id="txtEndereco"]', 'get_attribute', 'value'],
                ['//*[@id="txtNr_End"]', 'get_attribute', 'value'],
                ['//*[@id="txtDc_Compl_End"]', 'get_attribute', 'value'],
                ['//*[@id="txtbairro"]', 'get_attribute', 'value'],
                ['//*[@id="txtCEP"]', 'get_attribute', 'value'],
                ['//*[@id="txtUF"]', 'option', ''],
                ['//*[@id="txt_id_municipio_cnpj"]', 'option', ''],
                ['//*[@name="txt_contato"]', 'get_attribute', 'value'],
                ['//*[@name="txtFone"]', 'get_attribute', 'value'],
                ['//*[@name="txt_email"]', 'get_attribute', 'value'],
                ['//*[@name="txt_fax"]', 'get_attribute', 'value'],
                ['//*[@name="txt_site"]', 'get_attribute', 'value'],
                ['//*[@id="txt_tipo"]', 'option', ''],
                ['//*[@name="txt_tipo_esp"]', 'option', ''],
                ['//*[@name="txt_serv_mo"]', 'option', ''],
                ['//*[@id="txt_regime_federal"]', 'option', ''],
                ['//*[@id="txt_regime_federal_esp_simples"]', 'option', ''],
                ['//*[@id="txt_regime_federal_esp"]', 'option', ''],
                ['//*[@id="txt_recolhe_irpj"]', 'option', ''],
                ['//*[@id="txtIEstadual"]', 'get_attribute', 'value'],
                ['//*[@id="txt_icms_esp"]', 'option', ''],
                ['//*[@id="txt_codigo_ref_sp"]', 'option', ''],
                ['//*[@name="txt_codigo_ref"]', 'get_attribute', 'value'],
                ['//*[@id="txtIMunicipal"]', 'get_attribute', 'value'],
                ['//*[@id="txt_ISS_dia_vcto"]', 'get_attribute', 'value'],
                ['//*[@id="txt_iss_tipo"]', 'get_attribute', 'value'],
                ['//*[@id="txtCodGps"]', 'get_attribute', 'value'],
                ['//*[@id="txtGFIP_cod_recolhim"]', 'get_attribute', 'value'],
                ['//*[@id="txtGFIP_FPAS"]', 'get_attribute', 'value'],
                ['//*[@name="txtGFIP_cod_outras_entid"]', 'get_attribute', 'value'],
                ['//*[@id="txtGFIP_simples"]', 'get_attribute', 'value'],
                ['//*[@id="txtcliente_desde"]', 'get_attribute', 'value'],
                ['//*[@id="txtmes_reaj"]', 'option', ''],
                ['//*[@id="txtvlr_honor"]', 'get_attribute', 'value'],
                ['//*[@id="txt_reemb_impostos"]', 'option', ''],
                ['//*[@id="txt_pag_honor"]', 'option', ''],
                ['//*[@id="txtvlr_desc"]', 'get_attribute', 'value'],
                ['//*[@id="txtvlr_13_mensal"]', 'get_attribute', 'value'],
                ['//*[@id="txt_ref_boleto"]', 'option', ''],
                ['//*[@id="txtdia_vcto_honor"]', 'get_attribute', 'value'],
                ['//*[@id="txtMinNF"]', 'get_attribute', 'value'],
                ['//*[@id="txtMaxNF"]', 'get_attribute', 'value'],
                ['//*[@id="txtSerie"]', 'option', ''],
                ['//*[@id="txtSerie_Padrao"]', 'get_attribute', 'value'],
                ['//*[@id="txtSerie_Outras"]', 'get_attribute', 'value'],
                ['//*[@id="checa_padrao"]', 'check', ''],
                ['//*[@name="txtObrigaTomador"]', 'option', ''],
                ['//*[@id="txtSerie_Venda"]', 'option', ''],
                ['//*[@id="txtSerie_Venda_Outras"]', 'get_attribute', 'value'],
                ['//*[@id="txtCupom_Equip"]', 'option', ''],
                ['//*[@id="txtCupom_Equip_Outras"]', 'get_attribute', 'value'],
                ['//*[@id="txtCupom_CFOP_tribu"]', 'get_attribute', 'value'],
                ['//*[@id="txtCupom_CFOP_nao_trib"]', 'get_attribute', 'value'],
                ['//*[@id="txtDuplicaContabil"]', 'option', ''],
                ['//*[@id="txtMes"]', 'option', ''],
                ['//*[@id="txtAno"]', 'option', ''],
                ['//*[@id="cboIcImpCtrExp"]', 'option', ''],
                ['//*[@id="txt_apelido"]', 'get_attribute', 'value'],
                ['//*[@id="txt_apelido_folha"]', 'get_attribute', 'value'],
                ['//*[@id="txtCod_Etb"]', 'get_attribute', 'value'],
                ['//*[@id="txtCd_Integ_Contab"]', 'get_attribute', 'value'],
                ['//*[@id="txt_cd_atv"]', 'get_attribute', 'value'],
                ['//*[@id="txtCodigo_tabela_eventos"]', 'get_attribute', 'value'],
                ['//*[@id="txt_obs"]', 'text', '']
            ]

            clist = self.get_record_from_page(id, fields)

            if dbg:
                print("company: ", clist)
            dfc.loc[len(dfc)] = clist

            dfu = feed_user_df(self, id, dfu, dbg)
            dfu = None
            dfs = [dfc, dfu]

            return dfs

        def feed_user_df(self, id, dfu, dbg):
            #To call this def, you need to be in the company's page.
            self.browser.execute_script("btnUsuario_onclick();")

            users_elem = self.browser.find_elements(By.XPATH, '//*[@id="dv_user"]/table/tbody/tr')

            for elm in users_elem:
                try:
                    seq = elm.find_element(By.XPATH, './/td[1]').text
                    if seq.isdigit():
                        nome = elm.find_element(By.XPATH, './/td[2]').text
                        premissao = elm.find_element(By.XPATH, './/td[3]').text
                        impostos = elm.find_element(By.XPATH, './/td[4]').text
                        transferencia = elm.find_element(By.XPATH, './/td[5]').text
                        email = elm.find_element(By.XPATH, './/td[6]').text
                        ulist = [id, seq, nome, premissao, impostos, transferencia, email]
                        if dbg:
                            print('user:', ulist)
                        dfu.loc[len(dfu)] = ulist
                except:
                    pass


            return dfu

        def save_to_df(self, dfs):
            dfc = dfs[0]
            dfu = dfs[1]
            dfc.to_sql(name='Clientes', con=self.cnx)
            dfu.to_sql(name='Usuários', con=self.cnx)
            pass

        load_page(self)
        pages_to_check = get_pages_commands(self)
        companys_to_check = get_companys_commands(self, pages_to_check)

        dfc = pd.DataFrame(columns=['id', 'v_nome', 'v_cadastro', 'ativo', 'v_razao', 'entidade_financeira', 'v_carteira', 'v_login', 'senha', 'v_insc_fed', 'v_cnae', 'v_nat_jur', 'txt_ramo', 'txt_capital_social', 'txt_dt_abertura', 'txt_registro_num', 'txt_registro_org', 'txt_resp_nome', 'txt_resp_cpf', 'txt_resp_cod_qualif', 'txtEndereco', 'txtNr_End', 'txtDc_Compl_End', 'txtbairro', 'txtCEP', 'txtUF', 'txt_id_municipio_cnpj', 'txt_contato', 'txtFone', 'txt_email', 'txt_fax', 'txt_site', 'txt_tipo', 'txt_tipo_esp', 'txt_serv_mo', 'txt_regime_federal', 'txt_regime_federal_esp_simples', 'txt_regime_federal_esp', 'txt_recolhe_irpj', 'txtIEstadual', 'txt_icms_esp', 'txt_codigo_ref_sp', 'txt_codigo_ref', 'txtIMunicipal', 'txt_ISS_dia_vcto', 'txt_iss_tipo', 'txtCodGps', 'txtGFIP_cod_recolhim', 'txtGFIP_FPAS', 'txtGFIP_cod_outras_entid', 'txtGFIP_simples', 'txtcliente_desde', 'txtmes_reaj', 'txtvlr_honor', 'txt_reemb_impostos', 'txt_pag_honor', 'txtvlr_desc', 'txtvlr_13_mensal', 'txt_ref_boleto', 'txtdia_vcto_honor', 'txtMinNF', 'txtMaxNF', 'txtSerie', 'txtSerie_Padrao', 'txtSerie_Outras', 'checa_padrao', 'txtObrigaTomador', 'txtSerie_Venda', 'txtSerie_Venda_Outras', 'txtCupom_Equip', 'txtCupom_Equip_Outras', 'txtCupom_CFOP_tribu', 'txtCupom_CFOP_nao_trib', 'txtDuplicaContabil', 'txtMes', 'txtAno', 'cboIcImpCtrExp', 'txt_apelido', 'txt_apelido_folha', 'txtCod_Etb', 'txtCd_Integ_Contab', 'txt_cd_atv', 'txtCodigo_tabela_eventos', 'txt_obs'])
        dfu = pd.DataFrame(columns=['ID', 'Seq', 'Nome', 'Permissao', 'Impostos', 'Transferencia', 'Email'])
        dfs = [dfc, dfu]
        for company_command in companys_to_check:
            dfs = feed_company_df(self, dfs, company_command, dbg)

        save_to_df(self, dfs)
        return dfs

    def get_movfolha(self, dbg=None): # ok
        
        if dbg == None:
            dbg = self.dbg

        def load_page(self):
            self.browser.get(self.main_url + "/adm/movfolha.asp?" + self.session + "&q=1")

        def get_eventos(self,dgb):
            pages_elem = self.browser.find_elements(By.XPATH, '//*[@id="frmClientes"]/table[2]/tbody/tr')
            eventos = []
            for elm in pages_elem:
                f1 = elm.find_element(By.XPATH, './/td').text
                if f1 == "PROVENTOS" or f1 == "DESCONTOS":
                    tipo = f1
                try:
                    id = elm.find_element(By.XPATH, './/td[1]/input').get_attribute('value')
                    if id.replace('.','').strip().isdigit():
                        desc = elm.find_element(By.XPATH, './/td[2]').text
                        uni = elm.find_element(By.XPATH, './/td[3]').text
                        lista = elm.find_element(By.XPATH, './/td[4]').text
                        seq = elm.find_element(By.XPATH, './/td[5]').text
                        qtd_config = elm.find_element(By.XPATH, './/td[6]').text
                        itens_adicionar = [id, tipo, desc, uni, lista, seq, qtd_config]
                        eventos.append(itens_adicionar)
                    else:
                        raise
                except:
                    id = np.nan
                    desc = np.nan
                    uni = np.nan
                    lista = np.nan
                    seq = np.nan
                    qtd_config = np.nan
                    itens_adicionar = []
                    pass
                if dbg:
                    print(itens_adicionar)
            
            df = pd.DataFrame(eventos, columns=['Cod', 'Tipo', 'Descrição', 'Unidade', 'Lista', 'Seq', 'Qtd Configurado no Cliente'])
            return df
        
        def save_to_db (self, df):
            df.to_sql(name='Eventos', con=self.cnx)

        load_page(self)
        df = get_eventos(self, dbg)
        save_to_db(self, df)
        pass

    def get_outros_pag(self, dbg=None): # ok

        if dbg == None:
            dbg = self.dbg

        def load_page(self):
            self.browser.get(self.main_url + "/adm/pagamentos_cadastro.asp?" + self.session + "&q=2")

        def get_o_pgto(self, dbg):
            pages_elem = self.browser.find_elements(By.XPATH, '//*[@id="frmClientes"]/table/tbody/tr')

            outros_pagamentos = []
            for elm in pages_elem:
                try:
                    id = elm.find_element(By.XPATH, './/td[1]').text
                    if id.replace('.','').strip().isdigit():
                        pagamento = elm.find_element(By.XPATH, './/td[2]').text
                        lista = elm.find_element(By.XPATH, './/td[3]').text
                        qtd_usado = elm.find_element(By.XPATH, './/td[4]').text
                        tipo = elm.find_element(By.XPATH, './/td[5]').text
                        itens_adicionar = [id, pagamento, lista, qtd_usado, tipo]
                        outros_pagamentos.append(itens_adicionar)
                    else:
                        raise
                except:
                    pass
                if dbg:
                    print(itens_adicionar)
            
            df = pd.DataFrame(outros_pagamentos, columns=['Id', 'Pagamento', 'Lista', 'Qtd_Usado', 'Tipo'])
            return df

        def save_to_db (self, df):
            df.to_sql(name='Outros_Pagamentos', con=self.cnx)
        
        load_page(self)
        df = get_o_pgto(self, dbg)
        save_to_db(self, df)
        pass
    
    def get_solicitacoes(self, dbg=None): # ok

        if dbg == None:
            dbg = self.dbg

        def load_page(self):
            self.browser.get(self.main_url + "/adm/solicitacoes_cadastro.asp?" + self.session + "&q=1")

        def get_soliciacoes(self, dbg):
            pages_elem = self.browser.find_elements(By.XPATH, '//*[@id="DivSolicita"]/table[1]/tbody/tr')

            solicitaceos = []
            itens_adicionar=[]
            for elm in pages_elem:
                try:
                    id = elm.find_element(By.XPATH, './/td[1]/b').text
                    if id.replace('.','').strip().isdigit():
                        tipo = elm.find_element(By.XPATH, './/td[2]').text
                        departamento = elm.find_element(By.XPATH, './/td[3]').text
                        qtd_solicitado = elm.find_element(By.XPATH, './/td[4]').text
                        itens_adicionar = [id, tipo, departamento, qtd_solicitado]
                        solicitaceos.append(itens_adicionar)
                    else:
                        raise
                except:
                    pass
                if dbg:
                    print(itens_adicionar)
            
            df = pd.DataFrame(solicitaceos, columns=['Id', 'Tipo', 'Departamento', 'Qtd Solicitado'])
            self.save_to_db(df, 'Tipos_Solicitações')
            return df

        def get_locais(self, dbg):
            pages_elem = self.browser.find_elements(By.XPATH, '//*[@id="DivSolicita"]/table[2]/tbody/tr')

            locais = []
            itens_adicionar=[]
            for elm in pages_elem:
                try:
                    id = elm.find_element(By.XPATH, './/td[1]/b').text 
                    if id.replace('.','').strip().isdigit():
                        desc = elm.find_element(By.XPATH, './/td[2]').text
                        qtd_solicitado = elm.find_element(By.XPATH, './/td[3]').text
                        itens_adicionar = [id, desc, qtd_solicitado]
                        locais.append(itens_adicionar)
                    else:
                        raise
                except:
                    pass
                if dbg:
                    print(itens_adicionar)
            
            df = pd.DataFrame(locais, columns=['Id', 'Descrição', 'Qtd Solicitado'])
            self.save_to_db(df, 'Locais_Cadastrados')
            return df
        
        load_page(self)
        get_soliciacoes(self, dbg)
        get_locais(self, dbg)
        pass

    def get_user(self, dbg=None): # ok

        if dbg == None:
            dbg = self.dbg

        def load_page(self):
            self.browser.get(self.main_url + "/adm/senhas.asp?" + self.session + "&q=2")

        def get_user_data(self, dbg):
            
            elms_xpath = '//*[@id="frmClientes"]/table/tbody/tr/td/input/../..'
            fields = [
                ['.//td[1]/input', 'get_attribute', 'value'],
                ['.//td[2]', 'text'],
                ['.//td[3]', 'text'],
                ['.//td[4]', 'text'],
                ['.//td[5]/span', 'text'],
                ['.//td[6]', 'text']
            ]
            column_names = ['Id', 'Nome', 'Permissão', 'Email', 'Edita', 'Apaga']
            df = self.get_data(elms_xpath, fields, column_names, dbg)
            
            self.save_to_db(df, 'Usuarios_Escritorio')
            return df

        
        load_page(self)
        get_user_data(self, dbg)
        pass
    
    def get_saved_clients(self): # ok
        comando = "Select * from Clientes order by v_nome"
        df = pd.read_sql(comando, con=self.cnx)

        return df

    def is_mov_client_done(self, id, anos, meses):
        con = sqlite3.connect('scraped.db')
        cursor = con.cursor()
        cursor.execute(f"SELECT compet FROM Andamento where id = '{id}' and status = 'finalizado';")
        try:
            x = cursor.fetchall()
            dones = [y[0] for y in x]
        except:
            dones = []
        
        compets = []
        for ano in anos:
            compets += [f'{str(ano)}.{"{:02d}".format(mes)}' for mes in meses]

        result = True
        for compet in compets:
            if compet not in dones:
                result = False
                break

        return result

    def get_mov(self, df_clientes, anos, meses, dbg=None):
        if dbg == None:
            dbg = self.dbg

        def load_page(self):
            self.browser.get(self.main_url + "/adm/impostos.asp?" + self.session + "&q=2")

        def get_this_page_tribute(self, id, name, compet, dbg):

            def get_value_from_element(elm_parent, field):
                        
                try:
                    elm = elm_parent.find_element(By.XPATH, field[0])
                    if field[1] == 'get_attribute':
                        value = elm.get_attribute(field[2])
                    elif field[1] == 'text':
                        value = elm.text
                    elif field[1] == 'option':
                        value = Select(elm).first_selected_option.text
                except Exception as e:
                    if dbg:
                        print(f'Error field: {field[0]}; msg: {str(e)[:100]}')
                    value = np.nan
                return value

            def get_info_on_tribute(elm):
                # detect type
                fn_call = elm.find_element(By.XPATH, './/input[@class="button_tb"][@value="Visualizar"]/../a').get_attribute('href')

                # [code, value, date, Obs]
                if 'Ver_Prot_GPS' in fn_call:
                    tipo = 'GPS Avulsa'
                    pos = [['.//td[2]', 'text', ''],
                        ['.//td[3]', 'text', ''],
                        ['.//td[4]', 'text', ''],
                        ['.//td[6]', 'text', '']]
                elif 'Ver_Prot_Darf' in fn_call:
                    tipo = 'DARF'
                    try:
                        first = elm.find_element(By.XPATH, './/td[1]/input').get_attribute('value')
                    except:
                        first = 'A'
                    if first.strip().isdigit():
                        pos = [['.//td[1]/input', 'get_attribute', 'value'],
                                ['.//td[2]', 'text', ''],
                                ['.//td[3]', 'text', ''],
                                ['.//td[5]', 'text', '']]
                    else:
                        pos = [['.//td[4]/select', 'option', ''],
                                ['.//td[2]/input', 'get_attribute', 'value'],
                                ['.//td[3]', 'text', ''],
                                ['.//td[6]/input', 'get_attribute', 'value']]
                elif 'Ver_Prot_ICMS' in fn_call:
                    tipo = 'Estatuais'
                    pos = [['.//td[1]/input', 'get_attribute', 'value'],
                        ['.//td[2]/input', 'get_attribute', 'value'],
                        ['.//td[3]', 'text', ''],
                        ['.//td[5]/input', 'get_attribute', 'value']]
                elif 'Ver_Outros_Pgtos' in fn_call:
                    tipo = 'Outros'
                    pos = [['.//td[1]/input', 'get_attribute', 'value'],
                        ['.//td[2]', 'text', 'value'],
                        ['.//td[3]', 'text', ''],
                        ['.//td[5]', 'text', '']]
                else:
                    raise
                protocol = elm.find_element(By.XPATH, './/td/input[@class="button_tb"][@value="Visualizar"]').get_attribute('onclick')
                info = [get_value_from_element(elm, x) for x in pos]
                valores = {'type':tipo, 'code': info[0], 'value': info[1],
                           'date': info[2], 'obs': info[3], 'fn_tribute': fn_call, 'fn_protocol':protocol}
                
                return valores

            def save_pdf(self, protocol_name, command):
                #Open Protocol
                self.browser.execute_script(command)
                # chenge to new page
                for window_handle in self.browser.window_handles:
                    if window_handle != self.main_window:
                        current_window = window_handle
                        self.browser.switch_to.window(window_handle)
                        break

                #rename and print
                self.browser.execute_script("document.title = \'{}\'".format(protocol_name + '.pdf'))
                self.browser.execute_script("window.print();")

                #close window
                self.browser.switch_to.window(current_window)
                self.browser.close()

                #focus on main again
                self.browser.switch_to.window(self.main_window)

            #get tributes elements
            impostos_elem = self.browser.find_elements(By.XPATH, '//form[@id="frmClientes"]//input[@class="button_tb"][@value="Visualizar"]/../..')

            impostos = []
            caracteres_especiais = r'[\\/*?:"<>|]'

            for i, elm in enumerate(impostos_elem):
                #obter informações do imposto:
                valores = get_info_on_tribute(elm)
                valores['seq'] = i
                valores['emp'] = id
                valores['compet'] = compet                
                impostos.append(valores)

                #Base nome
                vtype = valores['type']
                code = valores['code']
                obs = valores['obs']
                self.reprint(f'{name} - {compet} - ({i}) Guia - {vtype} {code} - {obs}')

                #Salvar protocolo obs= 47 length
                protocol_name = f'{name} - {compet} - ({i}) Protocolo - {vtype} {code} - {obs}'
                protocol_name = re.sub(caracteres_especiais,"-",protocol_name).replace("'","-")[:125]
                save_pdf(self, protocol_name, valores['fn_protocol'])

                #Salvar guia
                tribute_name = f'{name} - {compet} - ({i}) Guia - {vtype} {code} - {obs}'
                tribute_name = re.sub(caracteres_especiais,"-",tribute_name).replace("'","-")[:125]
                save_pdf(self, tribute_name, valores['fn_tribute'])
            
            return impostos

        def go_to_date(self, year, month):
            current_year = datetime.now().year +1
            inder_y = str(current_year-year)
            inder_m = str(month-1)
            command = f'document.getElementById("cboMes").selectedIndex = {inder_m};'
            command += f'document.getElementById("cboAno").selectedIndex = {inder_y};'
            command += 'return txtAnoMes_onchange()'
            self.browser.execute_script(command)

            pass

        def get_mov_of_client(self, cliente, anos, meses, dbg):
            id = cliente['id']
            name = cliente['v_nome']
            self.browser.find_element(By.XPATH, f'//*[@id="cboEmpresasT"]/option[@value={id}]').click()
            for ano in anos:
                for mes in meses:
                    compet = f"{str(ano)}.{'{:02d}'.format(mes)}"
                    self.reprint(f'Iniciando cliente {name} na competência {compet}')
                    status_atual = self.get_andamento_status(id, compet)
                    if status_atual != 'finalizado':
                        go_to_date(self, ano, mes)
                        self.update_andamento(id, compet, 'iniciado')
                        impostos = get_this_page_tribute(self, id, name, compet, dbg)
                        self.update_andamento(id, compet, 'salvando')
                        df_imp = pd.DataFrame(impostos)
                        df_imp.to_sql('Impostos', con=self.cnx, if_exists='append', index=False)
                        self.update_andamento(id, compet, 'finalizado')

            pass
                

        #RUN
        load_page(self)
        for index, cliente in df_clientes.iterrows():
            id = cliente['id']
            done = self.is_mov_client_done(id, anos, meses)
            name = cliente['v_nome']
            self.reprint(f'Inciciando cliente {name}')
            if done:
                if dbg:
                    print(f'Client {id} all ready done.')
            else:
                get_mov_of_client(self, cliente, anos, meses, dbg)

        pass

    def re_enable_clients(self, skip_to=None, dbg=None): # ok
        
        if dbg == None:
            dbg = self.dbg
        def load_page(self):
            self.browser.get(self.main_url + "/adm/clientes.asp?" + self.session + "&q=2&a=consultar")
        
        def re_enable_client(self, id):
            #goto client
            self.browser.execute_script(f'Clica_Empresa({[id]})')
            
            #enable
            command = f'document.getElementById("txt_fAtiva").selectedIndex = {0};'
            
            #change key
            self.browser.execute_script(command)
            with open('config.json') as file:
                random_new_key = json.load(file)['random_new_key']
            element = self.browser.find_element(By.ID, 'txtSenha')
            self.browser.execute_script("arguments[0].setAttribute('value',arguments[1])",element, random_new_key)
            #save
            self.browser.find_element(By.ID, 'cmdGrava').click()

            pass
        
        load_page(self)
        df = self.get_saved_clients()
        df = df[df['ativo'] == 'Não']
        for id in list(df['id']):
            if str(id) == str(skip_to):
                skip_to = None
                if dbg:
                    print(f'Todos os clientes antes até o {str(id)} foram pulados.')
            if skip_to is None:
                re_enable_client(self, id)
                if dbg:
                    print(f'Cliente {id} reativado e senha alterada.')
        
        pass

    def test(self, t=1):
        sleep(t)
        return f'This calls is runing in the window {self.main_window}, sllep {t} sec.'

    def matar(self):
        self.browser.quit()
        pass

def insistir(quant_to_split, number_bot, anos, meses):
    es = eContabilSite()
    df_clientes = es.get_saved_clients()

    #split load
    n = quant_to_split
    dfs = []
    for i in range(0,n):
        dfs.append(df_clientes[df_clientes['index'] % n == i % n])

    again = True
    while again:
        start = datetime.now()
        try:
            es = eContabilSite()
            es.get_mov(dfs[number_bot], anos, meses)
            again = False
            es.matar()
            dur = datetime.now() - start
            print('')
            print(f'\033[92m Finish Dutration {dur.total_seconds()}\033[0m')
        except:
            dur = datetime.now() - start
            es.matar()
            print('')
            print(f'\033[93m Fail Dutration {dur.total_seconds()}\033[0m')
            pass
    
    return es

def main(console_log=False):

    if console_log: print('login')
    es = eContabilSite(console_log)

    if console_log: print('create_tables')
    es.create_tables()

    if console_log: print('get_clients')
    #es.get_clients()

    if console_log: print('get_movfolha')
    #es.get_movfolha()
    if console_log: print('get_outros_pag')
    #es.get_outros_pag()

    input("Press Enter to end this.")
    es.browser.quit()

if __name__ == '__main__':
    main()
