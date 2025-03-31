import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import os
from store_mapper import StoreMapper
from open_driver import OpenDriver
from helper_methods import HelperMethods
from selenium.webdriver.common.keys import Keys
import logging

# Configuração do log
logging.basicConfig(
    filename="log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

class InternalMovimentProcess:
    def __init__(self, caminho, type):
        self.type = type
        self.caminho = caminho
        self.driver = None  # Driver será atribuído dinamicamente


    def path_control(self):
        dir_name, file_name = os.path.split(self.caminho)
        base_name, ext = os.path.splitext(file_name)      
        control_file_name = f"{base_name}_CONTROLE{ext}"
        control_file_path = os.path.join(dir_name, control_file_name)
        return control_file_path


    def open_excel(self):
        store_mapper = StoreMapper()
        df = pd.read_excel(self.caminho, engine='openpyxl')
        df["Loja Origem"] = df["Loja Origem"].astype(str).map(store_mapper.dict_num_lojas).fillna(df["Loja Origem"])
        df["Loja Destino"] = df["Loja Destino"].astype(str).map(store_mapper.dict_num_lojas).fillna(df["Loja Destino"])
        df.to_excel(self.path_control(), index=False)
        return df
    
    
    def df_by_group(self):
        try:
            df = pd.read_excel(self.path_control())
        except:
            df = self.open_excel()
        df["Status"] = df["Status"].astype(str)
        df_filtrado = df[pd.isna(df["Status"]) | 
                        (df["Status"] == "nan") | 
                        (~df["Status"].str.contains("-Listado", na=False))]
        if df_filtrado.empty:
            HelperMethods.notificar("Concluido", "Sem Pedidos Novos")
            return None
        grouped_dfs = df_filtrado.groupby(["Loja Origem", "Loja Destino", "Cond. Pagamento", "Operacao"])
        return grouped_dfs
    

    def update_status(self, row, info):
        df = pd.read_excel(self.path_control())
        mask = (
            (df["Loja Origem"] == row["Loja Origem"]) &
            (df["Loja Destino"] == row["Loja Destino"]) &
            (df["Quantidade"] == row["Quantidade"]) &
            (df["Codigo"] == row["Codigo"]) &
            (df["Cond. Pagamento"] == row["Cond. Pagamento"]) &
            (df["Operacao"] == row["Operacao"]))
        
        df["Status"] = df["Status"].astype(str).replace('nan', '').fillna('')

        if mask.any():
            df.loc[mask, 'Status'] = df.loc[mask, 'Status'] + f"{info}"
            df.to_excel(self.path_control(), index=False)
            return True
        else:
            HelperMethods.notificar("STATUS ERRO", f"'{info}' Nao atualizada.")
            return False



    def select_mov_int(self):
        self.driver.find_element(By.XPATH, "//li[@onclick=\"simulaClick('op57')\"]").click()
        HelperMethods.carregando(self.driver)


    def select_vendas_op(self, op):
        self.driver.execute_script("document.getElementById('op1').click();")
        self.driver.execute_script(f"document.getElementById('op{op}').click();")
        HelperMethods.carregando(self.driver)


    def select_type(self):
        select_field = WebDriverWait(self.driver, 5).until(
            EC.presence_of_element_located((By.ID, 'incCentral:formCrudBase:selPadraoLancamento'))
        )
        select = Select(select_field)
        select.select_by_visible_text(self.type)
        iniciar = self.driver.find_element(By.ID, 'incCentral:formCrudBase:btnIniciar')
        iniciar.click()
        HelperMethods.carregando(self.driver)


    def select_cnpj_origem(self, cnpj_origem):
        cnpj_origem = str(cnpj_origem)
        select_cnpj = self.driver.find_element(By.XPATH, f".//option[contains(@value, '{cnpj_origem}')]")
        select_cnpj.click()
        HelperMethods.carregando(self.driver)


    def import_item(self, lista):
        HelperMethods.carregando(self.driver)
        for row in lista:
            logging.info(row)
            quantite_field = self.driver.find_element(By.ID, 'incCentral:formCrudBase:txtProduto')
            quantite_field.click()
            quantite_field.clear()
            quantite_field.send_keys(row)
            self.driver.find_element(By.ID, 'incCentral:formCrudBase:btnAdicionar').click()
            HelperMethods.carregando(self.driver)


    def click_cnpj_field(self, cnpj_destino):
        HelperMethods.carregando(self.driver)
        rows = self.driver.find_elements(By.XPATH, "//tbody[@id='tblPsqInfoPtcControladoBody']/tr")
        for row in rows:
            cnpj_text = row.find_element(By.XPATH, ".//td[contains(@class, 'p-2')]").text
            if cnpj_destino in cnpj_text:
                icon_element = row.find_element(By.XPATH, ".//i[contains(@class, 'material-icons')]")
                icon_element.click()
                HelperMethods.carregando(self.driver)
                return


    def select_cnpj_destino(self, cnpj_destino):
        self.driver.find_element(By.ID, 'incCentral:formCrudBase:btnPesqProduto').click()
        HelperMethods.carregando(self.driver)
        cnpj = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'incCentral:formPnlModalPesquisaParticipante:txtPtcCoCnpjFiltro'))
        )
        cnpj.click()
        cnpj.send_keys(cnpj_destino)
        self.driver.find_element(By.ID, 'incCentral:formPnlModalPesquisaParticipante:btnPsqPtcControlado').click()
        HelperMethods.carregando(self.driver)
        self.click_cnpj_field(cnpj_destino)
        HelperMethods.carregando(self.driver)


    def select_payment_condition(self, value):
        try:
            select_element = self.driver.find_element(By.ID, "incCentral:formCrudBase:selCondicaoPagto")
            select = Select(select_element)
            select.select_by_value(value)
            HelperMethods.carregando(self.driver)
        except:
            pass


    def finalizar_processo(self):
        self.driver.find_element(By.ID, 'incCentral:formCrudBase:btnFinalizar').click()
        HelperMethods.carregando(self.driver)


    def colect_pd_number(self):
        HelperMethods.carregando(self.driver)
        element = self.driver.find_element(By.XPATH, "//li[contains(@class, 'alert-success')]")
        number = element.text.split(' ')[2]
        return number

    def button_pesquisar_pd(self):
        def found_tr():
            tentativa = 0
            while tentativa <= 5:
                rows = self.driver.find_elements(By.XPATH, "//tbody[@id='tblPrdBodyPesquisa']/tr")
                if len(rows) > 0:
                    return True
                tentativa =+ 1
                time.sleep(1)
            return False
        try:
            parent_element = self.driver.find_element(By.ID, "incCentral:formCrudBase:btnPedPesquisar")
        except:
            parent_element = self.driver.find_element(By.ID, 'incCentral:formCrudBase:btnPdfPesquisar')
        material_icon = parent_element.find_element(By.XPATH, ".//i[contains(@class, 'material-icons')]")
        material_icon.click()
        HelperMethods.carregando(self.driver)
        
        if found_tr() is True:
            return
        material_icon.click()
        HelperMethods.carregando(self.driver)
        found_tr()

    def select_pd(self, pd_number):
        rows = self.driver.find_elements(By.XPATH, "//tbody[@id='tblPrdBodyPesquisa']/tr")
        for index, row in enumerate(rows, start=1):
            codigo_element = row.find_element(By.XPATH, ".//span[contains(@class, 'th-responsive') and contains(text(), 'Código')]")
            codigo_td = codigo_element.find_element(By.XPATH, "./ancestor::td")
            codigo_value = codigo_td.text.split()[-1]
            if codigo_value == pd_number:
                try:
                    element = row.find_element(By.XPATH, ".//span[contains(@class, 'th-responsive') and contains(text(), 'OS')]")
                except:
                    element = row.find_element(By.XPATH, ".//span[contains(@class, 'th-responsive') and contains(text(), 'Alterar')]")
                alterar_td = element.find_element(By.XPATH, "./ancestor::td")
                alterar_link = alterar_td.find_element(By.TAG_NAME, "a")
                alterar_link.click()
                HelperMethods.carregando(self.driver)
                return

    def liberar_faturamento(self):
        faturamento = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'incCentral:formCrudBase:formCapa:btnPedLiberar'))
        )
        self.driver.execute_script("arguments[0].scrollIntoView(true);", faturamento)
        faturamento.click()
        HelperMethods.carregando(self.driver)


    def accept_confirm(self):
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div#confirmaLiberacaoFaturamento.modal.fade.show")))
        id_liberacao = self.driver.find_element(By.ID, 'incCentral:confirmaLiberacaoFaturamentoformFooter')
        button_confirm = id_liberacao.find_element(By.CSS_SELECTOR, "a.btn.btn-primary")
        button_confirm.click()
        HelperMethods.carregando(self.driver)


    def select_operation(self, operation_number):
        select_element = self.driver.find_element(By.ID, "incCentral:frmmodalOperacao:selOpeCoOperacao")
        select = Select(select_element)
        select.select_by_value(operation_number)
        self.driver.find_element(By.XPATH, "//*[contains(@id, 'incCentral:frmmodalOperacao:')]").click()
        select_element.send_keys(Keys.TAB)
        button = self.driver.find_element(By.XPATH, "//a[text()='Selecionar Operação']")
        button.click()
        HelperMethods.carregando(self.driver)
    
    def listar_pedido(self):
        self.driver.find_element(By.XPATH, "//*[contains(@id, 'incCentral:confirmaListarPedidoformFooter:')]").click()

    def _liberar_e_listar(self, origem, operation, pd_number, group_df):
        logging.info(f"Vendas>Pedido")
        self.select_vendas_op(6)
        self.select_cnpj_origem(origem)
        self.button_pesquisar_pd()
        self.select_pd(pd_number)
        self.liberar_faturamento()
        self.accept_confirm()
        logging.info(f"STATUS PD.{pd_number}-Liberado")
        
        for _, row in group_df.iterrows():
            if not self.update_status(row, "-Liberado"):
                logging.error(f"Erro ao atualizar status durante o processo de liberação.")
                raise Exception("Erro no update_status. Processo interrompido.")
        self._listar(origem, operation, pd_number, group_df)

    def _listar(self, origem, operation, pd_number, group_df):
        logging.info(f"Vendas>Faturamento")
        self.select_vendas_op(8)
        self.select_cnpj_origem(origem)
        self.button_pesquisar_pd()
        self.select_pd(pd_number)
        self.select_operation(operation)
        logging.info(f"STATUS PD.{pd_number}-Liberado-Listado")
        
        for _, row in group_df.iterrows():
            if not self.update_status(row, "-Listado"):
                logging.error(f"Erro ao atualizar status durante o processo de listagem.")
                raise Exception("Erro no update_status. Processo interrompido.")
        logging.info(f"Finalizado com sucesso.\n\n")

    def processo_inclusao_pedidos(self):
        grouped_dfs = self.df_by_group()
        if grouped_dfs is None:
            logging.info(f"Concluído: Sem Pedidos Novos\n\n")
            return
        
        open_driver = OpenDriver()
        self.driver = open_driver.open_driver("vendas")

        for group_name, group_df in grouped_dfs:
            origem, destino, payment, operation = map(str, group_name)
            group_df['Qtd&Code'] = group_df['Quantidade'].astype(str) + "&" + group_df['Codigo'].astype(str)
            lista = group_df['Qtd&Code'].tolist()
            status = group_df.iloc[0]['Status']
            
            if pd.isna(status) or status == 'nan':
                logging.info(f"STATUS NaN")
                logging.info(f"Movimentacao Interna")
                logging.info(f"ORIGEM {StoreMapper().get_loja_by_cnpj(origem)}")
                logging.info(f"DESTINO {StoreMapper().get_loja_by_cnpj(destino)}")
                self.select_mov_int()
                self.select_type()
                self.select_cnpj_origem(origem)
                self.import_item(lista)
                self.select_cnpj_destino(destino)
                self.select_payment_condition(payment)
                self.finalizar_processo()
                pd_number = self.colect_pd_number()
                logging.info(f"STATUS PD.{pd_number}")

                for _, group_row in group_df.iterrows():    
                    if not self.update_status(group_row, f"PD.{pd_number}"):
                        logging.error(f"Erro ao atualizar status inicial.")
                        raise Exception("Erro no update_status. Processo interrompido.")
                self._liberar_e_listar(origem, operation, pd_number, group_df)

            elif "PD." in status and "-Liberado" not in status:
                logging.info(f"Liberar Pedido")
                pd_number = status.split('.')[1]
                self._liberar_e_listar(origem, operation, pd_number, group_df)

            elif "-Liberado" in status and "-Listado" not in status:
                logging.info(f"Listar Pedido")
                pd_number = status.split('.')[1].split('-')[0]
                self._listar(origem, operation, pd_number, group_df)

        HelperMethods.notificar("Concluído", "Pedidos Finalizados")






if __name__ == "__main__":
    caminho = "C:/Mac/Home/Downloads/Layout_Mov_Int.xlsx"
    processor = InternalMovimentProcess(caminho, "[TRANSFERENCIA]")
    processor.processo_inclusao_pedidos()
    