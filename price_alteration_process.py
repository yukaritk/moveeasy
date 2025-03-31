from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
from selenium.webdriver.common.keys import Keys
import os
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from open_driver import OpenDriver
from helper_methods import HelperMethods
from store_mapper import StoreMapper
import logging


# Configuração do log
logging.basicConfig(
    filename="log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)


class PriceAlterationProcessor:
    def __init__(self, caminho):
        self.caminho = caminho
        self.driver = None  # Driver será atribuído dinamicamente

    def set_driver(self, driver):
        self.driver = driver

    def novo_nome_csv(self):
        dir_name, file_name = os.path.split(self.caminho)
        base_name, ext = os.path.splitext(file_name)
        csv_file_name = f"{base_name}_alteracao_preco_parcial.csv"
        return os.path.join(dir_name, csv_file_name)

    def xml_csv(self):
        df = pd.read_excel(self.caminho, engine='openpyxl')
        df['Tipo do Codigo'] = df['Tipo do Codigo'].str.lower()

        try:
            df['Vl. Custo'] = df['Vl. Custo'].astype(str).str.replace(',', '.').astype(float)
        except:
            pass
        try:
            df['Vl. Revenda'] = df['Vl. Revenda'].astype(str).str.replace(',', '.').astype(float)
        except:
            pass
        try:
            df['Data inicio'] = pd.to_datetime(df['Data inicio'], errors='coerce')
            df['Data inicio'] = df['Data inicio'].dt.strftime('%d/%m/%Y')
        except:
            pass

        df.to_csv(self.novo_nome_csv(), sep=";", index=False)
        return df

    def arquivo_final(self):
        old_file_path = self.novo_nome_csv()
        dir_name, file_name = os.path.split(old_file_path)
        new_file_name = file_name.replace('parcial', 'final')
        new_file_path = os.path.join(dir_name, new_file_name)
        os.rename(old_file_path, new_file_path)

    def select_price_alt(self):
        field_manutencao = self.driver.find_element(By.ID, "op58")
        self.driver.execute_script("arguments[0].click();", field_manutencao)

    def fecha_calendario(self):
        div_id = "ui-datepicker-div"
        while True:
            div_element = self.driver.find_element(By.ID, div_id)
            estilo = div_element.get_attribute("style")
            if "display: none" in estilo:
                break
            time.sleep(0.5)

    def inclusao_data_inicio(self, data):
        HelperMethods.carregando(self.driver)
        time.sleep(1)
        field_data = self.driver.find_element(By.ID, "incCentral:formCrudBase:txtPvdDtIniValidade")
        field_data.click()
        field_data.send_keys(Keys.CONTROL, 'a')
        field_data.send_keys(Keys.DELETE)
        field_data.send_keys(data)
        botao_fechar = self.driver.find_element(By.CLASS_NAME, "ui-datepicker-close")
        botao_fechar.click()
        self.fecha_calendario()

    def seleciona_loja(self, loja):
        HelperMethods.carregando(self.driver)
        element_loja = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "incCentral:formCrudBase:selEmiCoCnpj"))
        )
        select_loja = Select(element_loja)
        sm = StoreMapper()
        
        select_loja.select_by_visible_text(sm.get_grupo_by_num(loja))

        start_time = time.time()

        # XPath específico para o <option> dentro do <select>
        xpath_option = f"//select[@id='incCentral:formCrudBase:selEmiCoCnpj']/option[contains(text(), '{sm.get_grupo_by_num(loja)}')]"

        # Loop para verificar até que o <option> esteja selecionado
        while True:
            try:
                # Localizar o elemento <option> correspondente dentro do <select>
                option_element = self.driver.find_element(By.XPATH, xpath_option)
                # Verificar se o atributo 'selected' está presente
                if option_element.get_attribute("selected") == "true":
                    return True
                else:
                    time.sleep(0.5)
            
            except Exception as e:
                # Exceção caso o elemento não seja encontrado
                return False

            if time.time() - start_time > 5:
                return False
            

    def abre_pesquisa_grupo(self):
        modal_id = "pnlPsqGrupoPreco"
        while True:
            modal = self.driver.find_element(By.ID, modal_id)
            aria_hidden = modal.get_attribute("aria-hidden")
            if aria_hidden != "true":
                break
            time.sleep(0.5)

    def abre_pesquisa_produto(self):
        modal_id = "pnlPsqProduto"
        while True:
            modal = self.driver.find_element(By.ID, modal_id)
            aria_hidden = modal.get_attribute("aria-hidden")
            if aria_hidden != "true":
                break
            time.sleep(0.5)

    def captura_codigo(self, tipo, code):
        HelperMethods.carregando(self.driver)
        tabela_id = "tblPsqInfoGrupoPrecoBody" if tipo == "grupo" else "tblPsqInfoParticipanteBody"
        elementos_nome = self.driver.find_elements(By.XPATH, f"//tbody[@id='{tabela_id}']//span[contains(@class, 'th-responsive') and text()='Nome']/parent::td")

        if elementos_nome:
            for elemento in elementos_nome:
                texto_completo = elemento.text
                match = re.match(r"^\d+", texto_completo)
                if match:
                    codigo = match.group(0)
                    if code == codigo:
                        proximo_elemento = elemento.find_element(By.XPATH, "./following-sibling::td[@class='p-2']")
                        link = proximo_elemento.find_element(By.TAG_NAME, "a")
                        id_link = link.get_attribute("id")
                        return code, id_link
        return None, None

    def pertence_grupo(self):
        try:
            # Localizar todos os elementos com a classe "form-group"
            elementos = self.driver.find_elements(By.CLASS_NAME, "form-group")
            
            # Iterar sobre os elementos para encontrar o texto desejado
            for elemento in elementos:
                texto = elemento.text.strip()  # Remover espaços extras no texto
                # Verificar se o texto começa com a frase específica
                if texto.startswith("Produto pertence a um grupo de preços."):
                    # Procurar o número entre parênteses usando regex
                    match = re.search(r"\((\d+)\)", texto)
                    if match:
                        return match.group(1)  # Retorna o número encontrado
            return None
        except:
            return None

    def fechar_campo_pesquisa(self):
        while True:  # Loop infinito até encontrar e interagir com o botão
            try:
                # Buscar todos os botões novamente
                botoes = self.driver.find_elements(By.TAG_NAME, "button")
                
                for botao in botoes:
                    # Verificar se o botão possui a classe "close"
                    classes = botao.get_attribute("class")
                    if "close" in classes:
                        # Verificar se o botão está visível e habilitado
                        if botao.is_displayed() and botao.is_enabled():
                            try:
                                # Clicar no botão
                                action = ActionChains(self.driver)
                                action.move_to_element(botao).click().perform()
                                return True
                            except:
                                continue
            except:
                continue

    def selecionar_produto(self, tipo, code):
        HelperMethods.carregando(self.driver)
        code = str(code)

        time.sleep(1)
        # Clica no botão para selecionar o produto
        button_produto = self.driver.find_element(By.ID, 'incCentral:formCrudBase:btnPesqProduto')
        button_produto.click()

        self.abre_pesquisa_produto()

        value, id_seleciona = self.captura_codigo(tipo, code)

        if value == code:
            if id_seleciona is None:
                self.fechar_campo_pesquisa()
                return "ERRO"
            else:
                self.driver.find_element(By.ID, id_seleciona).click()
                return 'OK'

        field_code = WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.ID, "incCentral:formPnlModalPesquisaPrd:txtPrdCoProdutoFiltro")))
        field_code.click()
        field_code.clear()
        field_code.send_keys(code)

        button_pesquisar = WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.ID, "incCentral:formPnlModalPesquisaPrd:btnPsqProduto")))
        button_pesquisar.click()

        HelperMethods.carregando(self.driver)

        value, id_seleciona = self.captura_codigo(tipo, code)

        if id_seleciona is None:
            self.fechar_campo_pesquisa()
            return "ERRO"
        else:
            if value == code:
                self.driver.find_element(By.ID, id_seleciona).click()
                HelperMethods.carregando(self.driver)
                span = self.pertence_grupo()
                if span != None:
                    return f"ERRO - Grupo de preco {span}"
                return 'OK'

    def selecionar_grupo(self, tipo, code):
        HelperMethods.carregando(self.driver)
        code = str(code)

        time.sleep(1)

        # Selecionar "Grupo"
        btn_grupo = self.driver.find_element(By.ID, "incCentral:formCrudBase:btnPesqGrupoPreco")
        btn_grupo.click()

        self.abre_pesquisa_grupo()

        value, id_seleciona = self.captura_codigo(tipo, code)

        if value == code:
            if id_seleciona is None:
                self.fechar_campo_pesquisa()
                return "ERRO"
            else:
                self.driver.find_element(By.ID, id_seleciona).click()
                return 'OK'

        field_code = WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.ID, "incCentral:formPnlModalPesquisaGpr:txtGrpCoGrupoFiltro")))
        field_code.click()
        field_code.clear()
        field_code.send_keys(code)

        button_pesquisar = WebDriverWait(self.driver,10).until(EC.element_to_be_clickable((By.ID, "incCentral:formPnlModalPesquisaGpr:btnPsqGrupoPreco")))
        button_pesquisar.click()

        HelperMethods.carregando(self.driver)

        value, id_seleciona = self.captura_codigo(tipo, code)

        if id_seleciona is None:
            self.fechar_campo_pesquisa()
            return "ERRO"
        else:
            if value == code:
                self.driver.find_element(By.ID, id_seleciona).click()
                return 'OK'
            else:
                return 'ERRO'

    def inclusao_preco(self, tipo, code, vlcusto, vlrevenda):
        # ID do elemento dinâmico
        elemento_id = "incCentral:formCrudBase:pnlNoPrdFiltro"

        # Loop para verificar o atributo 'xmlns' dinamicamente
        while True:
            try:
                # Localizar o elemento
                elemento = self.driver.find_element(By.ID, elemento_id)
                
                # Obter o valor do atributo 'xmlns'
                xmlns = elemento.get_attribute("xmlns")
                
                if xmlns is not None:
                    break
                else:
                    time.sleep(0.5)  # Aguarde 1 segundo antes de tentar novamente
            
            except Exception as e:
                time.sleep(0.5)  # Aguardar antes de tentar novamente

        # Localizar o elemento com ID que começa com "incCentral:formCrudBase:pnlNoPrd" e que tem o atributo 'value'
        elemento = self.driver.find_element(By.XPATH, "//input[starts-with(@id, 'incCentral:formCrudBase:pnlNoPrd') and @value]")
        # Obter o valor do atributo 'value'
        value = elemento.get_attribute("value")
        # Capturar apenas o número entre colchetes usando regex
        match = re.match(r"\[(\d+)\]", value)

        if match:
            numero = match.group(1)  # Valor do número extraído
            # Comparar com o code informado
            if str(code) == numero:
                # print("Código corresponde, continuando...")
                pass
            else:
                return False
        else:
            return False

        if tipo == "grupo":
            # Localizar o elemento de preço de reposição e alterar o valor
            custo_reposicao = self.driver.find_element(By.ID, "incCentral:formCrudBase:txtGpvVlCustoReposicao")
            custo_reposicao.clear()  # Limpar o campo antes de inserir o novo valor
            custo_reposicao.send_keys(str(vlcusto))

            # Localizar o elemento de preço de revenda e alterar o valor
            venda_revenda = self.driver.find_element(By.ID, "incCentral:formCrudBase:txtGpvVlVendaRevenda")
            venda_revenda.clear()  # Limpar o campo antes de inserir o novo valor
            venda_revenda.send_keys(str(vlrevenda))

        else:
            # Localizar o elemento de preço de reposição e alterar o valor
            custo_reposicao = self.driver.find_element(By.ID, "incCentral:formCrudBase:txtPvdVlCustoReposicao")
            custo_reposicao.clear()  # Limpar o campo antes de inserir o novo valor
            custo_reposicao.send_keys(str(vlcusto))

            # Localizar o elemento de preço de revenda e alterar o valor
            venda_revenda = self.driver.find_element(By.ID, "incCentral:formCrudBase:txtPvdVlVendaRevenda")
            venda_revenda.clear()  # Limpar o campo antes de inserir o novo valor
            venda_revenda.send_keys(str(vlrevenda))
        
        # Localizar o botão "Salvar" pelo ID
        botao_salvar = self.driver.find_element(By.ID, "incCentral:formCrudBase:btngpvPrecoAddEdt")

        # Clicar no botão "Salvar"
        botao_salvar.click()

        # Loop para verificar a frase "Salvo com sucesso!"
        start_time = time.time()
        while True:
            try:
                # Localizar o elemento pelo texto
                salvo_sucesso = self.driver.find_element(By.XPATH, "//li[contains(@class, 'alert alert-success') and contains(text(), 'Salvo com sucesso!')]")
                
                # Se o elemento for encontrado, exibir a mensagem e sair do loop
                if salvo_sucesso:
                    return True
            
            except Exception:
                # Exceção esperada caso o elemento ainda não esteja presente
                time.sleep(0.5)  # Aguardar antes de verificar novamente

            if time.time() - start_time > 5:
                return False
            
    def analisar_linha(self, df):
        open_driver = OpenDriver()
        self.driver = open_driver.open_driver("vendas")  # O driver é iniciado aqui
        self.select_price_alt()
        logging.info(f"Manutencao de Preco")

        for idx, row in df.iterrows():
            # Verificar o status da linha
            data = row['Data inicio']
            status = str(row['Status'])
            codigo = row['Produto/Grupo']
            lojas_total = row['Loja/Grupo']
            lojas = str(lojas_total).split(',')
            tipo = row['Tipo do Codigo']
            vl_custo = row['Vl. Custo']
            vl_revenda = row['Vl. Revenda']

            if status.startswith("OK") or status.startswith("ERRO"):
                # Se o status for OK ou ERRO, pular para a próxima linha
                continue

            elif status.startswith("PARCIAL"):
                # Se o status for PARCIAL, obter as lojas já analisadas
                lojas_analisadas = status.split('-')[-1].split(',')
                # Encontrar a próxima loja/grupo a ser analisada
                lojas_pendentes = [loja for loja in lojas if loja not in lojas_analisadas]
                analisadas = status
                if len(lojas_pendentes) == 0:
                    df.at[idx, 'Status'] = "OK"
                    df.to_csv(self.novo_nome_csv(), sep=";" ,index=False)
                    continue
            else:
                # Se não for PARCIAL, todas as lojas/grupos estão pendentes
                lojas_pendentes = lojas
                analisadas = "PARCIAL-"

            self.inclusao_data_inicio(data)
            logging.info(f"Data Inicio {data}")

            start_time = time.time()
            for loja in lojas_pendentes:
                elapsed_time = time.time() - start_time
                if elapsed_time > 60:
                    logging.warning(f"CRASH")
                    self.driver.quit()
                    return self.analisar_linha(df) 
                bool_loja = self.seleciona_loja(loja)
                if bool_loja is False:
                    analisadas_parcial = analisadas.split("-")[0] + f" Loja {loja} nao localizada."
                    logging.info(f"Loja Nao Localizada: {loja}")
                    if analisadas.split("-")[-1]:
                        analisadas = analisadas_parcial + "-" + analisadas.split("-")[-1]
                    else:
                        analisadas = analisadas_parcial + "-"
                    df.at[idx, 'Status'] = analisadas
                    df.to_csv(self.novo_nome_csv(), sep=";" ,index=False)
                    continue
                else:
                    logging.info(f"Preco Alterado da Loja: {loja}")
                    if tipo == 'produto':
                        logging.info(f"Produto: {codigo}")
                        mensagem = self.selecionar_produto(tipo, codigo)
                    else:
                        logging.info(f"Grupo: {codigo}")
                        mensagem = self.selecionar_grupo(tipo, codigo)
                    # Atualizar o status com a mensagem retornada
                    if mensagem.startswith("ERRO"):
                        df.at[idx, 'Status'] = mensagem
                        df.to_csv(self.novo_nome_csv(), sep=";" ,index=False)
                        logging.info(f"Codigo Nao Localizado: {codigo}")
                        break
                    else:
                        logging.info(f"PRECO CUSTO: {vl_custo}")
                        logging.info(f"PRECO REVENDA: {vl_revenda}")
                        bool_status = self.inclusao_preco(tipo, codigo, vl_custo, vl_revenda)
                        if bool_status is False:
                            logging.info(f"ERRO NA INCLUSAO DE PRECO")
                            df.at[idx, 'Status'] = "ERRO NA INCLUSAO DO PRECO"
                            continue
                        else:
                            if analisadas.split("-")[-1]:
                                analisadas = analisadas + "," + loja
                            else:
                                analisadas = analisadas + loja
                            df.at[idx, 'Status'] = analisadas
                            df.to_csv(self.novo_nome_csv(), sep=";" ,index=False)
                            logging.info(f"ALTERACAO REALIZADA")
                            logging.info("")
                            logging.info("")

            new_status = analisadas.split("-")[-1]
            if lojas_total == new_status:
                df.at[idx, 'Status'] = "OK"
                df.to_csv(self.novo_nome_csv(), sep=";" ,index=False)

    def analisar_planilha(self):
        try:
            df = pd.read_csv(self.novo_nome_csv(), sep=";")
        except Exception:
            df = self.xml_csv()

        tentativas = 0
        while tentativas <= 10:
            try:
                # Verificar se há algum status "PARCIAL" ou em branco
                if df['Status'].isnull().any() or any(df['Status'].str.startswith("PARCIAL")):
                    # Rodar a função para continuar o processamento
                    self.analisar_linha(df)
                    # Após finalizar analisar_linha, rodar arquivo_final e notificar sucesso
                    self.arquivo_final()
                    HelperMethods.notificar("Concluído", "Alteração de preço concluída")
                    break
                else:
                    # Se todos os status forem "OK" ou "ERRO", rodar arquivo_final
                    self.arquivo_final()
                    HelperMethods.notificar("Concluído", "Alteração de preço concluída")
                    break
            except Exception as e:
                # Em caso de erro, notificar falha
                HelperMethods.notificar('Erro', 'Falha na Alteração de preço.')
                print(f"Erro durante o processamento: {e}")
            tentativas += 1
            

if __name__ == "__main__":
    caminho = "C:/Mac/Home/Desktop/Sumire/alteracao_preco_Belliz_jan25.xlsx"
    processor = PriceAlterationProcessor(caminho)
    processor.analisar_planilha()