from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from open_driver import OpenDriver
from helper_methods import HelperMethods

class ImportType:
    def __init__(self):
        self.driver = None


    def select_mov_int(self):
        self.driver.find_element(By.XPATH, "//li[@onclick=\"simulaClick('op57')\"]").click()
        HelperMethods.carregando(self.driver)

    @staticmethod
    def update_list(item):
        with open("lista_padrao_lancamento.txt", "w") as file:
            file.write(str(item))


    def colect_padrao_lancamento(self):
        select_id = self.driver.find_element(By.ID, "incCentral:formCrudBase:selPadraoLancamento")
        select = Select(select_id)
        options_list = select.options
        new_list = []
        for option in options_list:
            new_list.append(option.text)
        self.update_list(new_list)


    def update_type(self):
        open_driver = OpenDriver()
        self.driver = open_driver.open_driver("vendas")
        self.select_mov_int()
        self.colect_padrao_lancamento()


if __name__ == "__main__":
    processor = ImportType()
    processor.update_type()