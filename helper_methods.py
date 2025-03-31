import time
from plyer import notification


class HelperMethods:
    @staticmethod
    def notificar(titulo, mensagem):
        notification.notify(
            title=titulo,
            message=mensagem,
            app_name="Robozinho",
            timeout=6  # Tempo em segundos que a notificação ficará visível
        )

    @staticmethod
    def carregando(driver):
        status_start_id = "_viewRoot:status.start"

        while True:
            # Localizar o elemento no navegador
            status_start = driver.find_element("id", status_start_id)
            
            # Obter o valor do atributo 'style' e verificar o 'display'
            estilo = status_start.get_attribute("style")
            if "display: none" in estilo:
                break  # Sair do loop quando o display for none
            
            # Mensagem indicando que a página ainda está carregando
            time.sleep(0.5)  # Espera antes de verificar novamente