from tkinter import *
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import load_workbook
import subprocess
from internal_moviment_process import InternalMovimentProcess
from import_type import ImportType
import ast


class InternalMovementPage:
    def __init__(self, root):
        # Configuração da janela
        self.window = Toplevel(root)
        self.window.title("Movimentação Interna")
        self.window.configure(background='DeepPink2')
        self.window.geometry("500x200")
        self.window.resizable(True, True)
        self.window.minsize(width=500, height=200)
        
        frame = Frame(self.window, bg='white')
        frame.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

        Label(frame, bg='white', text="Arquivo", anchor='w').place(relx=0.03, rely=0.1, relwidth=0.3, relheight=0.1)
        self.folder_entry = Entry(frame)
        self.folder_entry.place(relx=0.3, rely=0.1, relwidth=0.5, relheight=0.1)

        Button(frame, text="...", bd=3, command=self.open_file_dialog).place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.1)

        Label(frame, bg='white', text="Tipo de Lancamento", anchor='w').place(relx=0.03, rely=0.3, relwidth=0.3, relheight=0.1)
        with open("lista_padrao_lancamento.txt", "r") as file:
            content = file.read().strip()  # Lê todo o conteúdo do arquivo como string
            tipos_lancamento = ast.literal_eval(content)  # Converte a string para lista
        self.list_type_var = StringVar(frame)
        self.list_type_var.set(tipos_lancamento[0])
        self.list_type = OptionMenu(frame, self.list_type_var, *tipos_lancamento)
        self.list_type.place(relx=0.3, rely=0.3, relwidth=0.5, relheight=0.11)

        Button(frame, text="Atualizar", bd=3, font=("", 7), command=self.update_lista_lancamentos).place(relx=0.8, rely=0.3, relwidth=0.1, relheight=0.1)

        Button(frame, text="Layout", bd=3, command=self.download_layout).place(
            relx=0.5, rely=0.75, relwidth=0.2, relheight=0.15)

        # Botão para iniciar a movimentação
        Button(frame, text="Iniciar", bd=3, command=self.start_internal_movement).place(
            relx=0.7, rely=0.75, relwidth=0.2, relheight=0.15)

    def open_file_dialog(self):
        file_path = filedialog.askopenfilename(parent=self.window, title="Selecione o arquivo")
        if file_path:
            self.folder_entry.delete(0, END)  # Limpa a entrada
            self.folder_entry.insert(0, file_path)  # Insere o caminho do arquivo selecionado


    def update_lista_lancamentos(self):
        loader = ImportType()
        loader.update_type()
        self.window.destroy()


    def download_layout(self):
        try:
            # Define o nome do arquivo e o caminho para a pasta de downloads
            download_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            file_path = os.path.join(download_folder, "Layout_Mov_Int.xlsx")

            # Define os nomes das colunas conforme a especificação
            colunas = ["Loja Origem", "Loja Destino", "Quantidade", "Codigo", "Cond. Pagamento", "Operacao", "Status"]

            # Cria um dataframe vazio com as colunas desejadas
            df = pd.DataFrame(columns=colunas)

            # Salva em Excel
            df.to_excel(file_path, index=False, engine='openpyxl')

            # Ajusta o tamanho das colunas
            workbook = load_workbook(file_path)
            worksheet = workbook.active
            for col in worksheet.columns:
                col_letter = col[0].column_letter  # Pega a letra da coluna (A, B, C, etc.)
                worksheet.column_dimensions[col_letter].width = 20
            workbook.save(file_path)

            # Abre a pasta de downloads no Windows
            subprocess.Popen(f'explorer "{download_folder}"')

            print("Layout baixado com sucesso!")
        except Exception as e:
            print(f"Erro ao baixar o layout: {e}")

    def start_internal_movement(self):
        folder = self.folder_entry.get()
        type = self.list_type_var.get()
        InternalMovimentProcess(folder, type)

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    InternalMovementPage(root)
    root.mainloop()