from tkinter import *
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import load_workbook
import subprocess
from price_alteration_process import PriceAlterationProcessor



class PriceAlterationPage:
    def __init__(self, root):
        # Configuração da janela
        self.window = Toplevel(root)
        self.window.title("Alteracao de Preco")
        self.window.configure(background='DeepPink2')
        self.window.geometry("500x200")
        self.window.resizable(True, True)
        self.window.minsize(width=500, height=200)

        # Frame principal
        frame = Frame(self.window)
        frame.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

        # Inclusão do arquivo de entrada
        Label(frame, text="Arquivo").place(relx=0.0, rely=0.1, relwidth=0.3, relheight=0.1)
        self.folder_entry = Entry(frame)
        self.folder_entry.place(relx=0.3, rely=0.1, relwidth=0.5, relheight=0.1)

        Button(frame, text="...", command=self.open_file_dialog).place(relx=0.8, rely=0.1, relwidth=0.1, relheight=0.1)

        # Botão para baixar o layout
        Button(frame, text="Layout", bd=3, command=self.download_layout).place(
            relx=0.5, rely=0.75, relwidth=0.2, relheight=0.15
        )

        # Botão para iniciar a movimentação
        Button(frame, text="Iniciar", bd=3, command=self.start_pri).place(
            relx=0.7, rely=0.75, relwidth=0.2, relheight=0.15
        )

    def open_file_dialog(self):
        file_path = filedialog.askopenfilename(parent=self.window, title="Selecione o arquivo")
        if file_path:
            self.folder_entry.delete(0, END)  # Limpa a entrada
            self.folder_entry.insert(0, file_path)  # Insere o caminho do arquivo selecionado

    def download_layout(self):
        try:
            # Defina o nome do arquivo e o caminho para a pasta de downloads
            download_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            file_path = os.path.join(download_folder, "Layout_alteracao_preco.xlsx")
            
            # Defina os nomes das colunas conforme a imagem
            colunas = [
                "Tipo do Codigo", "Produto/Grupo", "Vl. Custo", "Vl. Revenda", "Loja/Grupo", "Data inicio" ,"Status"
            ]

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

    def start_price_alteration(self):
        folder = self.folder_entry.get()
        process = PriceAlterationProcessor(folder)
        process.analisar_planilha()