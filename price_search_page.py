from tkinter import *

class PriceSearchPage:
    def __init__(self, root):
        self.window = Toplevel(root)
        self.window.title("Consulta Preço")
        self.window.configure(background='DeepPink2')
        self.window.geometry("500x200")
        self.window.resizable(True, True)
        self.window.minsize(width=500, height=200)

        # Frame and elements
        frame = Frame(self.window)
        frame.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

        Label(frame, text="Arquivo").place(relx=0.0, rely=0.1, relwidth=0.3, relheight=0.1)
        folder_entry = Entry(frame)
        folder_entry.place(relx=0.3, rely=0.1, relwidth=0.6, relheight=0.1)

        Label(frame, text="Selecionar Loja").place(relx=0.0, rely=0.35, relwidth=0.3, relheight=0.1)
        loja_combobox = ttk.Combobox(frame, values=["Loja 1", "Loja 2", "Loja 3"])
        loja_combobox.place(relx=0.3, rely=0.35, relwidth=0.6, relheight=0.1)

        Button(frame, text="Iniciar", command=lambda: self.process_search(folder_entry.get(), loja_combobox.get())).place(
            relx=0.7, rely=0.75, relwidth=0.2, relheight=0.15)

    def process_search(self, folder, loja):
        print(f"Processing search with folder: {folder}, store: {loja}")