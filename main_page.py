from tkinter import *
from price_search_page import PriceSearchPage
from internal_movement_page import InternalMovementPage
from price_alteration_page import PriceAlterationPage

class MainPage:
    def __init__(self, root, type):
        self.type = type
        self.window = Toplevel(root)
        self.window.title('Página Principal')
        self.window.configure(background='DeepPink2')
        self.window_position(root, self.window)
        self.window.resizable(True, True)
        self.window.minsize(width=500, height=200)

        self.create_widgets()

    def create_widgets(self):
        frame = Frame(self.window, bg='white')
        frame.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

        Button(frame, text="Consulta Preço", bd=3, command=self.open_price_search).place(relx=0.2, rely=0.3, relwidth=0.2, relheight=0.3)
        Button(frame, text="Mov. Interna", bd=3, command=self.open_internal_movement).place(relx=0.4, rely=0.3, relwidth=0.2, relheight=0.3)
        Button(frame, text="Alteração Preço", bd=3, command=self.open_price_alteration).place(relx=0.6, rely=0.3, relwidth=0.2, relheight=0.3)

    def window_position(self, principal_window, top_level):
        x = principal_window.winfo_x()
        y = principal_window.winfo_y()
        top_level.geometry(f"500x200+{x}+{y}")
        top_level.grab_set()
    
    def open_price_search(self):
        subpage = PriceSearchPage(self.window)
        self.window_position(self.window, subpage.window)

    def open_internal_movement(self):
        subpage = InternalMovementPage(self.window)
        self.window_position(self.window, subpage.window)

    def open_price_alteration(self):
        subpage = PriceAlterationPage(self.window)
        self.window_position(self.window, subpage.window)

    def run(self):
        self.root.mainloop()