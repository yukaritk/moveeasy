from tkinter import *
from main_page import MainPage
from credential_manager import CredentialManager

class LoginPage:
    def __init__(self):
        self.root = Tk()
        self.root.title('Login')
        self.root.configure(background='DeepPink2')
        self.root.geometry("500x200")
        self.root.resizable(True, True)
        self.root.minsize(width=500, height=200)

        self.cred_manager = CredentialManager()

        self.create_widgets()
        self.load_credentials()

    def create_widgets(self):
        frame = Frame(self.root, bg='white')
        frame.place(relx=0.015, rely=0.03, relwidth=0.97, relheight=0.94)

        Label(frame,  bg='white',text="Nome de Usuário", anchor='w').place(relx=0.2, rely=0.2)
        self.username_entry = Entry(frame)
        self.username_entry.place(relx=0.5, rely=0.2, relwidth=0.3, relheight=0.1)

        Label(frame, bg='white', text="Senha", anchor='w').place(relx=0.2, rely=0.4)
        self.password_entry = Entry(frame, show="*")
        self.password_entry.place(relx=0.5, rely=0.4, relwidth=0.3, relheight=0.1)

        login_button = Button(frame, text="Logar", bd=3, command=self.login)
        login_button.place(relx=0.5, rely=0.6, relwidth=0.2)

    def load_credentials(self):
        username, password = self.cred_manager.load_credentials()
        if username:
            self.username_entry.insert(0, username)
        if password:
            self.password_entry.insert(0, password)

    def login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()
        self.cred_manager.save_credentials(username, password)

        main_page = MainPage(self.root, username)
        main_page.run()

    def run(self):
        self.root.mainloop()

# Ponto de entrada do programa
if __name__ == "__main__":
    app = LoginPage()
    app.run()