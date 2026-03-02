import customtkinter
from form import form

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("my-app")
        self.geometry("750x750")
       
        # O segredo para centralizar com pack é o expand=True
        self.meu_form = form(master=self, width=400, corner_radius=15)
        self.meu_form.pack(pady=20, expand=True) 

    def mostrar_dados(self):
        dados = self.meu_form.get_values()
        print(f"Usuário: {dados['User']}")
        print(f"Senha: {dados['Password']}")

app = App()
app.mainloop()