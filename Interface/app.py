import customtkinter
from components.form import form

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("my-app")
        self.geometry("750x750")
        # Instancia o componente passando 'self' como pai
        self.meu_form = form(master=self)
        self.meu_form.pack(expand=True)
      

    def mostrar_dados(self):
        dados = self.meu_form.get_values()
        print(f"Usuário: {dados['User']}")
        print(f"Senha: {dados['Password']}")

app = App()
app.mainloop()