import customtkinter
from Interface.colors import ThemeColors
from Interface.dashboard import Dashboard
from Interface.colors import APPEARANCE_MODE, DEFAULT_COLOR_THEME
from Interface.form import form

# Configurar tema Excel-like
customtkinter.set_appearance_mode(APPEARANCE_MODE)  # Modo escuro para melhor contraste
customtkinter.set_default_color_theme(DEFAULT_COLOR_THEME)  # Tema azul como Excel

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Leitor Automático - DaniTechnologia")
        self.geometry("750x750")

        # Configurar cores da janela principal (estilo Excel)
        
        self.configure(fg_color=ThemeColors.APP_BACKGROUND)  # Fundo escuro como Excel moderno
        
        # Instanciar e exibir o formulário de login
        self.meu_form = form(master=self)
        self.meu_form.pack(pady=20, expand=True, fill="both")

        self.dashboard = None

    def mostrar_dados(self):
        dados = self.meu_form.get_values()
        
        print(dados)
        
        print(f"Usuário: {dados['username']}")
        print(f"Senha: {dados['password']}")

    def login_success(self):
        # Esconder o formulário de login
        self.meu_form.pack_forget()
        
        # Mostrar o dashboard
        self.dashboard = Dashboard(master=self)
        self.dashboard.pack(fill="both", expand=True)

app = App()
app.mainloop()