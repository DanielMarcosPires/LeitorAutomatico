import customtkinter
from Interface.Fonts.fonts import Fonts
import tkinter.messagebox
from Interface.colors import FormColors

class form(customtkinter.CTkFrame):
    fonts = Fonts()
    def __init__(self,master,**kwargs):
        super().__init__(master,**kwargs)

        # Configurar cores do formulário (estilo Excel)
        self.configure(fg_color=FormColors.FRAME_BACKGROUND, border_width=2, border_color=FormColors.FRAME_BORDER)
        
        self.inputs = {}
        self.title("Login - DaniTechnologia")

        self.inputBox(
            label="Usuário:",
            entry="Digite seu usuário",
            dataType="username"
        )
        self.inputBox(
            label="Senha:",
            entry="Digite sua senha",
            dataType="password"
        )
        self.button = customtkinter.CTkButton(
            self,
            text="Entrar", 
            width=200,
            height=40,
            font=(self.fonts.BUTTON),
            fg_color=FormColors.BUTTON_BACKGROUND,  # Azul Excel
            hover_color=FormColors.BUTTON_HOVER,  # Azul mais escuro no hover
            text_color=FormColors.BUTTON_TEXT,
            command=self.login_action
        )
        self.button.pack(pady=20)


    def login_action(self):
        dados = self.get_values()

        username = dados["username"]
        password = dados["password"]

        usernameKey = "DanielMarcos"
        passwordKey = "1234"

        if username == usernameKey and password == passwordKey:
            # Login bem-sucedido, notificar o app
            self.master.login_success()  # type: ignore
            return True
        else:
            # Mostrar mensagem de erro
            tkinter.messagebox.showerror("Erro", "Credenciais inválidas!")
            return False
        
    #Junção de dois componentes
    def inputBox(self, label:str,dataType:str,entry:str):
        self.label(label)
        self.inputs[dataType] = self.entry(entry)

    #Componentes

    def title(self, text:str):
        # Título com estilo Excel
        ctkTitle = customtkinter.CTkTextbox(
            self, 
            height=40, 
            font=("Inter", 25, "bold"), 
            fg_color=FormColors.TITLE_BACKGROUND,  # Azul Excel header
            text_color=FormColors.TITLE_TEXT  # Texto branco para contraste
        ) 
        ctkTitle.tag_config("center", justify='center')
        ctkTitle.insert("0.0", text)
        ctkTitle.tag_add("center", "1.0", "end")
        ctkTitle.configure(state="disabled")
        
        ctkTitle.pack(fill="x", padx=20, pady=(20, 20))

    def label(self, text:str):
        ctkLabel = customtkinter.CTkLabel(
            self,
            text=text,
            font=(self.fonts.LABEL),
            text_color=FormColors.LABEL_TEXT  # Texto branco para contraste no fundo escuro
        )
        ctkLabel.pack(padx=20, pady=(10, 0), anchor="w")

    def entry(self, prompt:str):
        inputBox = customtkinter.CTkEntry(
            self, 
            placeholder_text=prompt, 
            font=(self.fonts.ENTRY), 
            width=350, 
            height=45,
            fg_color=FormColors.ENTRY_BACKGROUND,  # Fundo escuro para inputs
            border_color=FormColors.ENTRY_BORDER,  # Borda azul Excel
            text_color=FormColors.ENTRY_TEXT  # Texto branco
        )
        inputBox.pack(padx=20, pady=10) 
        return inputBox
    #Retornos dos valores
    def get_values(self):
        return {name: widget.get() for name, widget in self.inputs.items()}