import customtkinter

from Fonts.fonts import Fonts


class form(customtkinter.CTkFrame):
    fonts = Fonts()
    def __init__(self,master,**kwargs):
        super().__init__(master,**kwargs)
        self.inputs = {}
        self.title("Formulario")

        self.inputBox(
            label="Insert your username:",
            entry="Username",
            dataType="username"
        )
        self.inputBox(
            label="Insert your password:",
            entry="*******",
            dataType="password"
        )
        self.button = customtkinter.CTkButton(self,text="Enviar!",width=200,height=40,font=(self.fonts.BUTTON),command=self.login_action)
        self.button.pack(pady=20)


    def login_action(self):
        dados = self.get_values()

        username = dados["username"]
        password = dados["password"]

        usernameKey = "DanielMarcos"
        passwordKey = "1234"

        if username == usernameKey and password == passwordKey :
           return True
        return False

    #Junção de dois componentes
    def inputBox(self, label:str,dataType:str,entry:str):
        self.label(label)
        self.inputs[dataType] = self.entry(entry)

    #Componentes

    def title(self, text:str):
    # Adicionei pady=(30, 10) para dar espaço no topo
        ctkTitle = customtkinter.CTkTextbox(self, height=40, font=("Inter", 25, "bold"), fg_color="transparent") 
        ctkTitle.tag_config("center", justify='center')
        ctkTitle.insert("0.0", text)
        ctkTitle.tag_add("center", "1.0", "end")
        ctkTitle.configure(state="disabled")
        
        # pady=(distância_topo, distância_baixo)
        ctkTitle.pack(fill="x", padx=20, pady=(20, 20))

    def label(self, text:str):
        ctkLabel = customtkinter.CTkLabel(self,text=text,font=(self.fonts.LABEL))
        ctkLabel.pack(padx=20, pady=(10, 0), anchor="w")

    def entry(self, prompt:str):
        # Remova o fill="x" se quiser um tamanho fixo, 
        # ou mantenha se quiser que sigam a largura do Frame pai
        inputBox = customtkinter.CTkEntry(self, placeholder_text=prompt, font=(self.fonts.ENTRY), width=350, height=45)
        inputBox.pack(padx=20, pady=10) 
        return inputBox
    #Retornos dos valores
    def get_values(self):
        return {name: widget.get() for name, widget in self.inputs.items()}