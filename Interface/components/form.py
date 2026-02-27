import customtkinter 
from fonts.fonts import Fonts

class form(customtkinter.CTkFrame):
    fonts = Fonts()
    def __init__(self,master,**kwargs):
        super().__init__(master,**kwargs)
        self.inputs = {}

        self.inputBox(
            label="User:",
            entry="User:"
        )
        self.inputBox(
            label="Password:",
            entry="Password:"
        )
        self.button = customtkinter.CTkButton(self,text="Enviar!",width=200,height=40,font=(self.fonts.BUTTON),command=self.login_action)
        self.button.pack(pady=20)


    def login_action(self):
        dados = self.get_values()
        print(f"Tentativa de login com: {dados}")
    
    #Junção de dois componentes
    def inputBox(self, label:str,entry:str):
        self.label(label)
        self.inputs[label] = self.entry(entry)

    #Componentes
    def label(self, text:str):
        ctkLabel = customtkinter.CTkLabel(self,text=text,font=(self.fonts.LABEL))
        ctkLabel.pack(padx=20, pady=(10, 0), anchor="w")

    def entry(self,prompt:str):
        inputBox = customtkinter.CTkEntry(self, placeholder_text=prompt,font=(self.fonts.ENTRY), width=300, height=40)
        inputBox.pack(padx=20,pady=5, fill="x")
        return inputBox
    #Retornos dos valores
    def get_values(self):
        return {name: widget.get() for name, widget in self.inputs.items()}