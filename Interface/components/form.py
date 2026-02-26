import customtkinter 

class form(customtkinter.CTkFrame):
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
        self.button = customtkinter.CTkButton(self,text="Enviar!",command=self.login_action)
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
        ctkLabel = customtkinter.CTkLabel(self,text=text)
        ctkLabel.pack()

    def entry(self,prompt:str):
        inputBox = customtkinter.CTkEntry(self, placeholder_text=prompt)
        inputBox.pack(padx=10,pady=20)
        return inputBox
    #Retornos dos valores
    def get_values(self):
        return {name: widget.get() for name, widget in self.inputs.items()}