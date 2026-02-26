import customtkinter
from components.form import form
class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("my-app")
        self.geometry("750x750")
        self.grid_columnconfigure((0,1),weight=1)





app = App()
app.mainloop()