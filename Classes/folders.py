import os
from Classes.bgColors import bcolors

class folder:
    def __init__(self, name:str):
        try:
            os.mkdir(f"./{name}")
            print("Folder %s created suscefully! " %f"./{name}")
        except FileExistsError:
            print(bcolors.HEADER+ "Folder %s already exists " % f"./{name}")

    def generateFolders(self, title, names:list): 
        try: 
            os.mkdir(f'./{title}')
            print(bcolors.WARNING+"Folder %s created suscefully! \n" %f'./{title}')
            for name in names:
                try:
                    os.mkdir(f'./{title}/{name}')
                    print("Folder %s created suscefully! " %f'./{title}/{name}')
                except FileExistsError:
                    print("Folder %s already exists " % f"./{name}")
        except FileExistsError:
            print("Folder %s already exists " % f'./{title}')