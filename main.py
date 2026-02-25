import os
import openpyxl
import glob

class folder:
    def __init__(self, name:str):
        try:
            os.mkdir(f"./{name}")
            print("Folder %s created suscefully! " %f"./{name}")
        except FileExistsError:
            print("Folder %s already exists " % f"./{name}")

    def generateFolders(self, title, names:list): 
        try: 
            os.mkdir(f'./{title}')
            print("Folder %s created suscefully! " %f'./{title}')
            for name in names:
                try:
                    os.mkdir(f'./{title}/{name}')
                    print("Folder %s created suscefully! " %f'./{title}/{name}')
                except FileExistsError:
                    print("Folder %s already exists " % f"./{name}")
        except FileExistsError:
            print("Folder %s already exists " % f'./{title}')

        
class excel:
    def reader(self,folderReader:str) -> dict:
        listNames = []

        arquivos = glob.glob(f'./{folderReader}/*.xlsx')
        
        if not arquivos:
            print("Folder's not found!")
            return # type: ignore

        folderPath = arquivos[0]

        wb = openpyxl.load_workbook(folderPath)
        sheet = wb.active

        for linha in sheet.iter_rows(min_row=2, values_only=True): # pyright: ignore[reportOptionalMemberAccess]
            listNames.append(linha[0])

        return {'Title': sheet.title,'Names': listNames} # type: ignore


def main():
    print("Program Iniciated! \n")
    folderReader = "Leitor de Planilha"

    fold = folder(folderReader)
    excelSheet = excel().reader(folderReader) # type: ignore
    
    title = excelSheet['Title']
    names = excelSheet['Names']
    
    print(title)
    print(names)
    fold.generateFolders(title,names) 
    



if __name__ == "__main__":
     print("="*24)
     main()
     print("="*24)