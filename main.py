import os
import sys
import openpyxl
import glob

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

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

class excel:
    def reader(self,folderReader:str):
        listNames = []

        arquivos = glob.glob(f'./{folderReader}/*.xlsx')
        
        if not arquivos:
            print("Folder's in %s not found!"%f'./{folderReader}')
            return False

        folderPath = arquivos[0]

        wb = openpyxl.load_workbook(folderPath)
        sheet = wb.active

        for linha in sheet.iter_rows(min_row=2, values_only=True): # pyright: ignore[reportOptionalMemberAccess]
            listNames.append(linha[0])

        return {'Title': sheet.title,'Names': listNames} # type: ignore


def main():
    print(bcolors.OKBLUE+"Program Iniciated! \n")
    folderReader = "Leitor de Planilha"

    fold = folder(folderReader)
    excelSheet = excel().reader(folderReader) # type: ignore
    if excelSheet:
        title = excelSheet['Title']
        names = excelSheet['Names']
       
        fold.generateFolders(title,names) 
    



if __name__ == "__main__":
     print(bcolors.OKBLUE+"="*24)
     main()
     print(bcolors.OKBLUE+"="*24)

sys.exit()