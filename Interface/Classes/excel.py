import glob
import openpyxl

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
