import os

def initialFolder(name:str) -> str:
    try:
        os.mkdir(f"./{name}")
        return "Folder %s created! " % f"./{name}"
    except FileExistsError:
        return "Folder %s already exists " % f"./{name}"

def folderChildrens(loops:str):
    stockNames = []
    numberLoops = int(loops)

    print("="*12)
    for i in range(numberLoops):
        print(f"StockNames: {stockNames}\n")
        name = input(f"(${i+1}) Insira o nome do aluno:\n> ")
        if not name == "":
            stockNames.append(name)
        else:
            print("Não deve ser vazio!")
            break
    print("="*12)
    return stockNames

def createfoldersNames(folders:list,folderPrincipal:str):
    try:
        for name in folders:
            paths = f"./{folderPrincipal}/{name}"
            os.mkdir(paths)
        return "Folder's %s created!" 
    except FileExistsError:
        return "Folder %s already exists"

def main():
    folderPrincipal = input("Nome da pasta:\n> ")
    quantity = input("Insira o número de pastas:\n> ")

    folderCreated = initialFolder(folderPrincipal)
    foldersNames = folderChildrens(quantity)
    folderNamesCreated = createfoldersNames(foldersNames,folderPrincipal)
    
    print(folderNamesCreated)

if __name__ == "__main__":
    main()