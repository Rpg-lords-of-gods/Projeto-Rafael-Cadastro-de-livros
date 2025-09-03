import openpyxl

# Cria um novo arquivo Excel
workbook = openpyxl.Workbook()

# Seleciona a planilha ativa (padrão é chamada "Sheet")
sheet = workbook.active
sheet.title = "Livros"  # renomeia a aba

# Adiciona um cabeçalho
sheet.append(["Título", "Autor", "Ano", "Preço", "Quantidade"])

# Salva o arquivo
workbook.save("livros.xlsx")

print("Arquivo Excel criado com sucesso!")
