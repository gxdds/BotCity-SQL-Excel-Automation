import pandas as pd

produtos = pd.read_excel("produtos.xlsx")
print(produtos)

for indice, linha in produtos.iterrows():
    nome = linha["NOME"]
    preco = linha["PRECO"]
    quantidade = linha["QUANTIDADE"]
    comentario = linha["COMENTARIO"]

    print("Nome:", nome)
    print("Preço:", preco)
    print("Quantidade:", quantidade)
    print("Comentário:", comentario)