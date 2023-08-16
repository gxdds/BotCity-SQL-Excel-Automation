from botcity.core import DesktopBot
import pandas as pd

produtos = pd.read_excel("produtos.xlsx")
print(produtos)
class Bot(DesktopBot):
    def action(self, execution=None):

        self.execute(r"C:\Users\vinig\AppData\Local\Programs\Microsoft VS Code\Code.exe")

        self.wait(6000)

        if not self.find( "playbttn", matching=0.97, waiting_time=10000):
            self.not_found("playbttn")
        self.click()

        for indice, linha in produtos.iterrows():
            nome = linha["NOME"]
            preco = linha["PRECO"]
            quantidade = linha["QUANTIDADE"]
            comentario = linha["COMENTARIO"]

        if not self.find( "nome", matching=0.97, waiting_time=10000):
            self.not_found("nome")
        self.click_relative(20, 27)
        self.paste(nome)
    
        if not self.find( "preco", matching=0.97, waiting_time=10000):
            self.not_found("preco")
        self.click_relative(23, 27)
        self.paste(preco)
    
        if not self.find( "qtde", matching=0.97, waiting_time=10000):
            self.not_found("qtde")
        self.click_relative(18, 27)
        self.paste(quantidade)

        if not self.find( "coment", matching=0.97, waiting_time=10000):
            self.not_found("comentario")
        self.click_relative(6, 29)
        self.paste(comentario)

        self.wait(2500)

        if not self.find( "add", matching=0.97, waiting_time=10000):
            self.not_found("add")
        self.click()

        self.wait(1000)

        if not self.find( "okbtn", matching=0.97, waiting_time=10000):
            self.not_found("okbtn")
        self.click()
        


    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()



