from botcity.core import DesktopBot
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *



class Bot(DesktopBot):
    def action(self, execution=None):

        # lembre de atualizar os códios para o programa que quer abrir no seu computador
        self.execute(r'C:\Users\joaol\Desktop\Microsoft Teams.lnk')

        if not self.find( "pesquisar", matching=0.97, waiting_time=10000):
            self.not_found("pesquisar")
        self.click()
        self.paste("alon pinheiro")

        if not self.find( "alon", matching=0.97, waiting_time=10000):
            self.not_found("alon")
        self.click()

        if not self.find( "campo mensagem", matching=0.97, waiting_time=10000):
            self.not_found("campo mensagem")
        self.click()

        self.paste("Coe Alon, to automatizando tudo")

        self.enter()


    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()

































