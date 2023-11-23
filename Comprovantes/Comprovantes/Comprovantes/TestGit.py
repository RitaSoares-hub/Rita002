from botcity.core import DesktopBot
from time import sleep
from datetime import datetime
import os
import pandas as pd

tabelafit = pd.read_excel(r'C:\Spool\TST.xlsx', usecols="A")
data_atual = datetime.now()
nome_arq = "logfit_" + data_atual.strftime("%d%H%M%S") + ".csv"
print(f'nome do arquivo {nome_arq}')


# service=Service()
# options = webdriver.ChromeOptions()
# options.add_experimental_option("detach", True)
# navegador = webdriver.Chrome(service=service, options=options)
# url = ("https://secure.d4sign.com.br/login.html")
# navegador.get(url)

class Bot(DesktopBot):
    def action(self, execution=None):
        arq_log_completo = r"C:\Spool" + "\\" + nome_arq
        # with open(f"{arq_log_completo}", 'r') as arquivo_entrada:
        # with open(r'C:\Spool\logfit.txt',
        #          'w') as arquivo_saidafit:
        with open(f"{arq_log_completo}",
                  'w') as arquivo_saidafit:
            # with open(r'C:\Spool\{nome_arq}',
            #         'w') as arquivo_saidafit:
            for index, row in tabelafit.iterrows():
                # noinspection PyTypeChecker
                print(f"Linha {index + 1}: {row['matricula']}")
                teste = "rita"
                if not self.find("NomeAposBusca", matching=0.97, waiting_time=10000):
                    self.not_found("NomeAposBusca")
                    print("não achei Nome Completo")
                    teste="erro"
                    continue

                linhagravar = f"{teste}   \n"
                # Grava arquivo
                arquivo_saidafit.write(linhagravar)
        print("vai fechar o arquivo")
        arquivo_saidafit.close()

    def not_found(self, param):
        pass


if __name__ == '__main__':
    Bot.main()
