from botcity.core import DesktopBot
from time import sleep
from datetime import datetime
import os
import pandas as pd

tabelafit = pd.read_excel(r'C:\Spool\FIT.xlsx',usecols="A")
data_atual = datetime.now()
nome_arq = "logfit_" + data_atual.strftime("%d%H%M%S") + ".csv"
print(f'nome do arquivo {nome_arq}')

#service=Service()
#options = webdriver.ChromeOptions()
#options.add_experimental_option("detach", True)
#navegador = webdriver.Chrome(service=service, options=options)
#url = ("https://secure.d4sign.com.br/login.html")
#navegador.get(url)

class Bot(DesktopBot):
    def action(self, execution=None):
        arq_log_completo = r"C:\Spool" + "\\" + nome_arq
        #with open(f"{arq_log_completo}", 'r') as arquivo_entrada:
        #with open(r'C:\Spool\logfit.txt',
        #          'w') as arquivo_saidafit:
        with open(f"{arq_log_completo}",
                  'w') as arquivo_saidafit:
        #with open(r'C:\Spool\{nome_arq}',
        #         'w') as arquivo_saidafit:
            for index, row in tabelafit.iterrows():
                # noinspection PyTypeChecker
                print(f"Linha {index + 1}: {row['matricula']}")
                matr = str(row['matricula'])
                self.type_up()
                self.type_up()
                sleep(1)
                self.click_at(x=470, y=220)  # Posicao Clientes
                sleep(2)
                self.click_at(x=262, y=318) #Posicao para informar matricula 80%
                self.kb_type(matr)
                sleep(1)
                self.click_at(x=615, y=318) #Posicao para buscar matricula 80%
                sleep(3)

                if not self.find( "NomeAposBusca", matching=0.97, waiting_time=10000):
                    self.not_found("NomeAposBusca")
                    print("não achei Nome Completo")
                    continue
                self.click()

                y = self.get_last_y()
                x = self.get_last_x()
                #print(f'posição nome y= 420: {y}')
                self.click_at(x=x, y=y+30)  # Posicao para clicar sobre o cliente
                sleep(4)
                self.type_down()
                self.type_down()
                sleep(0.5)
                
                if  self.find( "NenhumCliente", matching=0.97, waiting_time=2000):
                    self.not_found("NenhumCliente")
                    linhagravar = f"{matr} ; sem dados  \n"
                    #Grava arquivo
                    arquivo_saidafit.write(linhagravar)
                    print("matricula sem dados")
                    #Clicar em Clientes novamente para limpar campo
                    sleep(1)
                    if not self.find( "Clientes", matching=0.97, waiting_time=10000):
                        self.not_found("Clientes")
                        print('Não conseguiu voltar clientes, {matr}')
                        break
                    self.click()
                     
                    #if not self.find( "Clientes", matching=0.97, waiting_time=10000):
                    #    self.not_found("Clientes")
                    #    print('Não conseguiu voltar tela clientes, {matr}')
                    #    break
                    #self.click()

                    continue

                if not self.find( "NomeCompleto", matching=0.97, waiting_time=10000):
                    self.not_found("NomeCompleto")
                    print("nao encontrei nome")
                    break
                self.click()
                
                y = self.get_last_y()
                #print(f'posição nome y= 420: {y}')

                #NOME
                # self.click_at(x=240, y=y)  # Posicao para copiar nome
                # sleep(1)
                # self.mouse_down()
                # sleep(1)
                # self.mouse_move(x=540, y=y)
                # sleep(1)
                # self.mouse_up()
                # sleep(1)
                # self.control_c()
                # nome = self.get_clipboard()
                # tamanho = len(nome)
                # nome=nome[15:tamanho]
                # print(nome)
                # sleep(1)
                # x = self.get_last_x()
                # y = self.get_last_y()
                # print(f'Posicao nome is: {x}, {y}')

                # Mail


                self.click_at(x=240, y=y+20)  # Posicao para copiar mail
                sleep(0.5)
                self.mouse_down()
                sleep(0.5)
                self.mouse_move(x=480, y=y+20)
                sleep(0.5)
                self.mouse_up()
                sleep(0.5)
                self.control_c()
                mail = self.get_clipboard()
                tamanho = len(mail)
                mail=mail[8:tamanho]
                #print(mail)

                # CPF
                self.click_at(x=240, y=y+50)  # Posicao para copiar mail
                sleep(0.5)
                self.mouse_down()
                sleep(0.5)
                self.mouse_move(x=420, y=y+50)
                sleep(0.5)
                self.mouse_up()
                sleep(0.5)
                self.control_c()
                cpf = self.get_clipboard()
                tamanho = len(cpf)
                cpf = cpf[5:tamanho]
                #print(cpf)

                # # Data Nascimento
                # self.click_at(x=240, y=y+70)  # Posicao para copiar mail
                # sleep(1)
                # self.mouse_down()
                # sleep(1)
                # self.mouse_move(x=420, y=y+70)
                # sleep(1)
                # self.mouse_up()
                # sleep(1)
                # self.control_c()
                # data = self.get_clipboard()
                # tamanho = len(data)
                # data = data[20:tamanho]
                # print(data)

                # endereço
                # self.click_at(x=240, y=y+95)  # Posicao para copiar mail
                # sleep(1)
                # self.mouse_down()
                # sleep(1)
                # self.mouse_move(x=480, y=y+95)
                # sleep(1)
                # self.mouse_up()
                # sleep(1)
                # self.control_c()
                # endereco = self.get_clipboard()
                # tamanho = len(endereco)
                # endereco = endereco[10:tamanho]
                # print(endereco)
                #
                # # cidade
                # self.click_at(x=240, y=y+120)  # Posicao para copiar mail
                # sleep(1)
                # self.mouse_down()
                # sleep(1)
                # self.mouse_move(x=460, y=y+120)
                # sleep(1)
                # self.mouse_up()
                # sleep(1)
                # self.control_c()
                # cidade = self.get_clipboard()
                # tamanho = len(cidade)
                # cidade = cidade[8:tamanho]
                # print(cidade)

                # Telefone
                self.click_at(x=240, y=y+170)  # Posicao para copiar mail
                sleep(0.5)
                self.mouse_down()
                sleep(0.5)
                self.mouse_move(x=520, y=y+170)
                sleep(0.5)
                self.mouse_up()
                sleep(0.5)
                self.control_c()
                telefone = self.get_clipboard()
                tamanho = len(telefone)
                telefone = telefone[11:tamanho]
                print(f'telefone {telefone}')
                sleep(1)
                self.page_up()
                if not self.find( "Clientes@", matching=0.97, waiting_time=10000):
                    self.not_found("Clientes@")
                    continue
                self.click()
                
                #if not self.find( "Clientes", matching=0.97, waiting_time=10000):
                #    self.not_found("Clientes")
                #    print('Não conseguiu voltar tela clientes, {matr}')
                #    break
                #self.click()
                # sleep(5)
                # x = self.get_last_x()
                # y = self.get_last_y()
                # print(f'Posicao cliente is: {x}, {y}')

                sleep(1)
                #x = self.get_last_x()
                #y = self.get_last_y()
                #print(f'The last saved mouse position is: {x}, {y}')

                #x = self.get_last_x()
                #y = self.get_last_y()
                #print(f'The last saved mouse position is: {x}, {y}')

                #tel = "a "

                #linhagravar = f"{matr} ; {nome} ; {mail} ; {endereco} ; {cidade} ; {data}; {telefone}  \n"
                linhagravar = f"{matr} ; {mail} ; {cpf}; {telefone}  \n"
                #Grava arquivo
                arquivo_saidafit.write(linhagravar)
        print("vai fechar o arquivo")
        arquivo_saidafit.close()

    def not_found(self, param):
        pass

if __name__ == '__main__':
    Bot.main()


























