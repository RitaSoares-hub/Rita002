from botcity.core import DesktopBot
import os
import pyautogui
from datetime import datetime
from shutil import copyfile
import PyPDF2
import logging
import win32com.client as win32
#import sys
from dotenv import load_dotenv
load_dotenv(override=True)
usuario = os.getenv("SAP_USER")
senha = os.getenv("USER_PSW")

print(usuario,senha)


caminho_compr = r'I:\TI\TI_Sistemas\Rita\Comprovantes'
diretoriodt = ""



def not_found(label):
    print(f"Element not found: {label}")


def criar_diretorio_com_data():
    global diretoriodt, arq_log
    data_atual = datetime.now()
    datalog = data_atual.strftime("%Y") + "-" + data_atual.strftime("%m") + "-" + data_atual.strftime("%d")
    print("data atual: ", data_atual)
    nome_diretorio = data_atual.strftime("%Y%m")
    diretoriodt = f"{nome_diretorio}"
    arq_log = "log_execucao_python" + "_" + diretoriodt
    logging.basicConfig(level=logging.INFO, filename=arq_log,
                        format="%(asctime)s - %(levelname)s - %(message)s ")
    print(diretoriodt)
    print(arq_log)

    try:
        global caminho_compr
        # global diretoriodt
        diretorio_atual = os.getcwd()
        os.chdir(caminho_compr)
        print(f"Diretório atual: {diretorio_atual}")
        os.mkdir(diretoriodt)
        print(f"O diretório '{diretoriodt}' foi criado com sucesso.")
        logging.info(f"O diretório '{diretoriodt}' foi criado com sucesso.")
    except FileExistsError:
        print(f"O diretório '{diretoriodt}' já existe.")
        logging.info(f"O diretório '{diretoriodt}' já existe.")
        # raise SystemExit("Programa cancelado .")


class Bot(DesktopBot):
    def action(self, execution=None):
        # Inicio bot de comprovantes
        # busca documentos na pasta
        global caminho_compr, empresa, diretoriodt, fornecedor, motivo, ini_valor, valor, partidaorig, nomearqren_re, arq_log
        lista_arquivos = os.listdir(caminho_compr)
        print(lista_arquivos)
        print(diretoriodt)
        print(arq_log)
        contador = 1
        # #################      Logon SAP
        self.type_windows()
        self.wait(2000)
        self.kb_type("SAP logon")
        self.enter()

        if not self.find( "TelaLogonSap", matching=0.97, waiting_time=10000):
            self.not_found("TelaLogonSap")
        
        self.wait(2000)
        self.type_down()
        self.enter()
        self.wait(2000)
        self.shift_tab()
        self.kb_type("200")
        self.tab()
        self.kb_type(usuario)
        self.tab()
        self.kb_type(senha)
        self.enter()

        
        
        
        if self.find( "JaLogado", matching=0.97, waiting_time=10000):
            print("achei cnpjja logado")
            exit(10)

        print("Logou")
        


        for arquivo in lista_arquivos:
            if "10052023" in arquivo:  # Seleciona os arquivos com data do dia
                print(contador)
                print(arquivo)
                num_f110 = arquivo[15:25]
                cnpj = arquivo[:14]
                datab = arquivo[49:57]
                ano = arquivo[53:57]
                print("CNPJ ", cnpj)
                print("DATAB ", datab)
                print("ANO ", ano)
                print("partida arquivo", num_f110)
                if "59291534" in cnpj:
                    empresa = "1000"
                elif cnpj == "13718634000126":
                    empresa = "1001"
                elif cnpj == "00278017000105":
                    empresa = "1009"
                elif cnpj == "20854704000139":
                    empresa = "1002"
                elif cnpj == "19760435000162":
                    empresa = "1003"
                elif cnpj == "26866107000100":
                    empresa = "1005"
                elif cnpj == "13396435000149":
                    empresa = "1006"
                elif cnpj == "39999769000109":
                    empresa = "1007"
                else:
                    logging.info(f"Empresa não SAP '{cnpj}' ")
                    continue
                # "31439951000195" CBAUTOMOTIVE GASTAO VIDIGAL
                #num_partida = "1200000007"
                #print("Numero da partida ", num_partida)
                contador = contador + 1
                #############  leitura pdv
                nomearq = caminho_compr + f"\{arquivo}"
                arq_pdf = open(nomearq, 'rb')  # leitura pdf 10082023
                pdf = PyPDF2.PdfReader(arq_pdf)
                pagina = pdf.pages[0]
                linhas = pagina.extract_text().split("\n")
                for item in linhas:
                    if "CNPJ" in item:
                        print("achei cnpj", item)
                        pos_cnpj = item.find("CNPJ")  # posição onde inicia CNPJ
                        fornecedor = (item)[0:pos_cnpj]
                    elif "Modalidade:" in item:
                        print("achei modalidade == ", item)
                        pos_modalidade = item.find(":") + 2  # posição onde termina texto modalidade
                        tamanho = len(item)
                        motivo = (item)[pos_modalidade:tamanho]  # modalidade
                    elif "Valor do documento:" in item:
                        print("achei valor == ", item)
                        tamanho = len(item)
                        ini_valor = item.find("Valor do documento:") + 22
                        valor = item[ini_valor:ini_valor + 15]
                    else:
                        pass

                #print(linhas[5])
                #print(linhas[9])
                #cnpj = linhas[5].find("CNPJ")  # posição onde inicia CNPJ
                #modalidade = linhas[9].find(":") + 2  # posição onde termina texto modalidade
                #fornecedor = (linhas[5])[0:cnpj]  # nome do fornecedor sem final CNPJ
                #tamanho = len(linhas[9])  # tamanho da linha da Modalidade - palavra modalidade
                #motivo = (linhas[9])[modalidade:tamanho]  # modalidade
                #ini_valor = linhas[12].find("R$ ") + 3
                #valor = (linhas[12])[ini_valor:ini_valor + 15]
                print("fornecedor: ", fornecedor)
                print("modalidade: ", motivo)
                print("inicio pos valor: ", ini_valor)
                print("valor: ", valor)
                if motivo == "Tributo - IPTU - Prefeituras":
                    desp = "IPTU"
                elif motivo == "Pagamento de Contas e Tributos com Código de Barras":
                    desp = "CONSUMO"
                else:
                    desp = "OUTRASDESPESAS"
                print("despesa: ", desp)
                arq_pdf.close()
                print(nomearq)
                logging.info(f"Arquivo lido: {nomearq}")
                ##### fim leitura pdf

                #
                # Buscando numero da partida original
                #
                self.click_at(x=69, y=54)
                self.wait(1000)
                self.kb_type("SE16N")
                self.enter()
                self.wait(1000)
                x = self.get_last_x()
                y = self.get_last_y()
                print(f'The last saved mouse position se16n: {x}, {y}')
                self.click_at(x=186, y=159)
                self.kb_type("xxxxx")
                self.enter()
                self.wait(1000)
                self.backspace()
                self.wait(1000)
                self.kb_type("REGUP")
                self.enter()


                #
                # if not self.find( "IrparaFB16n", matching=0.97, waiting_time=10000):
                #     self.not_found("IrparaFB16n")
                # self.click()
                # self.type_down()
                # self.type_right()
                # self.enter()
                # if not self.find( "EscolheVariante", matching=0.97, waiting_time=10000):
                #     self.not_found("EscolheVariante")
                # self.click()
                # self.wait(1000)
                # if not self.find( "EscolheRegup", matching=0.97, waiting_time=10000):
                #     self.not_found("EscolheRegup")
                # self.click()
                # self.wait(1000)
                # self.enter()
                # self.wait(1000)
                # #self.kb_type("/REGUP")
                # self.enter()
                x = self.get_last_x()
                y = self.get_last_y()
                #print(f'The last saved mouse position xy: {x}, {y}')
                self.wait(1000)
                self.click_at(x=190, y=275)
                #print(f'The last saved mouse position campo ZBUKR: {x}, {y}')
                self.wait(1000)
                self.kb_type("ZBUKR") #inclui campo empresa no top da lista da se16n
                self.enter()
                self.wait(1000)
                self.click_at(x=190, y=275)
                #print(f'The last saved mouse position campo WRBTR: {x}, {y}')
                self.wait(1000)
                self.kb_type("WRBTR") #inclui campo valor no top da lista da se16n
                self.enter(wait=1000)
                self.click_at(x=190, y=275)
                self.wait(2000)
                self.kb_type("VBLNR") #inclui campo documento no top da lista da se16n
                self.enter()
                self.wait(2000)
                x = self.get_last_x()
                y = self.get_last_y()
                #print(f'The last saved mouse position is: {x}, {y}')
                self.click_at(x=190, y=357)
                self.kb_type(empresa)
                self.wait(1000)
                self.enter()
                self.wait(1000)
                self.type_down()
                self.wait(1000)
                self.kb_type(valor)
                self.wait(2000)
                self.enter()
                self.wait(1000)
                self.type_down()
                self.wait(1000)
                self.kb_type(num_f110)
                self.wait(2000)
                self.enter()
                self.wait(1000)
                self.type_down()
                self.wait(1000)
                self.kb_type(datab)
                self.enter(wait=1000)
                #if not self.find( "executase16n", matching=0.97, waiting_time=10000):
                #    self.not_found("executase16n")
                #self.click()

                self.wait(1000)
                if not self.find( "exese16n", matching=0.97, waiting_time=10000):
                    self.not_found("exese16n")
                self.click()
                self.wait(1000)
                if not self.find( "PartidaOrig", matching=0.97, waiting_time=10000):
                    self.not_found("PartidaOrig")
                self.click()
                self.wait(1000)
                self.type_down()
                pyautogui.hotkey('ctrl', 'y')
                self.control_c()
                self.wait(1000)
                num_partida = self.get_clipboard()
                print("Numpartida identificado: ", num_partida)
                if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                    self.not_found("Retorno")
                self.click()
                self.wait(1000)
                if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                    self.not_found("Retorno")
                self.click()
                self.wait(1000)
                # break
                # ######  

                caminho_compr_ren = caminho_compr + f"\{diretoriodt}"
                self.wait(1000)
                self.kb_type("FB03")
                self.enter()
                self.wait(1000)
                self.kb_type(num_partida)
                self.tab()
                self.kb_type(empresa)
                self.tab()
                self.kb_type(ano)
                self.enter()
                # self.wait(2000)
                # if not self.find( "AnexoFB03_01", matching=0.97, waiting_time=10000):
                #     self.not_found("AnexoFB03_01")
                # self.click()
                # self.wait(1000)
                # self.type_down()
                # self.enter()
                # self.enter()
                # self.wait(1000)
                # self.kb_type(nomearq)
                # self.wait(1000)
                # self.enter()
                # self.wait(10000)
                logging.info(f"FB03 - ARQUIVO ANEXADO: {nomearq}")
                if not self.find( "VisaoRazao", matching=0.97, waiting_time=10000):
                    self.not_found("VisaoRazao")
                self.click()
                self.wait(1000)
                if not self.find( "CentroLucro", matching=0.97, waiting_time=10000):
                    self.not_found("CentroLucro")
                self.click()
                self.wait(1000)
                self.type_down()
                pyautogui.hotkey('ctrl', 'y')
                self.control_c()
                self.wait(1000)
                clucro = self.get_clipboard()
                # print("conteudo linha", clucro)
                self.wait(1000)
                cimovel = clucro[4:6]
                if cimovel == "I0":
                    imovel = clucro[6:10]
                    print("imovel", imovel)
                    print("num_partida[0:1]", num_partida[0:1])
                    #if num_partida[0:1] == "1":
                        # ver se partida começa com 1 avançar de 300 ate 342
                        #if not self.find( "Ambiente", matching=0.97, waiting_time=10000):
                        #    self.not_found("Ambiente")
                        #self.click()
                        #self.wait(1000)
                        #print("entrei aqui")
                        #self.type_down()
                        #self.wait(1000)
                        #self.type_down()
                        #self.wait(1000)
                        #self.type_down()
                        #self.wait(1000)
                        #self.enter()
                        #if not self.find( "NDOC", matching=0.97, waiting_time=10000):
                        #    self.not_found("NDOC")
                        #self.click()
                        #self.type_down()
                        # Copiar partida inicial
                        #self.wait(1000)
                        #self.click_at(x=215, y=309)
                        #pyautogui.hotkey('ctrl', 'y')
                        #self.wait(1000)
                        #x = self.get_last_x()
                        #y = self.get_last_y()
                        #print(f'The last saved mouse position is: {x}, {y}')
                        #self.mouse_down()
                        #self.wait(2000)
                        #self.mouse_move(x=280, y=315)
                        #self.wait(2000)
                        # x = self.get_last_x()
                        # y = self.get_last_y()
                        # print(f'The last saved mouse position is: {x}, {y}')
                        #self.mouse_up()
                        #self.wait(2000)
                        #self.control_c()
                        #self.wait(1000)
                        #partidaorig = self.get_clipboard()
                        #logging.info(f"Documento original: {partidaorig}")
                        # partidaorig = "1299999999"
                        #print("part orig", partidaorig)
                        #self.wait(1000)
                        #if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                        #    self.not_found("Retorno")
                        #self.click()
                    self.wait(1000)
                    if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                        self.not_found("Retorno")
                    self.click()
                    self.wait(1000)
                    if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                        self.not_found("Retorno")
                    self.click()
                    self.wait(1000)
                    self.kb_type("REBDBU")
                    self.enter()
                    self.wait(2000)
                    self.kb_type(empresa)
                    self.tab()
                    self.kb_type(imovel)
                    self.tab()
                    self.kb_type("1")
                    self.enter()
                    self.wait(3000)
                    # busca  parcela IPTU
                    print("inscricao")
                    self.wait(1000)
                    #Definindo mome do arquivo quando for RE mas não for relacionado a inscrição.
                    nomearq_re = diretoriodt + "-" + fornecedor + "-" + desp + ".pdf"
                    nomearqren_re = caminho_compr_ren + f"\{nomearq_re}"
                    print("numero partida ",num_partida)
                    if num_partida[0:1] == "1": #Comentar 367 ate 434
                        if not self.find( "AbaIns", matching=0.97, waiting_time=10000):
                            self.not_found("AbaIns")
                            print("nao encontrou AbaIns ", num_partida)
                        self.click()
                        print("numero partida ",num_partida)
                        self.click()
                        self.wait(3000)
                        #print("linha 419 ")
                        pyautogui.click(x=871, y=652)
                        pyautogui.click(x=871, y=652)
                        pyautogui.click(x=871, y=652)
                        pyautogui.click(x=871, y=652)
                        self.wait(2000)
                        # print(pyautogui.position())
                        pyautogui.click(x=155, y=580)
                        self.wait(1000)
                        #print("linha 428 - exercicio")
                        if not self.find( "ajexercicio", matching=0.97, waiting_time=10000):
                            self.not_found("ajexercicio")
                        self.click()
                        self.wait(1000)
                        if not self.find( "exercicio2", matching=0.97, waiting_time=10000):
                            self.not_found("exercicio2")
                        self.click()
                        self.wait(1000)
                        if not self.find( "ajexercicio3", matching=0.97, waiting_time=10000):
                            self.not_found("ajexercicio3")
                        self.click()
                        self.wait(1000)
                        self.kb_type(ano)
                        self.wait(1000)
                        self.enter()
                        self.wait(1000)
                        print("linha 445")
                        pyautogui.click(x=104, y=580)
                        self.wait(1000)
                        self.kb_type(num_partida)
                        #print("num_partida")
                        if not self.find( "BuscaPartida", matching=0.97, waiting_time=10000):
                            self.not_found("BuscaPartida")
                        self.click()

                        if self.find( "encontrapartida", matching=0.97, waiting_time=10000):
                            print("achei")
                            if not self.find("saidacxpesquisa", matching=0.97, waiting_time=10000):
                                self.not_found("saidacxpesquisa")
                            self.click()
                            self.wait(1000)
                            self.type_left()
                            self.type_left()
                            self.type_left()
                            self.type_left()
                            pyautogui.hotkey('ctrl', 'y')
                            self.control_c()
                            parcela = self.get_clipboard()
                            self.type_left()
                            self.type_left()
                            self.type_left()
                            pyautogui.hotkey('ctrl', 'y')
                            self.control_c()
                            inscricao = self.get_clipboard()
                            # parte que renomeia -nao executar
                            # caminho_compr_ren = caminho_compr + f"\{diretoriodt}" definido inicio bloco
                            nomearq_re = diretoriodt + "-" + inscricao + "-" + parcela + "-" + fornecedor + "-" + desp + ".pdf"
                            print("nome renomeado : ", nomearq_re)
                            nomearqren_re = caminho_compr_ren + f"\{nomearq_re}"
                        if self.find("saidacxpesquisa", matching=0.97, waiting_time=10000):
                            #self.not_found("saidacxpesquisa")
                            self.click()
                            print("saida da pesquisa - nao achou")
                            #self.not_found("encontrapartida")
                    #renomear arquivo de acordo com o conteudo
                    copyfile(nomearq, nomearqren_re)
                    print("nome renomeado com caminho : ", nomearqren_re)

                    #anexo

                    if not self.find( "AtualizaEd", matching=0.97, waiting_time=10000):
                        self.not_found("AtualizaEd")
                    self.click()
                    self.wait(3000)
                    if not self.find( "AnexoRebdbu", matching=0.97, waiting_time=10000):
                        self.not_found("AnexoRebdbu")
                    self.click()
                    self.wait(2000)
                    self.type_down()
                    self.enter()
                    self.enter()
                    self.wait(1000)
                    self.kb_type(nomearqren_re)
                    self.wait(1000)
                    self.enter()
                    self.wait(15000)
                    if not self.find( "NaoEncontraAnexo", matching=0.97, waiting_time=10000):
                        self.not_found("NaoEncontraAnexo")
                        if not self.find( "ANEXOOK", matching=0.97, waiting_time=10000):
                            self.not_found("ANEXOOK")                                          
                        logging.info(f"REFX - ARQUIVO ANEXADO: {nomearqren_re}")
                        print("Arquivo anexado : ", nomearqren_re)
                    else:
                        self.click()
                        logging.info(f"ANEXO NAO ENCONTRADO: {nomearqren_re}")
                        print("Arquivo nao encontrato : ", nomearqren_re)
                    self.wait(1000)
                    if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                        self.not_found("Retorno")
                    self.click()
                else:
                    pass
                self.wait(1000)
                if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                    self.not_found("Retorno")
                else:
                    self.click()
                if not self.find( "Retorno", matching=0.97, waiting_time=10000):
                    self.not_found("Retorno")
                else:
                    self.click()
                # nomearqren = caminho_compr_ren + f"\{arquivo}"  #mover para antes do rename
                # os.rename(nomearq, nomearqren)
                self.wait(1000)


        # Copia do arquivo de log para log do REFX  ############################################
        # Abrir o arquivo de entrada para leitura
        arq_log_completo = r"C:\Users\rita.soares\PycharmProjects\GrupoCB\IconRealty\Comprovantes\Comprovantes\Comprovantes" + "\\" + arq_log
        #with open(
        #        r'C:\Users\rita.soares\PycharmProjects\GrupoCB\IconRealty\Comprovantes\Comprovantes\Comprovantes\log_execucao_python.log',
        #with open(r'C:\Users\rita.soares\PycharmProjects\GrupoCB\IconRealty\Comprovantes\Comprovantes\Comprovantes\{arq_log}',
        #          'r') as arquivo_entrada:
        with open(f"{arq_log_completo}", 'r') as arquivo_entrada:
            # Ler todas as linhas do arquivo
            linhas = arquivo_entrada.readlines()
        # Filtrar as linhas que contêm o código "REFX"
        linhas_refx = [linha for linha in linhas if 'REFX' in linha]
        print("linhas_refx",linhas_refx)
        # Abrir o arquivo de saída para escrita
        with open(r'C:\Users\rita.soares\PycharmProjects\GrupoCB\IconRealty\Comprovantes\Comprovantes\Comprovantes\logrefx.txt', 'w') as arquivo_refx:
            if not linhas_refx:
                arquivo_refx.write("Nenhuma linha com o código REFX encontrada.")
            else:
                # Escrever as linhas filtradas no arquivo de saída
                arquivo_refx.writelines(linhas_refx)
        qtd_anexos_re = len(linhas_refx)
        arquivo_entrada.close()
        arquivo_refx.close()
        print(f"quantidade de linhas '{qtd_anexos_re}'.")

        # envio de e-mail ########################################
        outlook = win32.Dispatch('outlook.application')
        # criar um email
        email = outlook.CreateItem(0)

        # configurar as informações do seu e-mail
        #email.To = "rita.soares@grupocb.com.br;hudson.severo@grupocb.com.br"
        email.To = "rita.soares@grupocb.com.br"
        email.Subject = "E-mail automático do Python"
        email.HTMLBody = f"""
            <p>Mensagem programa anexo Comprovantes</p>
    
            <p>Quantidade de documentos anexados no módulo RE:  {qtd_anexos_re}</p>
     
            <p>Fim execução automática</p>
            """

        #anexo = "C://Users/rita.soares/PycharmProjects/pythonProject/Testegui1/logrefx.txt"
        anexo = r'C:\Users\rita.soares\PycharmProjects\GrupoCB\IconRealty\Comprovantes\Comprovantes\Comprovantes\logrefx.txt'
        attachment = anexo
        email.Attachments.Add(attachment)
        email.Send()
        print("Email Enviado")

    def not_found(self, param):
        pass


if __name__ == '__main__':
    criar_diretorio_com_data()
    Bot.main()



