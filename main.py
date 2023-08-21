
# 19/08/2023
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
#gmail
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
#alerta
import tkinter
import tkinter.filedialog
from tkinter import messagebox
from tkinter.simpledialog import askstring

root = tkinter.Tk()
root.withdraw()

remetente = askstring('gmail', 'digite o seu email')
senhaapp = askstring('senha app', 'digite a senha app')


def verificar_nome(nome_produto, nome_busca, termos_banidos):
    termos_procurados = nome_busca.split(' ')
    termos_banidos = termos_banidos.split(' ')
    nome_produto = nome_produto.lower()
    for termo in termos_banidos:
        if termo in nome_produto:
            return False
    for termo in termos_procurados:
        if not termo in nome_produto:
            return False
    return True

def string_para_float(preco):
    preco = preco.replace('R$', '').replace(
        ' ', '').replace('.', '').replace(',', '.')
    preco = float(preco)
    return preco

def enviar_mail(remetente, senhaapp, destinatario, assunto, texto, arquivos=[], use_tls=True, msg_html=False):
    
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = assunto
    if msg_html:
        msg.attach(MIMEText(texto, 'html'))
    else:
        msg.attach(MIMEText(texto))

    for path in arquivos:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={}'.format(Path(path).name))
        msg.attach(part)

    smtp = smtplib.SMTP('smtp.gmail.com: 587')
    if use_tls:
        smtp.starttls()
    smtp.login(remetente, senhaapp)
    smtp.sendmail(remetente, destinatario, msg.as_string())
    smtp.quit()



# começando a pesquisa 

nav = webdriver.Chrome()

busca_df = pd.read_excel('buscas.xlsx') 

for i in busca_df.index:
    nome_busca = busca_df.loc[i, 'Nome']
    termos_banidos = busca_df.loc[i, 'Termos banidos']

    nav.get('https://www.google.com')

    nav.find_element('tag name', 'textarea').send_keys(
        f'{nome_busca}', Keys.ENTER)

    # entrando na aba shopping
    while True:
        try:
            # clicar na aba shopping
            nav.find_element('link text', 'Shopping').click()
        except:
            continue  # se nao achar o elemento, é porque a pagina nao carregou, entao vai tentar denovo
        finally:
            break

    # pegando cada anuncio da pagina
    while True:
        # criando listas com todos os anuncios
        anuncios = nav.find_elements('class name', 'i0X6df')
        if anuncios:  # quando caregar a pagina e encher a lista, sairá do loop
            break
    time.sleep(1)

    lista_resultados = []

    for anuncio in anuncios:
        nome_produto = anuncio.find_element('class name', 'tAxDx').text
        preco = anuncio.find_element('class name', 'a8Pemb').text
        preco = string_para_float(preco)

        # filtrando os itens de acordo com os dados da tabela
        # filtrando os itens de acordo com os dados da tabela
        if verificar_nome(nome_produto, nome_busca, termos_banidos) and preco >= 2000:
            # acima de 2000 para nao pegar acessorios de celular
            # como nao deu para pegar o href diretamente do elemento, selecionei a tag filho para depois pegar o href da tag pai
            referencia = anuncio.find_element('class name', 'KoNVE')
            link = referencia.find_element('xpath', '..').get_attribute('href')  # href da tag pai
            lista_resultados.append((nome_produto, preco, link))
        else:
            pass

    # criando a tabela com os resultados
    resultado_df = pd.DataFrame(columns=['Nome', 'Preço', 'Link'])

    for item in lista_resultados:
        resultado_df.loc[len(resultado_df)] = item
    
    resultado_df = resultado_df.sort_values("Preço") #ordenando pelo preco do mais barato para o mais caro
    resultado_df = resultado_df.reset_index(drop=True) #organizando index
    resultado_df.index = resultado_df.index + 1

    nome_arquivo = '_'.join(nome_busca.split(' ')[:3])
    resultado_df.to_csv(f'resultado_{nome_arquivo}.csv', sep=';',encoding='utf-8')
    resultado_html = resultado_df.to_html()

    qtd_resultado = len(resultado_df)
    qtd_resultado = f'<p style="font-size:20px; color:blue; font-weight: bold">{qtd_resultado} Resultados </p>'
    assunto = f'Resultado de Pesquisa {i+1} - {nome_arquivo.upper().replace("_", " ")}'
    
    #enviando email
    enviar_mail(remetente, senhaapp, remetente, assunto, (qtd_resultado + resultado_html), arquivos=[f'resultado_{nome_arquivo}.csv'], msg_html=True)
    
    

nav.close()


messagebox.showinfo("Mensagem", "Pesquisa conlcuída e email enviado com sucesso")