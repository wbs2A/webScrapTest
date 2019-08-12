import os, time
import requests,math
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from xlwt import *
import xlrd

#email
import smtplib


def doSearch(driver):
    #Processo de captura das informações
    search = driver.find_element_by_xpath("//input[@ng-model='termo']")
    search.send_keys("bolsa família", Keys.ENTER)

    #Aguarda os elementos serem renderizados
    time.sleep(10)
    wait = WebDriverWait(driver, 10)
    elm = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@ng-repeat='td in resultado.tesesDissertacoes']")))


def changePage(driver, page):
    print('entrei no change', page)
    #pega os links da lista
    elements = driver.find_elements_by_xpath("//a[@ng-click='selectPage(page.number, $event)']")

    #procura a página desejada
    for el in elements:
        if str(el.get_attribute('innerHTML')) == str(page):
            print("Achei o elemento aqui:",el.get_attribute('innerHTML'),el.get_attribute('outerHTML'))
            el.click()
            time.sleep(15)
            return


def getOnPage(driver, page):

    #Conferindo se estamos na página desejada
    activePageLiTag = driver.find_element_by_class_name('active') #elemento da lista de paginação que está ativo
    activePageLink = activePageLiTag.find_element_by_class_name('ng-binding')
    assert str(page) in str(activePageLink.get_attribute('innerHTML'))

    #pega todas as divs de resultado da pesquisa
    elements = driver.find_elements_by_xpath("//div[@ng-repeat='td in resultado.tesesDissertacoes']")
    noLinks = []

   #Separa os trabalhos com e sem links
    with open(os.path.join(os.getcwd()+'\\tmp\\','links.txt'), 'a', encoding='utf-8') as file:
        for el in elements:
            if "Trabalho anterior à Plataforma Sucupira" in str(el.text):
                noLinks.append(str(el.text))
            else:
                link = el.find_element_by_tag_name('a')
                file.write(str(link.get_attribute('href')))
                file.write('\n\n')
    #insere os trabalhos sem links noutro arquivo
    with open(os.path.join(os.getcwd()+'\\tmp\\','semlinks.txt'), 'a', encoding='utf-8') as file:
        for work in noLinks:
            file.write(work)
            file.write('\n')

    #muda para a página seguinte
    changePage(driver, page+1)


def extractInfoFromLink(url):
    print('entrei no extract')
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    options.add_argument('window-size=1920x1080')

    driver = webdriver.Chrome(options=options,
                              executable_path='C:\\Users\\Betelgeuse\\Downloads\\chromedriver_win32\\chromedriver.exe')
    driver.get(url)
    time.sleep(.5)

    print(driver.find_element_by_id('autor').get_attribute('innerHTML'))

    data = {'aluno': driver.find_element_by_id('autor').get_attribute('innerHTML'),
            'nome_tese': driver.find_element_by_id('nome').get_attribute('innerHTML'),
            'universidade': driver.find_element_by_id('ies').get_attribute('innerHTML'),
            'orientador': driver.find_element_by_id('orientador').get_attribute('innerHTML'),
            'area': driver.find_element_by_id('area').get_attribute('innerHTML'),
            'resumo': driver.find_element_by_id('resumo').get_attribute('innerHTML')
            }

    driver.close()
    return data


def createTable():
    w = Workbook()
    ws = w.add_sheet('Tabela de Teses_Dissertações')

    ws.write(0, 0, 'Autor')
    ws.write(0, 1, 'Nome da Tese')
    ws.write(0, 2, 'Universidade')
    ws.write(0, 3, 'Orientador')
    ws.write(0, 4, 'Área de concentração')
    ws.write(0, 5, 'Resumo')
    ws.write(0, 6, "Link")
    try:
        saveOnTable(ws)
        w.save('teses_dissertacoes_capes.xls')
    except Exception as e:
        print(e)
        w.save('teses_dissertacoes_capes.xls')


def extractLink(response):
    # Separa os trabalhos com e sem links
    noLinks = []
    with open(os.path.join(os.getcwd() + '\\tmp\\', 'links.txt'), 'a', encoding='utf-8') as file:
        for el in response['tesesDissertacoes']:
            if not el['link']:
                noLinks.append(str(el))
            else:
                print(el['link'])
                link = el['link']
                file.write(link)
                file.write('\n\n')
    # insere os trabalhos sem links noutro arquivo
    with open(os.path.join(os.getcwd() + '\\tmp\\', 'semlinks.txt'), 'a', encoding='utf-8') as file:
        for work in noLinks:
            file.write(work)
            file.write('\n')


def saveOnTable(worksheet):
    i = 1
    with open(os.path.join(os.getcwd() + '\\tmp\\', 'links.txt'), 'r') as file:
        lines = [line.rstrip('\n') for line in file]
        lines = filter(lambda a: a!= '', lines)
        for line in lines:
            try:
                info = extractInfoFromLink(line)
                with open('concluidos.txt', 'a') as concluidos:
                    concluidos.write(line)
                    concluidos.write('\n')
            except:
                continue
            finally:
                worksheet.write(i, 0, info['aluno'])
                worksheet.write(i, 1, info['nome_tese'])
                worksheet.write(i, 2, info['universidade'])
                worksheet.write(i, 3, info['orientador'])
                worksheet.write(i, 4, info['area'])
                worksheet.write(i, 5, info['resumo'])
                worksheet.write(i, 6, line)
                i += 1


def getInformations(url, anoinicio, anofim):
    payload = "{\"termo\":\"bolsa famÃ­lia\",\"pagina\":1,\"filtros\":[{\"campo\":\"Ano\",\"valor\":\"2003\"}],\"registrosPorPagina\":20}"
    headers = {
        'accept': "application/json, text/plain, */*",
        'user-agent': "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.87 Safari/537.36",
        'sec-fetch-mode': "cors",
        'content-type': "application/json;charset=UTF-8",
        'cache-control': "no-cache",
    }
    i, j = anoinicio, 1
    # ano a ano
    while i<=anofim:
        try:
            response = requests.request("POST", url, data=payload, headers=headers)
            resp_data = eval(response.text)
        except:
            print(i,i-1)
            payload = payload.replace(str(i), str(i - 1), 1)
            i-=1
            continue
        else:
            quantpaginas = math.ceil(resp_data['total'] / resp_data['registrosPorPagina'])
            print(resp_data, quantpaginas)
            # página a página
            j = 1
            while j <= quantpaginas:
                try:
                    response = requests.request("POST", url, data=payload, headers=headers)
                    resp_data = eval(response.text)
                except:
                    print(j)
                    payload = payload.replace(str(j), str(j - 1), 1)
                    j-=1
                    print(payload, j)
                    continue
                else:
                    extractLink(resp_data)
                    time.sleep(.5)
                    payload = payload.replace(str(j), str(j + 1), 1)
                    print(j,payload)
                    j+=1
            # volta para a página 1
            payload = payload.replace(str(j), str(1), 1)
            # muda para o ano seguinte
            payload = payload.replace(str(i), str(i + 1), 1)
            print(payload)
            i+=1

def sendEmail(text):
    smtp = smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465)
    smtp.login("barbosawesley101@yahoo.com", "Eli@ne9275")
    de = "barbosawesley101@yahoo.com"
    para = "barbosawesley101@gmail.com"
    msg = """
    From: % s
    To: % s
    Subject: Deu ruim
    Email
    de
    teste
    do
    SempreUpdate.""".format(de, para)
    smtp.sendmail(de, para, msg)
    smtp.quit()