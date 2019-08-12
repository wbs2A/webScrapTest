import webFunctions


def run(url='', anoinicio=2003, anofim=2018):
    #webFunctions.getInformations(url, anoinicio, anofim)
    webFunctions.createTable()
    """
    retrys = 0
    while True:
        try:
            
            print('entrei')
            raise BaseException("Deu ruim, amigo")
        except Exception as err:
            webFunctions.sendEmail(err)
            break
            #if retrys == 3:
            #else:
            #   retrys+=1
            #    continue
    """


if __name__ == '__main__':
    run(url='https://catalogodeteses.capes.gov.br/catalogo-teses/rest/busca')