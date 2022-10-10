from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook
import time

navegador = webdriver.Chrome()
navegador.get(
    "https://api.manheim.com/auth/authorization.oauth2?adaptor=manheim_customer&client_id=6rxsc2as8cmntj5fd44w2db3&redirect_uri=https://2ndchance.manheim.com/oauthcallback&response_type=code")
input_1 = navegador.find_element(by=By.ID, value='user_username')
input_1.send_keys('mperroni')
input_2 = navegador.find_element(by=By.ID, value='user_password')
input_2.send_keys('mperroni')
navegador.find_element("xpath", '//*[@id="submit"]').click()

dt_select = navegador.find_element(by=By.NAME, value='ddSite')
aux_dt_select = []
aux_dt_select.append(dt_select.text.split('\n'))
aux_dt_select[0].remove(' IL - Manheim Arena Illinois')
flat_list_select = [item for sublist in aux_dt_select for item in sublist]

aux = []
aux2 = []
lista_aux = []

df = pd.DataFrame(columns=["Cidade", "Especificações", "Oferta"])
writer = pd.ExcelWriter('veiculos.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='veiculos', index=False)
writer.save()

k = -1
for j in flat_list_select:
    select_1 = navegador.find_element(by=By.ID, value='ddSite')
    select_1.send_keys(j)
    navegador.find_element(by=By.CLASS_NAME, value='selection').click()
    size2 = []
    k = k+1
    time.sleep(1)
    try:
        #print(size2)
        size2 = navegador.find_element(by=By.CLASS_NAME, value='pagination_info')
        index_tam = size2.text.split()
        tam_total = int(index_tam[2])
    except:
        #print('Lista vazia!!!')
        pass
    else:
        calculo_pagina = tam_total / 50
        aux_calculo = int(calculo_pagina)
        i = 0

        #if aux_dt_select[0][k] == ' LA - Manheim New Orleans':
        #    k = k+1

        while i != aux_calculo + 1:

            dt2 = navegador.find_element(by=By.ID, value='sc_data')
            aux.append(dt2.text.split('\n'))

            teste2 = [aux[0][i:i + 3] for i in range(0, len(aux[0]), 3)]

            for element2 in navegador.find_elements(By.NAME, 'offer'):
                aux2.append(element2.get_attribute('value'))

            aux3 = []

            for element in aux_dt_select:
                aux3.append(element[k])

            teste3 = [aux3[0] for i in range(0, len(aux[0]), 1)]
            #print(teste3)

            #teste3 = []
            #teste3.append(j)

            data_tuples2 = list(zip(teste3, teste2, aux2))

            myFileName = 'veiculos.xlsx'

            wb = load_workbook(filename=myFileName)
            ws = wb['veiculos']

            lista_aux2 = []
            for item in data_tuples2:
                lista_aux2.append([''.join(map(str, x)) for x in item])

            tuples = []
            for x in lista_aux2:
                tuples.append(tuple(x))

            for row in tuples:
                ws.append(row)

            wb.save(filename=myFileName)

            aux.clear()
            aux2.clear()

            if i == 0 and tam_total > 50:
                navegador.find_element("xpath", '//*[@id="pagination_info"]/a').click()
            elif i < aux_calculo and tam_total > 50:
                navegador.find_element("xpath", '//*[@id="pagination_info"]/a[2]').click()

            i = i + 1
        index_tam.clear()

    if j == 76:
        wb.close()
        break
print('Programa Finalizado com sucesso!!!')