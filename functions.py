import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

### variável global para o driver ###
driver = None

### função para garantir que a página está completamente carregada ###
def esperar_pagina_carregar():
    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

### função para formatar a data corretamente ###
def formatar_data(data):
    ### verifica se a data é válida e a formata para o padrão 'dd/mm/yyyy hh:mm' ###
    return pd.to_datetime(data).strftime("%d/%m/%Y %H:%M")

### função para iniciar o processo de preenchimento do formulário ###
def preencher_formulario():
    global df
    ### caminho para o arquivo excel ###
    planilha_path = 'C:\\Users\\kauantavares\\Downloads\\Teste.xlsx'
    df = pd.read_excel(planilha_path, header=None)  ### carrega os dados da planilha sem cabeçalhos ###

    ### definir manualmente os nomes das colunas com base nos 17 campos fornecidos ###
    df.columns = [
        'DATA PROVISÃO', 'CLIENTE', 'UND NEGÓCIO', 'I.C', 'TIPO DOC', 'NÚM DOC',
        'CLASSIFICAÇÃO', 'RECEITA BRUTA', 'ICMS', 'ISS', 'PIS', 
        'COFINS', 'CPRB', 'RECEITA LÍQUIDA', 'OBSERVAÇÃO', 'STATUS', 'LINK'
    ]

    ### preencher o formulário com os dados da planilha ###
    for index, row in df.iterrows():
        try:
            ### verificar se a 'DATA PROVISÃO' tem valor ###
            if pd.isna(row['DATA PROVISÃO']):
                print(f"linha {index} está sem valor em 'data provisão'. pulando...")  ### se não houver valor, pula a linha ###
                continue

            ### preencher "data provisão" no formato correto ###
            data_provisao_formatada = formatar_data(row['DATA PROVISÃO'])
            data_provisao_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'sp-DateTimePicker-DateTimeTextField')))
            data_provisao_input.clear()
            data_provisao_input.send_keys(data_provisao_formatada)

            ### preencher "cliente" ###
            cliente_input = driver.find_element(By.ID, 'TextField10')
            cliente_input.clear()
            cliente_input.send_keys(str(row['CLIENTE']))

            ### preencher "und negócio" ###
            und_negocio_input = driver.find_element(By.ID, 'TextField16')
            und_negocio_input.clear()
            und_negocio_input.send_keys(str(row['UND NEGÓCIO']))

            ### preencher "i.c" ###
            ic_input = driver.find_element(By.ID, 'TextField22')
            ic_input.clear()
            ic_input.send_keys(str(row['I.C']))

            ### preencher "tipo doc" ###
            tipo_doc_input = driver.find_element(By.ID, 'TextField28')
            tipo_doc_input.clear()
            tipo_doc_input.send_keys(str(row['TIPO DOC']))

            ### preencher "número doc" ###
            num_doc_input = driver.find_element(By.ID, 'TextField34')
            num_doc_input.clear()
            num_doc_input.send_keys(str(row['NÚM DOC']))

            ### preencher "classificação" ###
            classificacao_input = driver.find_element(By.ID, 'TextField40')
            classificacao_input.clear()
            classificacao_input.send_keys(str(row['CLASSIFICAÇÃO']))

            ### preencher "receita bruta" ###
            receita_bruta_input = driver.find_element(By.ID, 'TextField46')
            receita_bruta_input.clear()
            receita_bruta_input.send_keys(str(row['RECEITA BRUTA']))

            ### preencher "icms" ###
            icms_input = driver.find_element(By.ID, 'TextField52')
            icms_input.clear()
            icms_input.send_keys(str(row['ICMS']))

            ### preencher "iss" ###
            iss_input = driver.find_element(By.ID, 'TextField58')
            iss_input.clear()
            iss_input.send_keys(str(row['ISS']))

            ### preencher "pis" ###
            pis_input = driver.find_element(By.ID, 'TextField64')
            pis_input.clear()
            pis_input.send_keys(str(row['PIS']))

            ### preencher "cofins" ###
            cofins_input = driver.find_element(By.ID, 'TextField70')
            cofins_input.clear()
            cofins_input.send_keys(str(row['COFINS']))

            ### preencher "cprb" ###
            cprb_input = driver.find_element(By.ID, 'TextField76')
            cprb_input.clear()
            cprb_input.send_keys(str(row['CPRB']))

            ### preencher "receita líquida" ###
            receita_liquida_input = driver.find_element(By.ID, 'TextField82')
            receita_liquida_input.clear()
            receita_liquida_input.send_keys(str(row['RECEITA LÍQUIDA']))

            ### preencher "observação" ###
            observacao_input = driver.find_element(By.ID, 'TextField88')
            observacao_input.clear()
            observacao_input.send_keys(str(row['OBSERVAÇÃO']))

            ### preencher "status" ###
            status_input = driver.find_element(By.ID, 'TextField94')
            status_input.clear()
            status_input.send_keys(str(row['STATUS']))

            ### clicar no botão "enviar" ###
            botao_enviar = driver.find_element(By.ID, 'id__0')
            botao_enviar.click()

            time.sleep(2)  ### aguarda 2 segundos antes de processar a próxima linha ###

            ### clicar em "enviar outra resposta" ###
            botao_enviar_outra = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.submitAnotherResponseButton_77b07534')))
            botao_enviar_outra.click()
            esperar_pagina_carregar()

        except Exception as e:
            print(f"erro ao processar a linha {index}: {e}")  ### exibe erro ao processar a linha ###

### função para abrir o driver e acessar o link ###
def abrir_driver():
    global driver
    driver = webdriver.Chrome()  ### inicia o driver do chrome ###

    driver.get('https://expressonepomucenobr.sharepoint.com/sites/Custos-DPO/_layouts/15/listforms.aspx?cid=MDEyZGFhNmYtOGFlYS00OGU2LWExYTQtY2ZlZTBiMjcyZDk4&nav=ZDVlNDQ3M2YtMzAxNi00Y2E2LWFjNGItNDU4NjhlYmE4MzY1')
    esperar_pagina_carregar()  ### espera a página carregar completamente ###
