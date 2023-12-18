from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver import ActionChains
import time
import os
import pandas as pd


#pegando a pasta
directory = os.getcwd()

#caminho do chrome drive na minha maquina
chrome_driver = r'C:\Program Files\chromedriver_win32\chromedriver.exe'

#Opções do navegador Chrome
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.default_directory': directory,
    'download.prompt_for_downloads':True,
    'download.directory_upgrade':True,
    'safebrowsing.enabled': True
})


#criando o navegador
driver = webdriver.Chrome(executable_path=chrome_driver, options=options)


processos = pd.read_excel('Processos.xlsx')


#procurar cidade
def search_city(city):
    try:
        while len(driver.find_elements(By.XPATH , '/html/body/div/div')) < 1:
            time.sleep(1)
            print('1')
        
        menu = driver.find_element(By.XPATH , '/html/body/div/div')
        ActionChains(driver).move_to_element(menu).perform()

        options = menu.find_elements(By.TAG_NAME, 'a')

        for option in options:
            if city.lower() in option.text.lower():
                option.click()

    except Exception as er:
        print(er)

#preencher dados
def fill_in_data(name,attorney,process_number):
    try:
        new_page = driver.window_handles[1]
        driver.switch_to.window(new_page)

        driver.find_element(By.ID, 'nome').send_keys(str(name))
        driver.find_element(By.ID, 'advogado').send_keys(str(attorney))
        driver.find_element(By.ID, 'numero').send_keys(str(process_number))

        time.sleep(1)
        driver.find_element(By.CLASS_NAME, 'registerbtn').click()


    except Exception as er:
        print(er)


def box_dialogue():
    try:
        alerta = Alert(driver)
        time.sleep(1)
        alerta.accept()

        time.sleep(5)

    except Exception as er:
        print(er)



def main():
   
    #entrando em um pagina local / ou da web
    driver.get(directory + r'/index.html')

    #Maximizar a janela
    driver.maximize_window()



    for i in range(len(processos)):
        time.sleep(0.5)
        main_page = driver.window_handles[0]
        driver.switch_to.window(main_page)

        row = processos.iloc[i]

        name = row['Nome']
        attorney = row['Advogado']
        process_number = row['Processo']
        cities = row['Cidade']


        search_city(cities)

        fill_in_data(name, attorney, process_number)

        box_dialogue()

        alerta = Alert(driver)
        text_alert = alerta.text
        print(text_alert)

        time.sleep(1)
        alerta.accept()

        processos.at[i, 'Status'] = str(text_alert)
        processos.to_excel('Processos.xlsx', index=False, engine='openpyxl')

        time.sleep(1)
        driver.close()

      
   
    driver.quit()





if __name__ == '__main__':
    main()