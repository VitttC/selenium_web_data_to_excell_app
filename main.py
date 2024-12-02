from selenium.webdriver import Chrome
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import pandas as pd
import xlsxwriter
from pathlib import Path
from pandas import DataFrame



# RICHIESTA GIORNO DA VERIFICARE
datadal = input('\ninserire data di affidamento: gg/mm/aaaa\n\n\t')
committ = input('inserire committente:\n\t')

# ASSEGNAZIONE VARIABILI
sito = 'https://link.to.site/'
bolle = 'https://path/to/resource/'
username = 'vittorio'
password = 'password'


# INSTALLAZIONE DRIVER CHROME
chrome_driver = ChromeDriverManager().install()
driver = Chrome(service=Service(chrome_driver))
driver.maximize_window()


# VAI AL SITO
driver.get(sito)


# LOGIN
usrweb = driver.find_element(By.NAME, 'Utente')
usrweb.send_keys(username)
sleep(0.2)


#se non si deve conservare il processo su una variabile si può concatenare:
driver.find_element(By.NAME, 'password').send_keys(password+'\n') #note: carattere a capo in sostituzione del tasto invio
sleep(0.2)


# PASSAGGIO A SCHEDA BOLLE
driver.find_element(By.XPATH, '//td /a [text()="Bolle"]').click()


# INSERIMENTO DATA DAL/AL
driver.find_element(By.NAME, 'DataBollaDal').send_keys(datadal)
driver.find_element(By.NAME, 'DataBollaAL').send_keys(datadal)
sleep(0.3)


# INSERIMENTO ANAGRAFICA CLIENTE
driver.find_element(By.NAME, 'txtAnagraficaCliente').send_keys(committ +'\n')
sleep(0.5)


### VERIFICA PRESENZA BOLLE MAGGIORE O MINORE DI DIECI ###


# FUNZIONE DI VERIFICA: se falsa ci sono dieci o meno bolle
def check_exists(xpath):
    try:
        driver.find_element(By.XPATH, xpath)
    except NoSuchElementException:
        return False
    return True

lista_dati = []


# PATH per minore di dieci /html/body/div[2]/center/table[4]/tbody/tr/td[2]/span
bollemagg= '/html/body/div[2]/center/table[4]/tbody/tr/td[2]/form/span'
flag = True
check = check_exists(bollemagg)

if check == True:

    #RECUPERO NUMERO BOLLE MAGGIORE DI DIECI
    bolletot = driver.find_element(By.XPATH, '/html/body/div[2]/center/table[4]/tbody/tr/td[2]/form/span').text #STRINGA NUMERO TOTALE BOLLLE
    bolletot = bolletot.rpartition(' ') #DIVIDI AL PRIMO SPAZIO DA DESTRA
    bolletot = int(bolletot[2]) #RECUPERO NUMERO TOTALE INT
    bolletot = str(bolletot) #CONVERSIONE IN STRINGA
    deci = int(bolletot[0]) #ESTRAZIONE DECINA DA STRINGA
    single = int(bolletot[1]) #ESTRAZIONE NUMERO DI BOLLE NELL'ULTIMO CICLO

    #DECINE = N PAGINE
    for decine in range(0,deci):
        rowmax = 13        
        for el in range(3,rowmax):
            rigaX = el
            classname_riga = 'f1'

            #DEFINIZIONE PERCORSI
            path_I_riga = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']'
            path_data_aff = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[3]'
            path_rif_mitt = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[4]'
            path_destinatario = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[8]'
            path_destinazione = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[9]'
            path_colli = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[10]'
            path_kg = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[17]'
            
            
            #RECUPERO DATI DA INSERIRE
            data_aff = driver.find_element(By.XPATH, path_data_aff).text
            data_aff = data_aff.partition(' ')
            rif_mitt = driver.find_element(By.XPATH, path_rif_mitt).text
            destario = driver.find_element(By.XPATH, path_destinatario).text
            destaone = driver.find_element(By.XPATH, path_destinazione).text
            colli = driver.find_element(By.XPATH, path_colli).text
            peso = driver.find_element(By.XPATH, path_kg).text


            #ESTRAZIONE CITTÀ E PROVINCIA
            citta = destaone.partition(' ')
            prov_str = destaone.replace(')', '(').split('(',4)
            provincia = prov_str[1]


            #INSERIMENTO DATI IN DICT
            oggetto_bolle = {
            'riferimento' : rif_mitt,
            'città' : citta[0],
            'provincia' : provincia,
            'destinatario' : destario,
            'colli' : colli,
            'peso' : peso,
            'affidamento' : data_aff[2]
                }
	    
	    #AGGIUNGI DICT ALLA LISTA
            lista_dati.append(oggetto_bolle)
            
            
    #VAI ALLA PAGINA SUCCESSIVA
    driver.find_element(By.NAME, 'SoloCodiceBolla').send_keys(Keys.ARROW_RIGHT)
            
        
    #CICLO PER ULTIMA PAGINA
    rowmax = single+3

    for el in range(3,rowmax):
        rigaX = el
        classname_riga = 'f1'

        #DEFINIZIONE PERCORSI

        path_I_riga = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']'
        path_data_aff = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[3]'
        path_rif_mitt = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[4]'
        path_destinatario = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[8]'
        path_destinazione = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[9]'
        path_colli = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[10]'
        path_kg = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[17]'
        
        
        #RECUPERO DATI DA INSERIRE
        data_aff = driver.find_element(By.XPATH, path_data_aff).text
        data_aff = data_aff.partition(' ')
        rif_mitt = driver.find_element(By.XPATH, path_rif_mitt).text
        destario = driver.find_element(By.XPATH, path_destinatario).text
        destaone = driver.find_element(By.XPATH, path_destinazione).text
        colli = driver.find_element(By.XPATH, path_colli).text
        peso = driver.find_element(By.XPATH, path_kg).text

        #ESTRAZIONE CITTÀ E PROVINCIA
        citta = destaone.partition(' ')
        prov_str = destaone.replace(')', '(').split('(',4)
        provincia = prov_str[1]

        #INSERIMENTO DATI IN DICT
        oggetto_bolle = {
            'riferimento' : rif_mitt,
            'città' : citta[0],
            'provincia' : provincia,
            'destinatario' : destario,
            'colli' : colli,
            'peso' : peso,
            'affidamento' : data_aff[2]
            }

        lista_dati.append(oggetto_bolle)

        
        
if check == False:

    # RECUPERO NUM BOLLE MINORE DI DIECI

    bolletot = driver.find_element(By.XPATH, '/html/body/div[2]/center/table[4]/tbody/tr/td[2]/span').text
    bolletot = bolletot.rpartition(' ')
    bolletot = int(bolletot[2])
    
    if bolletot >= 10:
        bolletot = str(bolletot)
        bolletot = int(bolletot[1])
    
    rowmax = bolletot+3


    for el in range(3,rowmax):
        rigaX = el
        classname_riga = 'f1'

        #DEFINIZIONE PERCORSI
        path_I_riga = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']'
        path_data_aff = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[3]'
        path_rif_mitt = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[4]'
        path_destinatario = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[8]'
        path_destinazione = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[9]'
        path_colli = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[10]'
        path_kg = '/html/body/div[2]/center/table[2]/tbody/tr['+str(rigaX)+']/td[17]'
        
        
        #RECUPERO DATI DA INSERIRE
        data_aff = driver.find_element(By.XPATH, path_data_aff).text
        data_aff = data_aff.partition(' ')
        rif_mitt = driver.find_element(By.XPATH, path_rif_mitt).text
        destario = driver.find_element(By.XPATH, path_destinatario).text
        destaone = driver.find_element(By.XPATH, path_destinazione).text
        colli = driver.find_element(By.XPATH, path_colli).text
        peso = driver.find_element(By.XPATH, path_kg).text


        #ESTRAZIONE CITTÀ E PROVINCIA
        citta = destaone.partition(' ')
        prov_str = destaone.replace(')', '(').split('(',4)
        provincia = prov_str[1]


        #INSERIMENTO DATI IN DICT
        oggetto_bolle = {
            
            'riferimento' : rif_mitt,
            'città' : citta[0],
            'provincia' : provincia,
            'destinatario' : destario,
            'colli' : colli,
            'peso' : peso,
            'affidamento' : data_aff[2]
            }

        lista_dati.append(oggetto_bolle)
       


driver.quit()


#STAMPA DEL DATAFRAME A VIDEO
df = pd.DataFrame(lista_dati)
print(df,flush=True)


#SALVATAGGIO EXCEL
output = input("nome file excel:\n\n\t")

df.to_excel(output+'.xlsx')



