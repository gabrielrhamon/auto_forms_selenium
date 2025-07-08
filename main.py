from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager as CDM
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas
import random
import time 

# G L O B A L

#W E B D R I V E R
def Initialize_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("--headless")
    return webdriver.Chrome(service=Service(CDM().install()), options=options)
#D A D O S   E X C E L
def load_dados_excel(excel_file, header=0):
    return pandas.read_excel(excel_file, header=header, dtype={'CIC': str}, converters= {'CIC': lambda x: str(x).strip().zfill(11)})
#C L I C K E R
def click_button(driver, xpath, timeout=600):
    element = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    WebDriverWait(driver,timeout).until(EC.visibility_of(element))
    element.click()
#W R I T E Rcd
def write_text(driver, xpath, text, timeout=600):
    element_writer = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
    element_writer.send_keys(Keys.CONTROL + "a")
    element_writer.send_keys(Keys.BACKSPACE)
    element_writer.send_keys(text)
#M A I N
def main():

    df = load_dados_excel(excel_file = "x.xlsx", header=0).sort_values(by="Random") # Carrega o arquivo Excel e ordena por coluna "Random" - Nome do arquivo removido por segurança
    driver = Initialize_driver()
    #   V   A   R   -----------------------------------------------------------------------------------------    F O R M S
    
    
    filial_xpaths = {
        "CBMSA Mina": '/html/body/div[2]/div/div[3]/span[2]/span'
    }
    função_xpaths ={
        "ÔNIBUS": '//*[@id="question-list"]/div[5]/div[2]/div/div/div[1]/div'
    }
    resid_xparths ={
        "'SIM - EC II": '//*[@id="question-list"]/div[5]/div[2]/div/div/div[2]',
        "CIDADE":       '//*[@id="question-list"]/div[5]/div[2]/div/div/div[1]'
    }
    resid_casa_xpaths ={
        "PRIMAVERA":'//*[@id="question-list"]/div[7]/div[2]/div/div/div[1]/div',
        "GUARÁ":'//*[@id="question-list"]/div[7]/div[2]/div/div/div[2]/div',
        "JANDAIÁ":'//*[@id="question-list"]/div[7]/div[2]/div/div/div[3]/div',
        "ANDORINHA":'//*[@id="question-list"]/div[7]/div[2]/div/div/div[4]/div'
    }
    casa_bloco_xpaths = {
        "1A": '/html/body/div[2]/div/div[1]',
        "1B": '/html/body/div[2]/div/div[2]',
        "1C": '/html/body/div[2]/div/div[3]',
        "1D": '/html/body/div[2]/div/div[4]',
        "1E": '/html/body/div[2]/div/div[5]',
        "1F": '/html/body/div[2]/div/div[6]',
        "1G": '/html/body/div[2]/div/div[7]',
        "2A": '/html/body/div[2]/div/div[8]',
        "2B": '/html/body/div[2]/div/div[9]',
        "2C": '/html/body/div[2]/div/div[10]',
        "2D": '/html/body/div[2]/div/div[11]',
        "2G": '/html/body/div[2]/div/div[12]',
        "3B": '/html/body/div[2]/div/div[13]',
        "3C": '/html/body/div[2]/div/div[14]',
        "3D": '/html/body/div[2]/div/div[15]',
        "4A": '/html/body/div[2]/div/div[16]',
        "4B": '/html/body/div[2]/div/div[17]',
        "4C": '/html/body/div[2]/div/div[18]',
        "4D": '/html/body/div[2]/div/div[19]',
        "5A": '/html/body/div[2]/div/div[20]',
        "5B": '/html/body/div[2]/div/div[21]',
        "5C": '/html/body/div[2]/div/div[22]',
        "5D": '/html/body/div[2]/div/div[23]',
        "6A": '/html/body/div[2]/div/div[24]',
        "6B": '/html/body/div[2]/div/div[25]',
        "6C": '/html/body/div[2]/div/div[26]',
        "6D": '/html/body/div[2]/div/div[27]',
        "7A": '/html/body/div[2]/div/div[28]',
        "7B": '/html/body/div[2]/div/div[29]',
        "7C": '/html/body/div[2]/div/div[30]',
        "7D": '/html/body/div[2]/div/div[31]',
        "OUTRO":'/html/body/div[2]/div/div[32]'
    }
    dist_via_xpaths = {
        "Até 100km (por ex: Parauapebas)":'//*[@id="question-list"]/div[13]/div[2]/div/div/div[1]/div',
        "Entre 100 e 300km (por ex: Marabá)":'//*[@id="question-list"]/div[13]/div[2]/div/div/div[2]/div',
        "Entre 300 e 500km (por ex: Tucuruí)":'//*[@id="question-list"]/div[13]/div[2]/div/div/div[3]/div',
        "Acima de 500km":'//*[@id="question-list"]/div[13]/div[2]/div/div/div[4]/div'
        }
    dist_via_xpathsb = {
        "Até 100km (por ex: Parauapebas)":'//*[@id="question-list"]/div[11]/div[2]/div/div/div[1]/div',
        "Entre 100 e 300km (por ex: Marabá)":'//*[@id="question-list"]/div[11]/div[2]/div/div/div[2]/div',
        "Entre 300 e 500km (por ex: Tucuruí)":'//*[@id="question-list"]/div[11]/div[2]/div/div/div[3]/div',
        "Acima de 500km":'//*[@id="question-list"]/div[11]/div[2]/div/div/div[4]/div'
        }
    inic_hora_xpaths = {
        7:'/html/body/div[2]/div/div[7]',
        16:'/html/body/div[2]/div/div[16]'}
    inic_min_xpaths = {
        0:'/html/body/div[2]/div/div[1]',
        5:'/html/body/div[2]/div/div[2]',
        10:'/html/body/div[2]/div/div[3]'
    }
    horas_cama_e_acordou_xpaths ={
        1:'/html/body/div[2]/div/div[1]',
        2:'/html/body/div[2]/div/div[2]',
        3:'/html/body/div[2]/div/div[3]',
        4:'/html/body/div[2]/div/div[4]',
        5:'/html/body/div[2]/div/div[5]',
        6:'/html/body/div[2]/div/div[6]',
        7:'/html/body/div[2]/div/div[7]',
        8:'/html/body/div[2]/div/div[8]',
        9:'/html/body/div[2]/div/div[9]',
        10:'/html/body/div[2]/div/div[10]',
        11:'/html/body/div[2]/div/div[11]',
        12:'/html/body/div[2]/div/div[12]',
        13:'/html/body/div[2]/div/div[13]',
        14:'/html/body/div[2]/div/div[14]',
        15:'/html/body/div[2]/div/div[15]',
        16:'/html/body/div[2]/div/div[16]',
        17:'/html/body/div[2]/div/div[17]',
        18:'/html/body/div[2]/div/div[18]',
        19:'/html/body/div[2]/div/div[19]',
        20:'/html/body/div[2]/div/div[20]',
        21:'/html/body/div[2]/div/div[21]',
        22:'/html/body/div[2]/div/div[22]',
        23:'/html/body/div[2]/div/div[23]',
        24:'/html/body/div[2]/div/div[24]'
    }
    horas_sono_xpaths={
        "Menos de 4":'/html/body/div[2]/div/div[4]',
        5:'/html/body/div[2]/div/div[5]',
        6:'/html/body/div[2]/div/div[6]',
        7:'/html/body/div[2]/div/div[7]',
        8:'/html/body/div[2]/div/div[8]',
        9:'/html/body/div[2]/div/div[9]',
        10:'/html/body/div[2]/div/div[10]',
        11:'/html/body/div[2]/div/div[11]',
        12:'/html/body/div[2]/div/div[12]',
        "Acima de 12":'/html/body/div[2]/div/div[13]'
    }    
    horas_cama_min_xpaths = {
        15:'//*[@id="question-list"]/div[3]/div[2]/div/div/div[1]/div',
        30:'//*[@id="question-list"]/div[3]/div[2]/div/div/div[2]/div',
        45:'//*[@id="question-list"]/div[3]/div[2]/div/div/div[3]/div'

    }
    demorou_para_dormir_xpaths = {
        15:'//*[@id="question-list"]/div[3]/div[2]/div/div/div[1]/div',
        30:'//*[@id="question-list"]/div[3]/div[2]/div/div/div[2]/div',
        45:'//*[@id="question-list"]/div[3]/div[2]/div/div/div[3]/div',
        60:'//*[@id="question-list"]/div[3]/div[2]/div/div/div[4]/div',
        "+ que 60 minutos":'//*[@id="question-list"]/div[3]/div[2]/div/div/div[5]/div'

    }
    acordou_trab_min_xpaths = {
        15:'/html/body/div[2]/div/div[4]',
        30:'/html/body/div[2]/div/div[7]',
        45:'/html/body/div[2]/div/div[10]'

    }
    esposa_parc_xpaths = {
        "Sim, parceiro ou colega, mas em outro quarto": '//*[@id="question-list"]/div[16]/div[2]/div/div/div[2]/div',
        "Sim, parceiro no mesmo quarto, mas em outra cama": '//*[@id="question-list"]/div[16]/div[2]/div/div/div[3]/div',
        "Sim, parceiro/a na mesma cama": '//*[@id="question-list"]/div[16]/div[2]/div/div/div[4]/div'
    }
    rota_casa_xpaths ={
        "1A" :'//*[@id="question-list"]/div[7]/div[2]/div/div/div[1]/div',
        "1B" :'//*[@id="question-list"]/div[7]/div[2]/div/div/div[2]/div',
        "2A" :'//*[@id="question-list"]/div[7]/div[2]/div/div/div[3]/div',
        "2B" :'//*[@id="question-list"]/div[7]/div[2]/div/div/div[4]/div',
        "3A" :'//*[@id="question-list"]/div[7]/div[2]/div/div/div[5]/div',
        "3B" :'//*[@id="question-list"]/div[7]/div[2]/div/div/div[6]/div',
        "4A" :'//*[@id="question-list"]/div[7]/div[2]/div/div/div[7]/div'
    }
    # A P L I C A T I O  N ---------------------------------------------------------------------------------- S T A R T   G E T
    try:
        driver.get("-") # URL real removida por segurança
        for index, row in df.iterrows():
            try:
                nome_colab = row['NOME']
                cic_colab = str(row['CIC']).zfill(11)
                fun_colab = row['RAC']
                equipamento_colab = row['EQUIPAMENTO']
                filial_colab = row['EMPRESA']
                resid_colab = row['RESIDÊNCIA']
                inic_hora_colab = row['INICIO(Hora)']
                inic_min_colab = row['INICIO(Min)']
                resid_casa_colab = row['QUAL CASA']
                casa_bloco_colab = row['QUAL SEU BLOCO?']
                quarto_alog_colab = row['ESCREVA QUAL QUARTO ESTÁ ALOJADO:']
                #rota_casa_colab = row ['QUAL ROTA']
                turno_colab = row['TURNO']
                #rota_peba_colab = row['ROTA PEBA']
                #folga_pag_colab = row['FOLGA PAGAMENTO']
                #dist_via_colab = row['QUAL A DISTÂNCIA DO SEU DESTINO?']
                horascama_colab = row['cama à noite dia? hora']
                horas_cama_min_colab = row['cama à noite dia? minuto']
                demorou_para_dormir_colab = row['demorou para dormir?']
                acordou_trab_colab = row['acordou para ir trabalhar? hora']
                acordou_trab_min_colab = row['acordou para ir trabalhar? minutos']
                horas_sono_colab = row['quantas horas de sono você teve por noite?']
                esposa_parc_colab = row ['Esposo(a) ou colega de quarto?']
                esposa_parc_atrap_colab = row ['quarto atrapalha o seu sono?']
                esposa_parc_atrap_desc_colab = row ['Descreva o motivo:']


                #P A G E    ----    0 1

                # Q1    ----    N O M E

                write_text(driver, '//*[@id="question-list"]/div[2]/div[2]/div/span/input', nome_colab)
                print(f"Nome: {nome_colab}")
                # Q2    ----    C I C

                write_text(driver, '//*[@id="question-list"]/div[3]/div[2]/div/span/input', cic_colab)

                print(f"CIC :{cic_colab}")
                # Q3    ----    E M P R E S A

                click_button(driver, '//*[@id="question-list"]/div[4]/div[2]/div/div/div/span[2]' )
                click_button(driver, '/html/body/div[2]/div/div[3]/span[2]/span' )
                
                
                # Q4    ----     A L O J A M E N T O

                if resid_colab == 'SIM - EC II':
                    click_button(driver, '//*[@id="question-list"]/div[5]/div[2]/div/div/div[2]/div')

                    # Q4.1    ----    Possui algum fato externo que interrompa seu sono? Não
                    
                    click_button(driver, '//*[@id="question-list"]/div[6]/div[2]/div/div/div[2]/div')
                    print(f"Residência: {resid_colab}")
                    # Q5    ----    T U R N O

                    if turno_colab == 'DIURNO':
                        click_button(driver, '//*[@id="question-list"]/div[7]/div[2]/div/div/div[2]/div')
                    elif turno_colab == 'NOTURNO':
                        click_button(driver, '//*[@id="question-list"]/div[7]/div[2]/div/div/div[1]/div')
                    print(f"Turno: {turno_colab}")

                    # Q6    ----    T U R N O   H O R A
                    
                    if inic_hora_colab in inic_hora_xpaths:
                        click_button(driver, '//*[@id="question-list"]/div[8]/div[2]/div/div/div/span[2]')
                        click_button(driver, inic_hora_xpaths[inic_hora_colab])
                    
                    # Q7    ----    T U R N O   M I N U T O

                    if inic_min_colab in inic_min_xpaths:
                        click_button(driver, '//*[@id="question-list"]/div[9]/div[2]/div/div/div/span[2]')
                        click_button(driver, inic_min_xpaths[inic_min_colab])
                    print(f"Hora de início: {inic_hora_colab}:{inic_min_colab}")
                    
                    # Q8    ----    R A C

                    if fun_colab == "RAC 02" and equipamento_colab =="Onibus":
                        click_button(driver, '//*[@id="question-list"]/div[10]/div[2]/div/div/div[2]/div')
                        
                    elif fun_colab == "RAC 02":
                        click_button(driver, '//*[@id="question-list"]/div[10]/div[2]/div/div/div[1]/div')
                        
                    elif fun_colab == "RAC 03":
                        click_button(driver, '//*[@id="question-list"]/div[10]/div[2]/div/div/div[3]/div')
                        
                    print(f"Função: {fun_colab}")
                
                # Q4    ----    C A S A    C A N A Ã
                elif resid_colab == 'CIDADE':
                    click_button(driver, '//*[@id="question-list"]/div[5]/div[2]/div/div/div[1]/div')
                    print(f"Residência: {resid_colab}")
                    
                    # Q5    ----    T U R N O

                    if turno_colab == 'DIURNO':
                        click_button(driver, '//*[@id="question-list"]/div[6]/div[2]/div/div/div[2]/div')
                    elif turno_colab == 'NOTURNO':
                        click_button(driver, '//*[@id="question-list"]/div[6]/div[2]/div/div/div[1]/div')
                    print(f"Turno: {turno_colab}")

                    # Q6    ----    T U R N O   H O R A

                    if inic_hora_colab in inic_hora_xpaths:
                        click_button(driver, '//*[@id="question-list"]/div[7]/div[2]/div/div/div/span[2]')
                        click_button(driver, inic_hora_xpaths[inic_hora_colab])
                    
                    # Q7    ----    T U R N O   M I N U T O

                    if inic_min_colab in inic_min_xpaths:
                        click_button(driver, '//*[@id="question-list"]/div[8]/div[2]/div/div/div/span[2]')
                        click_button(driver, inic_min_xpaths[inic_min_colab])
                    print(f"Hora de início: {inic_hora_colab}:{inic_min_colab}")
                    
                    # Q8    ----    R A C

                    if fun_colab == "RAC 02" and equipamento_colab =="Onibus":
                        click_button(driver, '//*[@id="question-list"]/div[9]/div[2]/div/div/div[2]/div')
                        
                    elif fun_colab == "RAC 02":
                        click_button(driver, '//*[@id="question-list"]/div[9]/div[2]/div/div/div[1]/div')
                        
                    elif fun_colab == "RAC 03":
                        click_button(driver, '//*[@id="question-list"]/div[9]/div[2]/div/div/div[3]/div')
                        
                    print(f"Função: {fun_colab}")

                # NEXT_PAGE
                click_button(driver,'//*[@id="form-main-content1"]/div/div/div[2]/div[3]/div')
                
                
                #P A G E    ----    0 2   -------------- -------------- -------------- -------------- -------------- 

                #Q10    ----    Durante o último mês, quando geralmente você foi para a cama à noite/ dia? (hora)
                
                if horascama_colab in horas_cama_e_acordou_xpaths:
                    click_button(driver, '//*[@id="question-list"]/div[2]/div[2]/div/div/div')                    
                    click_button(driver, horas_cama_e_acordou_xpaths[horascama_colab])
                print(f"Hora de cama: {horascama_colab}")
                
                #Q11    ----    Durante o último mês, quanto tempo em minutos você demorou para dormir?

                if horas_cama_min_colab in horas_cama_min_xpaths:
                    click_button(driver, horas_cama_min_xpaths[horas_cama_min_colab])
                print(f"Minutos para dormir: {horas_cama_min_colab}")

                #Q12    ----    Durante o último mês, que horas você geralmente acordou para ir trabalhar? (hora)

                if acordou_trab_colab in horas_cama_e_acordou_xpaths:
                    click_button(driver, '//*[@id="question-list"]/div[4]/div[2]/div/div/div')
                    click_button(driver, horas_cama_e_acordou_xpaths[acordou_trab_colab])

                #Q13    ----    Durante o último mês, que horas você geralmente acordou para ir trabalhar? (minutos)
                
                if acordou_trab_min_colab in acordou_trab_min_xpaths:
                    click_button(driver, '//*[@id="question-list"]/div[5]/div[2]/div/div/div')
                    click_button(driver, acordou_trab_min_xpaths[acordou_trab_min_colab])
                print(f"Minutos acordar: {acordou_trab_min_colab}")

                #Q14    ----    Durante o último mês, quantas horas de sono você teve por noite?
                
                if horas_sono_colab in horas_sono_xpaths:
                    click_button(driver,'//*[@id="question-list"]/div[6]/div[2]/div/div/div')
                    click_button(driver, horas_sono_xpaths[horas_sono_colab])
                print(f"Horas de sono: {horas_sono_colab}")

                #N E X T   P A G E
                
                click_button(driver, '//*[@id="form-main-content1"]/div/div/div[2]/div[3]/div/button[2]')
                
                #P A G E   0 3
                #Q18
                
                
                # VAR LOOP RANDOM                
                total_question_forms1 = 9 
                quant_pittsburgh_1 = random.randint(0,1)
                quest_pittsburgh_1 = random.sample(range(2, total_question_forms1 + 1 ), quant_pittsburgh_1)
                print(quest_pittsburgh_1)
                
                for i in range(2, 11):
                    pittsburgh_0 = f'//*[@id="question-list"]/div[{i}]/div[2]/div/div/div[1]/div'
                    pittsburgh_1 = f'//*[@id="question-list"]/div[{i}]/div[2]/div/div/div[2]/div'
                    if i in quest_pittsburgh_1:
                        click_button(driver, pittsburgh_1)
                    else:
                        click_button(driver, pittsburgh_0)
                #Q27
                click_button(driver,'//*[@id="question-list"]/div[11]/div[2]/div/div/div[2]/div')
                
                total_question_formsb = 4 
                quant_pittsburgh_1b = random.randint(0,2)
                quest_pittsburgh_1b = random.sample(range(12, total_question_formsb + 12 ), quant_pittsburgh_1b)
                print(quest_pittsburgh_1b)
                for i in range(12, 16):
                    pittsburgh_0b = f'//*[@id="question-list"]/div[{i}]/div[2]/div/div/div[1]/div'
                    pittsburgh_1b = f'//*[@id="question-list"]/div[{i}]/div[2]/div/div/div[2]/div'
                    if i in quest_pittsburgh_1b:
                        click_button(driver, pittsburgh_1b)
                    else:
                        click_button(driver, pittsburgh_0b)
                #32
                if esposa_parc_colab == "Não":
                    click_button(driver, '//*[@id="question-list"]/div[16]/div[2]/div/div/div[1]/div')   

                
                elif esposa_parc_colab == "Sim, parceiro ou colega, mas em outro quarto":
                    click_button(driver, '//*[@id="question-list"]/div[16]/div[2]/div/div/div[2]/div')
                    click_button(driver, '//*[@id="question-list"]/div[17]/div[2]/div/div/div[1]/div')
                    click_button(driver,'//*[@id="question-list"]/div[18]/div[2]/div/div/div[1]/div')
                 
                elif esposa_parc_colab == "Sim, parceiro no mesmo quarto, mas em outra cama":
                    click_button(driver, '//*[@id="question-list"]/div[16]/div[2]/div/div/div[3]/div')
                    click_button(driver, '//*[@id="question-list"]/div[17]/div[2]/div/div/div[1]/div')
                    click_button(driver,'//*[@id="question-list"]/div[18]/div[2]/div/div/div[1]/div')
                
                        
                        
                elif esposa_parc_colab == "Sim, parceiro/a na mesma cama":
                    click_button(driver, '//*[@id="question-list"]/div[16]/div[2]/div/div/div[4]/div')
                   
                tempomedio = random.randint(85, 367)
                tempomedioeminutos = tempomedio / 60
                print(f"Tempo de Resposta -> {tempomedioeminutos:02}" )
                time.sleep(tempomedio)

                # S E N D   B U T T O N
                click_button(driver, '//*[@id="form-main-content1"]/div/div/div[2]/div[4]/div/button[2]')
                #RECOMEÇAR
                click_button(driver, '//*[@id="form-main-content1"]/div/div/div[2]/div[1]/div[2]/div[5]')
                tempo_entre_respostas = random.randint(0,187) # Tempo de espera entre respostas de 3 minutos e 7 segundos 
                tempo_entre_respostas_em_minutos1 = tempo_entre_respostas/60
                print(f"Tempo entre resposta -> {tempo_entre_respostas_em_minutos1:02}")
                time.sleep(tempo_entre_respostas)
                
            except Exception as e:
                print(f"Erro na linha {index + 1}:{e}")

    finally:
        driver.quit()

if __name__ =="__main__":
    main()
