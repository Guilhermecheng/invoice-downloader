from selenium import webdriver
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains

# retry logic
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time
from excel import ordem_de_servico, nota_1_lista, nota_2_lista, loja_lista
from senha import dealernetLogin, dealernetPassword

# defining geckodriver and nav
path = r"C:\Users\guilh\AppData\Local\Programs\Python\Python39"

fp = webdriver.FirefoxProfile()

fp.set_preference("browser.download.folderList",2)
fp.set_preference("browser.download.manager.showWhenStarting", False)
fp.set_preference("browser.download.dir", r"D:\downpdf")
fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf") # ou outro
fp.set_preference("pdfjs.disabled", "true")

navegador = webdriver.Firefox(firefox_profile=fp)

# defining login and password from Dealernet
meu_login_dealernet = dealernetLogin
minha_senha_dealernet = dealernetPassword

loja_dict = {
    "SPG IPIRANGA 0004": 16,
    "BRG CURITIBA 0006": 5,
    "SPG GASTAO 0002": 17,
    "BRG FLORIANOPOLIS 0003": 3,
    "BRN DISTRITO FEDERAL 0008": 7,
    "BRG PORTO ALEGRE 0002": 2
}

def login_in_dealernet():
    # finding where to put data
    login = navegador.find_element_by_id("vUSUARIO_IDENTIFICADORALTERNATIVO")
    senha = navegador.find_element_by_id("vUSUARIOSENHA_SENHA")
    confirm_button = navegador.find_element_by_id("IMAGE3")    

    # send keys and confirming
    login.send_keys(meu_login_dealernet)
    senha.send_keys(minha_senha_dealernet)
    confirm_button.send_keys(Keys.ENTER)

def abrir_os(numeros_os, dealer):
    oficina_button = navegador.find_element_by_id("W5|_253_|Oficina")
    oficina_button.click()

    time.sleep(1)
    os_button = navegador.find_element_by_id("W5|_289_|OrdemdeServiço/Orçamento")
    os_button.click()

    time.sleep(10)
    # getting into iframe
    def listing_iframes():
        number_of_frames = len(navegador.find_elements_by_tag_name("iframe"))
        array_tries = 0
        for x in range(number_of_frames):
            # print(array_tries)
            navegador.switch_to.frame(x)
            try:
                os_num_placeholder = navegador.find_element_by_id("vCODIGO")
                navegador.switch_to.default_content()
                array_tries = x
                return array_tries

            except NoSuchElementException:
                navegador.switch_to.default_content()
                print("no OS num placeholder in " + str(x))

        return print("no frame found")

    iframe_num = listing_iframes()
    # print(iframe_num)
    if iframe_num != False:
        navegador.switch_to.frame(iframe_num)

        os_num_placeholder = navegador.find_element_by_id("vCODIGO")
        os_num_placeholder.send_keys(numeros_os)

        dealer_select = Select(navegador.find_element_by_id("vEMPRESA_CODIGO"))
        dealer_select.select_by_value(str(dealer))

        # navegador.find_element_by_id("IMAGE1").click()
        ActionChains(navegador).move_to_element(navegador.find_element_by_id("IMAGE1")).click().perform()

        time.sleep(5)
        navegador.find_element_by_id("vDISPLAYOS_0001").click()
        navegador.switch_to.default_content()
    # table = navegador.find_element_by_id("TABLE01")
    

def getting_os_nfs(downloads, nota_de_servico, nota_de_peca):
    ne_serv_downloaded = False
    ne_peca_downloaded = False
    nf_download_check = [ne_serv_downloaded, ne_peca_downloaded]
    
    notas = [nota_de_servico, nota_de_peca]
    download_counter = downloads

    def listing_iframes(quesito):        
        number_of_frames = len(navegador.find_elements_by_tag_name("iframe"))
        array_tries = 0
        for x in range(number_of_frames):
            # print(array_tries)
            navegador.switch_to.frame(x) 

            try:    
                

                os_num_placeholder = navegador.find_element_by_id(quesito)
                navegador.switch_to.default_content()
                array_tries = x
                return array_tries

            except NoSuchElementException:
                navegador.switch_to.default_content()
                # print("no OS num placeholder in " + str(x))

        return print("no frame found")


    for notas_length in range(len(notas)):
        if notas[notas_length] != "none":           
            iframe_num = listing_iframes("TABLENF")

            # downloading
            if iframe_num != False:            
                navegador.switch_to.frame(iframe_num)

                # table = navegador.find_element_by_id("GridnfContainerTbl")
                rowcount = len(navegador.find_elements_by_xpath('//*[@id="GridnfContainerTbl"]/tbody/tr'))
                for count in range(rowcount):
                    span_id = "span_vNOTAFISCAL_NUMERO_000" + str(count + 1)
                    # print(span_id)
                    nf_number = navegador.find_element_by_xpath(f'//*[@id="{span_id}"]/a').text
                    # .replace(" ","")
                    # print(nf_number)

                    if (nf_number == notas[notas_length] and nf_download_check[notas_length] == False):
                        navegador.find_element_by_id(span_id).click()             

                        navegador.switch_to.default_content()
                        time.sleep(10)

                        WebDriverWait(navegador, 40).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
                        iframe_from_os = listing_iframes("BTNIMPRIMIR")

                        if iframe_from_os != False:
                            navegador.switch_to.frame(iframe_from_os)
                            WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.ID, "BTNIMPRIMIR"))).click()
                            navegador.switch_to.default_content()
                            
                            time.sleep(5)

                            if download_counter < 1:
                                iframe_from_download = listing_iframes("download")
                                if iframe_from_download != False:
                                    navegador.switch_to.frame(iframe_from_download)
                                    navegador.find_element_by_id("download").click()
                                    download_counter = download_counter + 1
                                    print("files downloaded: " + str(download_counter))
                                    navegador.switch_to.default_content()
                                    time.sleep(5)

                                    navegador.find_element_by_id("ext-gen184").click()
                                    time.sleep(3)
                            else:

                                navegador.switch_to.default_content()
                                # navegador.find_element_by_class_name("x-tool-close").click()
                                ActionChains(navegador).move_to_element(navegador.find_element_by_class_name("x-tool-close")).click().perform()
                                download_counter = download_counter + 1
                                print(str(notas[notas_length]) + " downloaded")
                                print("files downloaded: " + str(download_counter))
                                time.sleep(5)

                            nf_download_check[notas_length] = True 
                            ActionChains(navegador).move_to_element(navegador.find_element_by_class_name("x-tool-close")).click().perform()
                            time.sleep(5)
                        break
                    else:
                        continue
                
    ActionChains(navegador).move_to_element(navegador.find_element_by_class_name("x-tool-close")).click().perform()
    return download_counter
            

# main logic
navegador.get("http://dealernet.shcnet.com.br/LoginAux.aspx?Windows")
time.sleep(5)
login_in_dealernet()
time.sleep(10)

# dealer_procurado = input("Opções: \n 13: CPO \n 5: CMT \n 3: FLO \n 17: GVI \n 20: JAF \n 12: NIT \n 16: NAC/ IPI \n 2: POC \n 9: REI \n 19: ABC \n 7: SIA \n Selecione a concessionária desejada: ")

download_num = 0
os = ordem_de_servico
nf_serv = nota_1_lista
nf_peça = nota_2_lista
lojas = loja_lista

for os_os in range(len(os)):
    loja_proc = lojas[os_os]
    abrir_os(os[os_os], loja_dict[loja_proc])
    time.sleep(10)  
    download_num2 = getting_os_nfs(download_num, nf_serv[os_os], nf_peça[os_os])
    download_num = download_num2

navegador.quit()
print("Script over!")
print("Total files downloaded: "+ str(download_num))