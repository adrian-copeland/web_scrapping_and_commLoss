from selenium import webdriver 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service # i added this line
from selenium.webdriver.support.ui import Select
import os
import time
import datetime
import win32com.client as win32
# from selenium.webdriver.chrome.service import Service
# from selenium.common.exceptions import TimeoutException , NoSuchElementException

DOWNLOAD_DIR = os.path.join(os.getcwd(), "lists_downloaded")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)


def get_chrome_options():
    chrome_options = Options()
    ### add the language argument
    chrome_options.add_argument("--lang=en")
    ###add experimental options for file downloads
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    return chrome_options

def inicio_pasword(file_path: str, driver: webdriver.Chrome, wait: WebDriverWait, lista_correo_errores: list):
    """This is a special function that introduces credentials in Connect +. 
    -----------------------------------------
    Parameters:
    ----------
    file_path : str
        Path where the credentials file is located
    driver : webdriver.Chrome
        Chrome webdriver instance   
    wait : WebDriverWait
        WebDriverWait instance for waiting for elements
    lista_correo_errores : list
        List of email addresses for error notifications
    -----------------------
    Returns:
    ------- 
    None
    """
    sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
    driver.switch_to.frame(sesion_frame)

    username_input = wait.until(EC.presence_of_element_located((By.NAME, "_UserName")))
    password_input = wait.until(EC.presence_of_element_located((By.NAME, "_Password")))
    login_button = wait.until(EC.element_to_be_clickable((By.NAME, "loginButton")))
    #password_input = driver.find_element(By.NAME, "_Password") i changed this
    #login_button = driver.find_element(By.NAME, "loginButton") i changed this
    print('found input fields\n')
    #k_path = path_programa + '\\txt\\Credenciales_Walmart.txt'
    k_path = file_path
    print(k_path)
    with open(k_path, 'r', encoding='utf-8') as file:
        lineas = file.readlines()

    print('Attempting login...\n')
    #cada línea a una variable
    u = lineas[0].strip()
    p = lineas[1].strip()

    username_input.send_keys(u)
    password_input.send_keys(p)
    login_button.click()

    #login_button = driver.find_element(By.ID, "ssoButton")
    #login_button.click()
    #print("Inicio de sesión exitoso!\n")
                #login_button.click()
    #switch back out of the frame after the clic, as the logging process usually loads a new page OUSIDE the frame.
    driver.switch_to.default_content()

    #print("Inicio de sesión exitoso!\n")
    try:
        element = driver.find_element(By.CLASS_NAME, "invalidFields")
        if "Password will expire in" in element.text:
            login_button.click()
            send_mail_app_escritorio(2, lista_correo_errores, lista_correo_errores, 'Contraseña de cuenta de descarga de alarmas por expirar', f'La contraseña de la cuenta {username_input_walmart} está por expirar', lista_correo_errores)
        else:
            print("El elemento existe, pero el texto no coincide.")
            time.sleep(2)
    except:
        pass



# Use the actual directory path from your options
#NEW_FILENAME = "My_Renamed_Report.xlsx"
def new_filename():
    """Generates a new filename based on the current date and time."""
    now = datetime.datetime.now()
    return f"Report_{now.strftime('%Y%m%d_%H%M')}.xlsx"


def rename_downloaded_file(download_dir, new_name):
    # 1. Wait for the download to finish (a simple, but often reliable wait)
    #time.sleep(5) # Adjust this wait time based on expected download speed

    # 2. Find the newest file in the directory
    # List all files and sort them by creation time (newest first)
    files = [os.path.join(download_dir, f) for f in os.listdir(download_dir)]
    files.sort(key=os.path.getmtime, reverse=True)
    
    if files:
        original_filepath = files[0] # Assumes the newest file is the one you just downloaded
        new_filepath = os.path.join(download_dir, new_name)

        # 3. Rename the file
        try:
            os.rename(original_filepath, new_filepath)
            print(f"Successfully renamed '{os.path.basename(original_filepath)}' to '{new_name}'")
        except FileNotFoundError:
            print("Error: Downloaded file not found.")
        except Exception as e:
            print(f"Error renaming file: {e}")
    else:
        print("No files found in the download directory.")

# Call this function after your script initiates the download and navigates away
# Example usage:
# driver.find_element(By.ID, "download_button").click()
# rename_downloaded_file(DOWNLOAD_DIR, NEW_FILENAME)
#def inicio_walmart():
 #   sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
  #  driver.switch_to.frame(sesion_frame)

    #username_input_walmart = wait.until(EC.presence_of_element_located((By.NAME, "_UserName")))
             #   password_input_walmart = driver.find_element(By.NAME, "_Password")
             #   login_button = driver.find_element(By.ID, "ssoButton") #//*[@id="ssoButton"]

             #   login_button.click()

             #print("Inicio de sesión exitoso!\n")
            #iniciar sesión
            #try:
                
             #   a = inicio_walmart()

#Funcion para envio de alarmas 
def send_mail_app_escritorio(importancia: int,destinatario: list,copia:list ,subjet: list ,cuerpo, lista_correo_errores:list):
    destinatario = ';'.join(destinatario)
    copia = ';'.join(copia) 

    olApp = win32.Dispatch('Outlook.Application')
    mailItem = olApp.CreateItem(0)
    mailItem.Importance = importancia
    mailItem.to = destinatario
    mailItem.CC = copia
    mailItem.Subject = subjet
    # Creamos el objeto mensaje
    mailItem.HTMLbody = cuerpo

    try:
        mailItem.send
    except Exception as e:
        error_message = str(e)

        send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, 'Fallo en envío de alarmas de Copeland', error_message)
        print('Error al enviar correo:',repr(e))
        print(f'Error: {e}\n')
        time.sleep(5)

def extraer_alarmas_connect(previous_days: int, lista_correo_errores: list):
    for i in range(3):
    # try:
        path_programa = os.getcwd()   ## path where the code is living
        driver_path = os.path.join(path_programa, "chromedriver-win64", "chromedriver.exe")
        credentials_path = os.path.join(path_programa, "credentials.txt")

        # Configurar opciones de Chrome para abrir en inglés
        #options = Options()
        #options.add_argument("--lang=en")  # forza el idioma de chrome a inglés



        service = webdriver.ChromeService(driver_path)  ### I changed this line
        #service = Service(executable_path=driver_path) ### updated line
        driver = webdriver.Chrome(service=service, options=get_chrome_options())
        driver.maximize_window()
        zoom_lever = '50%'
        driver.execute_script(f'document.body.style.zoom="{zoom_lever}"')

        #tiempo de espera
        wait = WebDriverWait(driver, 20)

        try:
            # Abrir la página web
            #url_walmart = "https://walmartca.my-connectplus.com/walmartca/"   ###### una vez aqui, probably i need to put this as a parameter
            url_cpla = "https://cpla.my-connectplus.com/cpla/"     ###### una vez aqui, enfoque en este
            driver.get(url_cpla)
            #driver.get(url_walmart)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "html")))
            print("Página cargada\n")
            print('Entering login function\n')
            
            inicio_pasword(file_path= credentials_path, driver=driver, wait=wait, lista_correo_errores=lista_correo_errores)
            print('get out password function\n')
            

            #def inicio_walmart():
             #   sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
             #   driver.switch_to.frame(sesion_frame)

             #   username_input_walmart = wait.until(EC.presence_of_element_located((By.NAME, "_UserName")))
             #   password_input_walmart = driver.find_element(By.NAME, "_Password")
             #   login_button = driver.find_element(By.ID, "ssoButton") #//*[@id="ssoButton"]

             #   login_button.click()

             #print("Inicio de sesión exitoso!\n")
            #iniciar sesión
            #try:
                
             #   a = inicio_walmart()
            print("You get successfully logged in\n")
            #return None

        except Exception as e:
            error_message = str(e)
            print('Error during login:', str(e))
            #send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, 'No se encontró el formulario de inicio de sesión.', error_message, lista_correo_errores)
            print("Fallo el inicio de sesión.\n")
            driver.quit()
            exit()

        try:
            print('Enter try to click the button to go to the data section\n')
            driver.switch_to.default_content()
            sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "navId")))
            driver.switch_to.frame(sesion_frame)

            #wait.until(EC.element_to_be_clickable((By.XPATH, '//img[@title="Alarm"]'))).click() #<img class="header-img" src="/cpla/Images/alarm.png" title="Alarm">
            wait.until(EC.element_to_be_clickable((By.XPATH, '//img[@title="Graph/Watch"]'))).click() #<img class="header-img" src="/cpla/Images/graph.png" title="Graph/Watch">
            #wait.until(EC.element_to_be_clickable((By.XPATH, "//img[@id='clearCheckboxesButton']/following-sibling::img[@title='Graph/Watch']"))).click()

            #graph_image = wait.until(
             #   EC.presence_of_all_elements_located((By.XPATH, "//img[@title='Graph/Watch']"))
            #)
            ##force the click using JavaScript
            #driver.execute_script("arguments[0].click();", graph_image[0])
            #driver.execute_script("arguments[0].click();", graph_image)

            print('Clic en boton alarmas')

            #driver.switch_to.default_content()
            #graph_image = wait.until(   
                #EC.element_to_be_clickable((By.XPATH, '//*[@id="ext-gen17"]/div[2]/div/table/tbody/tr/td[3]/div/img[2]'))
             #   EC.element_to_be_clickable((By.XPATH, '//imag[@title="Graph/Watch"]'))
                #EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, 'graph.png')]"))
            #)
            #graph_image.click()
            print('Clic en boton graph/watch')
            
            # espera que la pantalla gris desaparezca
            #wait.until(EC.invisibility_of_element((By.CLASS_NAME, "ext-el-mask"))) 
            #################### selecting +Tiendas CommLoss from the dropdown #########################
            driver.switch_to.default_content()
            sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "mainId")))
            driver.switch_to.frame(sesion_frame)
            #find the dropdown element by its id    
            dropdown_element = wait.until(EC.presence_of_element_located((By.ID, "listsselection")))
            
            #driver.find_element(By.ID, "listselection")
            #pass the located dropdown element to the select class constructor
            
            select_object = Select(dropdown_element)
            wait.until(
                EC.presence_of_element_located((By.XPATH, '//option[@value="599"]'))
            )
            select_object.select_by_value("599")
            #select_object.select_by_visible_text("+Tiendas CommLoss")
            #<option value="599" class="publiclistoption">+Tiendas CommLoss</option>

            ############## SELecting button Retrieve Logs  + Export ######################
            ##goes to the right frame
            driver.switch_to.default_content()
            sesion_frame = wait.until(EC.presence_of_element_located((By.ID, "mainId")))
            driver.switch_to.frame(sesion_frame)

            button_element = wait.until(
                EC.element_to_be_clickable((By.ID, "exportLogsButton"))
            )
            button_element.click()
            ##<button type="button" id="exportLogsButton" class="button contentNormal" onclick="showDateRangeSelection(true)">
                        # Retrieve Logs + Export
                     #</button>
            
            
            ##################### Select option Condense from the dropdown menu (stays in the same frame)########################
            dropdown_element = wait.until(EC.presence_of_element_located((By.ID, "exportFormat")))
            select_object = Select(dropdown_element)
            wait.until(
                EC.presence_of_element_located((By.XPATH, '//option[@value="2"]'))
            )
            ##select_object.select_by_value("2")
            select_object.select_by_visible_text("Condensed")
            ##<select id="exportFormat" class="controls">
		            #<option value="1">Comprehensive</option>
		            #<option value="2">Condensed</option>
		          #</select>

            ######################### start time and end time (stays in the same frame) ##########################
        
            now = datetime.datetime.now()
            end_time = now.strftime("%Y-%m-%d %H:%M:%S")
            
            initial_date = now - datetime.timedelta(days=previous_days)
            intitial_date = initial_date.strftime("%Y-%m-%d %H:%M:%S")

            ### start time    
            start_time_input_field = wait.until(
                EC.presence_of_element_located((By.ID, "startTimeField"))
            )
            start_time_input_field.clear()
            start_time_input_field.send_keys(intitial_date)
            #<input id="startTimeField" name="startTimeField" class="controls" type="text">
            
            ####end time 
            end_time_input_field = wait.until(
                EC.presence_of_element_located((By.ID, "endTimeField"))
            )
            end_time_input_field.clear()
            end_time_input_field.send_keys(end_time)
            #<input id="endTimeField" name="endTimeField" class="controls" type="text">

            ################## Click on Go button and download  (still in the same frame)#######################
            button_element = wait.until(
                EC.element_to_be_clickable((By.ID, "userDownloadStart"))
            )
            button_element.click()
            time.sleep(100)
            time.sleep(10)
            ##<button id="userDownloadStart" type="button" class="controls dialogButton">Go</button>#


            time.sleep(5)
            print('Closing the browser...')
            driver.quit()
            return None

            
            
            #wait.until(EC.element_to_be_clickable((By.XPATH, '//img[@title="Alarm"]'))).click() #<img class="header-img" src="/cpla/Images/alarm.png" title="Alarm">
            #print('Clic en boton alarmas')
        except Exception as e:
                print("Fallo la descarga de alarmas desde Connect+.\n")
                error_message = str(e)
                print(str(e))
                #send_mail_app_escritorio(int(2), lista_correo_errores, lista_correo_errores, f'Fallo en el intento {i+1} en descarga de alarmas de Copeland', error_message, lista_correo_errores)
                if i > 2:
                    return None
                ##continue  ################
                return None 


if __name__ == "__main__":
    #formato_tienda = ['MBG', 'BGA'] , not used anymore
    previous_days = 1
    lista_correo_errores = ['adrianpebus@gmail.com']

    #path_archivo_descargado = extraer_alarmas_connect( formato_tienda, lista_correo_errores)
    extraer_alarmas_connect(previous_days=previous_days, lista_correo_errores=lista_correo_errores)
    new_name = new_filename()
    rename_downloaded_file(DOWNLOAD_DIR, new_name=new_name)
    print(f'Successfull run!!!')

