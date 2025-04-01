from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
import pyautogui
import time
import zipfile
import os
import shutil
from datetime import datetime

# Configurar opções do Edge
edge_options = Options()

# Flags para desabilitar a proteção de download e outras configurações de segurança
edge_options.add_argument("--safebrowsing-disable-download-protection")  # Desativa a proteção de download
edge_options.add_argument("--disable-web-security")  # Desabilita a segurança da web
edge_options.add_argument("--allow-running-insecure-content")  # Permite conteúdo inseguro
edge_options.add_argument("--ignore-certificate-errors")  # Ignora erros de certificado
edge_options.add_argument("--disable-features=IsolateOrigins,site-per-process")  # Desativa recursos de segurança extra
edge_options.add_experimental_option("prefs", {
    "download.default_directory": r"LOCAL QUE SERÁ BAIXADO",  # Caminho do download
    "download.prompt_for_download": False,  # Não perguntar onde salvar
    "download.directory_upgrade": True,  # Permitir atualizações do diretório
    "safebrowsing.enabled": False,  # Desativa o Safe Browsing
    "profile.default_content_setting_values.automatic_downloads": 1  # Permitir múltiplos downloads automáticos
})

# Inicializar o WebDriver com as opções
driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()), options=edge_options)

# Agora, o WebDriver abrirá com as configurações para permitir downloads sem proteção.

# Configurar o WebDriver
driver.get("SITE")

# Aguarda carregamento da página
time.sleep(3)

# Preencher login e senha
driver.find_element(By.ID, "filled-adornment-username").send_keys("LOGIN")
driver.find_element(By.ID, "filled-adornment-pass").send_keys("SENHA")

# Clicar no botão "Entrar"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Entrar')]"))
).click()

time.sleep(5)

# Clicar no botão "Relatórios"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div/div'))
).click()

time.sleep(5)

# Clicar no botão "Gerar Relatórios"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/main/div/div[1]/button[2]'))
).click()

# Aguarda o menu expansível aparecer
menu_xpath = "/html/body/div[4]/div[3]/div"
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, menu_xpath)))
print("✅ Menu de relatórios carregado!")

time.sleep(1)

# Clique no botão desejado
button_xpath = "/html/body/div[4]/div[3]/div/div[3]/div[1]/div/div/div/button"
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, button_xpath))).click()
print("✅ Botão clicado!")

# Aguarda o calendário abrir e encontra o dia 1
calendar_xpath = "/html/body/div[5]/div[2]/div/div/div/div[2]/div/div[2]//button[text()='1']"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, calendar_xpath))
).click()
print("✅ Dia 1 clicado!")

# Clicar novamente no botão original
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, button_xpath))
).click()
print("✅ Botão original clicado novamente!")

# Descer a barra de navegação até o final da página
scroll_xpath = "/html/body/div[4]/div[3]"
scroll_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, scroll_xpath)))
driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", scroll_element)
print("✅ Rolagem para o final da página concluída!")

# Clicar no primeiro input
first_input_xpath = "/html/body/div[4]/div[3]/div/div[4]/label/span[1]/input"
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, first_input_xpath)))
time.sleep(2)
input_element = driver.find_element(By.XPATH, first_input_xpath)
driver.execute_script("arguments[0].click();", input_element)
print("✅ Primeiro input clicado!")

# Clicar no segundo input
second_input_xpath = "/html/body/div[4]/div[3]/div/div[5]/div/label/span[1]/input"
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, second_input_xpath)))
time.sleep(2)
second_input_element = driver.find_element(By.XPATH, second_input_xpath)
driver.execute_script("arguments[0].click();", second_input_element)
print("✅ Segundo input clicado com sucesso!")

# Clicar no botão adicional
additional_button_xpath = "/html/body/div[4]/div[3]/div/div[6]/button"
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, additional_button_xpath))
).click()
print("✅ Botão adicional clicado!")

# Monitorar carregamento
timeout = 180
start_time = time.time()
loading_xpath = "//*[@id='root']/div/main/div/div[2]/table/tbody/div[1]/div/div/tr[1]/td[4]"
page_loaded = False

while time.time() - start_time < timeout and not page_loaded:
    try:
        # Espera o elemento de carregamento aparecer
        loading_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, loading_xpath))
        )
        print(f"⏳ Aguardando carregamento...")
        time.sleep(10)
        driver.refresh()  # Atualiza a página a cada 10 segundos
        print("🔄 Página recarregada!")
    except:
        # Quando o carregamento termina
        print("✅ Carregamento finalizado! Prosseguindo...")
        page_loaded = True  # Marca que a página está carregada
        break

# Clicar no botão uma vez
button_xpath = "//*[@id='root']/div/main/div/div[2]/table/tbody/div[1]/div/div/tr[1]/td[4]/button"
WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, button_xpath))
).click()
print("✅ Botão da primeira linha clicado!")

time.sleep(30)

def extract_and_rename_zip(download_directory, target_directory):
    # Listar arquivos na pasta de downloads
    files = [f for f in os.listdir(download_directory) if f.endswith('.zip')]
    
    if not files:
        print("Nenhum arquivo .zip encontrado!")
        return None

    # Obter o arquivo .zip mais recente baseado na data de modificação
    latest_file = max(files, key=lambda f: os.path.getmtime(os.path.join(download_directory, f)))
    download_path = os.path.join(download_directory, latest_file)
    
    # Extrair o nome do arquivo sem a extensão .zip
    zip_filename = os.path.basename(download_path)
    folder_name = zip_filename.replace(".zip", "")
    
    # Caminho para onde o arquivo será extraído
    extracted_path = os.path.join(download_directory, folder_name)
    
    # Extrair o arquivo .zip
    with zipfile.ZipFile(download_path, 'r') as zip_ref:
        zip_ref.extractall(extracted_path)
    
    print(f"✅ Arquivo {zip_filename} extraído para {extracted_path}")
    
    # Renomear o arquivo extraído
    current_month_year = datetime.now().strftime('%m_%Y')
    renamed_file_name = f"PONTAL_{current_month_year}.csv"
    
    extracted_file_path = os.path.join(extracted_path, os.listdir(extracted_path)[0])  # Assume que o arquivo dentro da pasta extraída é o primeiro
    renamed_file_path = os.path.join(extracted_path, renamed_file_name)
    os.rename(extracted_file_path, renamed_file_path)
    
    print(f"✅ Arquivo renomeado para {renamed_file_name}")
    
    # Mover o arquivo renomeado para o diretório de destino
    final_path = os.path.join(target_directory, renamed_file_name)
    shutil.move(renamed_file_path, final_path)
    
    print(f"✅ Arquivo movido para {final_path}")
    return final_path

# Caminho do diretório de downloads
download_directory = r"LOCAL QUE FOI BAIXADO"

# Caminho do diretório de destino
target_directory = r"SALVAMENTO FINAL DO ARQUIVO"

# Chama a função de extração e renomeação
final_file_path = extract_and_rename_zip(download_directory, target_directory)

# Remover o arquivo ZIP original da pasta Downloads após a movimentação
if final_file_path:
    download_path = os.path.join(download_directory, os.path.basename(final_file_path).replace(".csv", ".zip"))
    extracted_path = os.path.join(download_directory, os.path.basename(final_file_path).replace(".csv", ""))

    if os.path.exists(download_path):
        try:
            os.remove(download_path)
            print(f"🗑️ Arquivo ZIP original removido: {download_path}")
        except Exception as e:
            print(f"❌ Erro ao remover o arquivo ZIP original: {e}")

    # Remover a pasta extraída após a movimentação do arquivo
    if os.path.exists(extracted_path):
        try:
            shutil.rmtree(extracted_path)
            print(f"🗑️ Pasta extraída removida: {extracted_path}")
        except Exception as e:
            print(f"❌ Erro ao remover a pasta extraída: {e}")

print("✅ Processo concluído com sucesso!")
