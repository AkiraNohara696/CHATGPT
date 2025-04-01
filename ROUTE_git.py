import time
import os
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuração do Selenium
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")  # Desabilitar aceleração de hardware
options.add_experimental_option("detach", True)  # Manter o navegador aberto

# Configuração para evitar perguntas antes de baixar arquivos
download_dir = os.path.join(os.path.expanduser("~"), "Downloads")  # Diretório para downloads (pasta Downloads local)
preferences = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "download.default_directory": download_dir
}
options.add_experimental_option("prefs", preferences)

# Inicializar o navegador
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Acessando o site
driver.get("SITE")

# Esperar até que o botão "Avançadas" esteja presente e clicar nele
wait = WebDriverWait(driver, 20)  # Espera de até 20 segundos
advanced_button = wait.until(EC.presence_of_element_located((By.ID, "details-button")))
advanced_button.click()

# Esperar até que o link "Ir para sipserver.routetelecom.com.br (não seguro)" esteja presente e clicar nele
proceed_link = wait.until(EC.presence_of_element_located((By.ID, "proceed-link")))
proceed_link.click()

# Aguardar 10 segundos para garantir que a página foi carregada
time.sleep(10)

# Esperar até que o campo de username esteja presente
username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
# Esperar até que o campo de senha esteja presente
password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))

# Preenchendo os campos
username_field.send_keys("LOGIN")
password_field.send_keys("SENHA")

# Submetendo o formulário
password_field.send_keys(Keys.RETURN)

# Aguardar a página carregar após o login
time.sleep(5)

# Esperar até que o último ícone da lista esteja visível e clicável
last_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[@id='main-menu']/li[last()]/a")))
last_icon.click()

# Aguardar o menu dropdown aparecer e esperar que ele esteja visível
try:
    dropdown_menu = wait.until(EC.visibility_of_element_located((By.XPATH, "//ul[@class='dropdown-menu']")))
    print("Menu dropdown visível.")
except Exception as e:
    print(f"Erro ao esperar o menu dropdown: {e}")

# Esperar até que a opção desejada dentro do menu dropdown esteja visível e clicável
try:
    did_option = wait.until(EC.element_to_be_clickable((By.ID, "relatorioLigacoesDID")))
    did_option.click()
    print("Opção 'Ligações Recebidas - DID' clicada.")
except Exception as e:
    print(f"Erro ao clicar na opção 'Ligações Recebidas - DID': {e}")

# Aguardar a página carregar após o clique
time.sleep(5)

# Preencher o campo de data com o primeiro dia do mês
try:
    # Obter o primeiro dia do mês atual no formato DD/MM/YYYY
    primeiro_dia_mes = datetime.now().replace(day=1).strftime("%d/%m/%Y")
    # Esperar até que o campo de data esteja presente e clicável
    data_input = wait.until(EC.element_to_be_clickable((By.ID, "txtDataI")))
    # Clicar no campo para ativá-lo
    data_input.click()
    # Limpar o campo antes de inserir a nova data
    data_input.clear()
    # Inserir a nova data
    data_input.send_keys(primeiro_dia_mes)
    # Pressionar TAB para confirmar a entrada e sair do campo
    data_input.send_keys(Keys.TAB)
    print(f"Campo de data preenchido com: {primeiro_dia_mes}")
except Exception as e:
    print(f"Erro ao preencher o campo de data: {e}")

# Esperar até que o botão "Pesquisar" esteja visível e clicável
try:
    search_button = wait.until(EC.element_to_be_clickable((By.ID, "btnSearch")))
    search_button.click()  # Clicar no botão Pesquisar
    print("Botão 'Pesquisar' clicado.")
except Exception as e:
    print(f"Erro ao clicar no botão 'Pesquisar': {e}")

# Aguardar um momento para garantir que a pesquisa foi realizada
time.sleep(3)

# Esperar até que o botão "Exportar" esteja visível e clicável
try:
    export_button = wait.until(EC.element_to_be_clickable((By.ID, "btnExport")))
    export_button.click()  # Clicar no botão Exportar
    print("Botão 'Exportar' clicado.")
except Exception as e:
    print(f"Erro ao clicar no botão 'Exportar': {e}")

# Aguardar o download ser concluído
time.sleep(10)

# Obter o nome do arquivo baixado
downloaded_files = os.listdir(download_dir)
latest_file = max([os.path.join(download_dir, f) for f in downloaded_files], key=os.path.getctime)

# Renomear o arquivo com o formato "Route_mês_ano"
current_month_year = datetime.now().strftime("%m_%Y")
new_file_name = f"Route_{current_month_year}.csv"

# Renomear o arquivo baixado
os.rename(latest_file, os.path.join(download_dir, new_file_name))

# Definir o destino
destination = r"LOCAL FINAL DE SALVAMENTO"

# Verificar se o arquivo já existe no destino
destination_file_path = os.path.join(destination, new_file_name)
if os.path.exists(destination_file_path):
    try:
        # Se o arquivo já existir, removê-lo
        os.remove(destination_file_path)
        print(f"Arquivo existente {new_file_name} removido.")
    except Exception as e:
        print(f"Erro ao remover arquivo existente {new_file_name}: {e}")

# Mover o arquivo para a pasta de destino
shutil.move(os.path.join(download_dir, new_file_name), destination_file_path)
print(f"Arquivo movido e renomeado para: {destination_file_path}")

# Remover o arquivo da pasta Downloads após a movimentação
final_file_path = os.path.join(download_dir, new_file_name)
if os.path.exists(final_file_path):
    try:
        os.remove(final_file_path)
        print("Arquivo original removido da pasta Downloads para evitar acúmulo.")
    except Exception as e:
        print(f"Erro ao remover o arquivo original da pasta Downloads: {e}")

# Fechar o navegador (opcional)
driver.quit()

print("Processo finalizado!!!")