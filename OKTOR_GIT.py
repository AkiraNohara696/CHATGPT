import os
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
import time
from datetime import datetime

# Configuração do Selenium
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")  # Desabilitar aceleração de hardware
options.add_experimental_option("detach", True)  # Manter o navegador aberto

# Configuração para evitar perguntas antes de baixar arquivos
preferences = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", preferences)

# Inicializar o navegador
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    # Acessar o site OKTOR
    driver.get("COLOQUE O SITE")

    # Espera o campo de email estar presente e inserir o email
    email = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "email")))
    email.send_keys("USÚARIO")

    # Espera o campo de senha estar presente e inserir a senha
    senha = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, "senha")))
    senha.send_keys("SENHA")

    # Clicar no botão de login
    btn_login = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'btn')]")))
    btn_login.click()

    # Espera o menu lateral estar presente após o login
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.CLASS_NAME, "main-header")))

    print("Login realizado com sucesso!")

    # **Clicar no botão do menu lateral (fa-bars), caso necessário**
    try:
        menu_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//nav[contains(@class, 'main-header')]//i[contains(@class, 'fa-bars')]")))
        menu_button.click()
        print("Menu lateral expandido!")
    except Exception as menu_error:
        print(f"Erro ao tentar expandir o menu lateral: {menu_error}")

    # **Clicar na aba 'Relatórios' no menu lateral**
    try:
        relatorios_menu = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, "//p[normalize-space()='Relatórios']"))
        )
        # Scroll para garantir visibilidade
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", relatorios_menu)
        time.sleep(1)  # Pequena pausa para garantir que o scroll foi concluído
        relatorios_menu.click()
        print("Aba 'Relatórios' acessada!")
    except Exception as relatorios_error:
        print(f"Erro ao tentar acessar a aba 'Relatórios': {relatorios_error}")

    # **Aguardar a opção 'Detalhamento de consumo' aparecer e clicar**
    try:
        detalhamento_consumo = WebDriverWait(driver, 60).until(
            EC.element_to_be_clickable((By.XPATH, "//p[normalize-space()='Detalhamento de consumo']"))
        )
        detalhamento_consumo.click()
        print("Acesso à aba 'Detalhamento de consumo' realizado com sucesso!")
    except Exception as detalhamento_error:
        print(f"Erro ao tentar acessar 'Detalhamento de consumo': {detalhamento_error}")

    # **Aguardar a página recarregar após o clique**
    time.sleep(5)  # Aguarda 5 segundos para o recarregamento da página

    # **Aguardar a página carregar completamente após o recarregamento**
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Detalhamento de consumo')]")))
        print("Página recarregada com sucesso!")
    except Exception as recarregamento_error:
        print(f"Erro ao aguardar o recarregamento da página: {recarregamento_error}")

    # **Sair da seleção de 'Detalhamento de consumo'**
    try:
        # Clicar em um elemento neutro para 'sair' da seleção atual
        elemento_neutro = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@class='main-header']"))  # Substitua pelo XPath de um elemento neutro
        )
        elemento_neutro.click()
        print("Saiu da seleção de 'Detalhamento de consumo'.")
    except Exception as sair_error:
        print(f"Erro ao tentar sair da seleção: {sair_error}")

    # **Selecionar a opção 'Personalizado' no campo de período**
    try:
        periodo_select = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "periodo"))
        )
        periodo_select.click()
        opcao_personalizado = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//option[@value='personalizado']"))
        )
        opcao_personalizado.click()
        print("Opção 'Personalizado' selecionada no período!")
    except Exception as periodo_error:
        print(f"Erro ao selecionar o período: {periodo_error}")

    # **Preencher o campo 'data_inicio' com o primeiro dia do mês**
    try:
        data_inicio_input = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "data_inicio"))
        )
        # Obter o primeiro dia do mês atual
        primeiro_dia_mes = datetime(datetime.now().year, datetime.now().month, 1).strftime("%d/%m/%Y")
        data_inicio_input.clear()
        data_inicio_input.send_keys(primeiro_dia_mes)
        print(f"Campo 'data_inicio' preenchido com o primeiro dia do mês: {primeiro_dia_mes}")
    except Exception as data_inicio_error:
        print(f"Erro ao preencher o campo 'data_inicio': {data_inicio_error}")

    # **Preencher o campo 'data_fim' com o dia atual**
    try:
        data_fim_input = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, "data_fim"))
        )
        # Obter o dia atual
        dia_atual = datetime.now().strftime("%d/%m/%Y")
        data_fim_input.clear()
        data_fim_input.send_keys(dia_atual)
        print(f"Campo 'data_fim' preenchido com o dia de hoje: {dia_atual}")
    except Exception as data_fim_error:
        print(f"Erro ao preencher o campo 'data_fim': {data_fim_error}")

    # **Aguardar o botão 'Buscar' estar presente e clicável**
    try:
        btn_buscar = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and @id='btnSubmit']"))
        )
        btn_buscar.click()
        print("Botão 'Buscar' clicado! A pesquisa foi iniciada.")
    except Exception as buscar_error:
        print(f"Erro ao tentar clicar no botão 'Buscar': {buscar_error}")

    # Pausa para garantir que a pesquisa tenha sido processada
    time.sleep(5)

    # **Aguardar o botão CSV estar presente e clicável**
    try:
        btn_csv = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@class='dt-button buttons-csv buttons-html5 btn btn-warning fa fa-file']"))
        )
        print("Botão CSV encontrado!")

        # Garantir que o botão esteja visível
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_csv)

        # Usar JavaScript para clicar no botão
        driver.execute_script("arguments[0].click();", btn_csv)
        print("Botão 'CSV' clicado! Download iniciado.")
    except Exception as csv_error:
        print(f"Erro ao tentar clicar no botão CSV: {csv_error}")

    # Pausa para o download do arquivo
    time.sleep(10)

    download_folder = str(Path.home() / "Downloads")
    destination = r"CAMINHO PARA SALVAR ARQUIVO"

    # Verificar se a pasta de destino existe, se não, criar
    if not os.path.exists(destination):
        try:
            os.makedirs(destination)
            print(f"Pasta de destino criada: {destination}")
        except Exception as e:
            print(f"Erro ao criar a pasta de destino: {e}")
            raise

    # Listar os arquivos na pasta de downloads
    files = os.listdir(download_folder)
    # Procurar o arquivo CSV mais recente
    downloaded_file = max([f for f in files if f.endswith('.csv')], key=lambda x: os.path.getctime(os.path.join(download_folder, x)))
    
    # Caminho completo do arquivo baixado
    downloaded_file_path = os.path.join(download_folder, downloaded_file)
    
    # Obter o mês e o ano atual
    current_month_year = datetime.now().strftime("%m_%Y")

    # Criar o novo nome para o arquivo como 'CUSTO_OKTOR_{current_month_year}.csv'
    new_file_name = f"CUSTO_OKTOR_{current_month_year}.csv"

    # Caminho de destino com o novo nome
    destino_file_path = os.path.join(destination, new_file_name)
    
    # Verificar permissões antes de mover
    if os.access(destination, os.W_OK):
        try:
            # Verificar se o arquivo já existe no destino e remover se necessário
            if os.path.exists(destino_file_path):
                os.remove(destino_file_path)  # Remove o arquivo com o mesmo nome
                print(f"Arquivo existente removido: {new_file_name}")
            
            # Mover o arquivo para o destino com o novo nome
            shutil.move(downloaded_file_path, destino_file_path)
            print(f"Arquivo {downloaded_file} renomeado e movido para {destino_file_path}.")
        except Exception as move_error:
            print(f"Erro ao mover o arquivo: {move_error}")
    else:
        print(f"Erro: Sem permissão para escrever em {destination}")

except Exception as e:
    print(f"Ocorreu um erro: {e}")

finally:
    # Fechar o navegador
    if driver:
        driver.quit()
        print("Navegador fechado.")
    print("Processo concluído!")
