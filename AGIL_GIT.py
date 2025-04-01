from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import os
import shutil
from datetime import datetime

# Configuração do Selenium
options = webdriver.ChromeOptions()
options.add_argument("--disable-gpu")
preferences = {
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", preferences)
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Acessar o site
driver.get("COLOQUE O SITE")
driver.maximize_window()

# Fazer login
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtNumero")))
usuario = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtNumero")
senha = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtSenha")
btn_login = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnLogar")
usuario.send_keys("USÚARIO")
senha.send_keys("SENHA") 
btn_login.click()

# Acessar Demonstrativo
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "Demonstrativo")))
demonstrativo = driver.find_element(By.LINK_TEXT, "Demonstrativo")
demonstrativo.click()

# Clicar no botão de play correto dependendo do dia do mês
dia_atual = datetime.now().day

if dia_atual == 28:
    xpath_botao_play = "(//img[contains(@src, 'images/play.png')])[3]"  # Terceiro botão de play
else:
    xpath_botao_play = "(//img[contains(@src, 'images/play.png')])[2]"  # Segundo botão de play

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_botao_play)))
driver.find_element(By.XPATH, xpath_botao_play).click()

# Exportar CSV
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_btnExportar")))
driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExportar").click()

# Espera o download
download_destino = os.path.join(os.path.expanduser("~"), "Downloads")
timeout = 480
inicio_espera = time.time()
caminho_arquivo = None
while time.time() - inicio_espera < timeout:
    arquivos = os.listdir(download_destino)
    for arquivo in arquivos:
        if arquivo.endswith(".csv"):
            caminho_arquivo = os.path.join(download_destino, arquivo)
            print(f"Arquivo encontrado: {arquivo}")
            break
    else:
        time.sleep(1)
        continue
    break

if not caminho_arquivo:
    print("Arquivo não foi baixado a tempo.")
    driver.quit()
    exit()

df = pd.read_csv(caminho_arquivo, encoding='ISO-8859-1', sep=';')

# Processar arquivos CSV no diretório
with open(caminho_arquivo, 'r', encoding='latin-1') as file:
    linhas = file.readlines()
    linhas[0] = linhas[0].rstrip('\n') + ';'
    linhas[0] = linhas[0].rstrip(';;\n') + ';'

with open(caminho_arquivo, 'w', encoding='latin-1') as file:
    file.writelines(linhas)

df = pd.read_csv(caminho_arquivo, encoding='ISO-8859-1', sep=';')

# Transformação de dados
if 'DATA' in df.columns:
    # Converter a coluna 'DATA' para o formato correto de data
    df['DATA'] = pd.to_datetime(df['DATA'], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')

if 'VALOR' in df.columns:
    # Converter os valores da coluna 'VALOR' para formato numérico
    df['VALOR'] = df['VALOR'].astype(str).str.replace(',', '.', regex=False).astype(float)
    df['VALOR'] = df['VALOR'].fillna(0)

# Agora vamos agrupar os dados pela 'DATA' e somar os valores da coluna 'VALOR'
if 'DATA' in df.columns and 'VALOR' in df.columns:
    df_aggregated = df.groupby('DATA', as_index=False)['VALOR'].sum()
    df = df_aggregated

# Substituir o ponto por vírgula na parte decimal e formatar com ponto como separador de milhar
if 'VALOR' in df.columns:
    df['VALOR'] = df['VALOR'].apply(lambda x: f"{x:,.2f}".replace('.', '_').replace(',', '.').replace('_', ','))

# Exibir o DataFrame agregado
print(df)

# Definir destino final
diretorio_destino_final = r"CAMINHO PARA SALVAR ARQUIVO PRONTO"
data_atual = datetime.now()

if data_atual.day == 28:
    # No dia 28, o intervalo é do dia 28 do mês anterior até 27 do mês atual
    if data_atual.month == 1:
        inicio_intervalo = datetime(data_atual.year - 1, 12, 28)
    else:
        inicio_intervalo = datetime(data_atual.year, data_atual.month - 1, 28)
    fim_intervalo = datetime(data_atual.year, data_atual.month, 27)
else:
    # A partir do dia 29, o intervalo é do dia 28 do mês atual até 27 do próximo mês
    inicio_intervalo = datetime(data_atual.year, data_atual.month, 28)
    if data_atual.month == 12:
        fim_intervalo = datetime(data_atual.year + 1, 1, 27)
    else:
        fim_intervalo = datetime(data_atual.year, data_atual.month + 1, 27)

# Nome do arquivo ajustado conforme o intervalo definido
novo_nome = f"AGIL_{inicio_intervalo.strftime('%d_%m_%Y')} a {fim_intervalo.strftime('%d_%m_%Y')}_APÓS3S.csv"
destino_final_renomeado = os.path.join(diretorio_destino_final, novo_nome)

df.to_csv(destino_final_renomeado, index=False, sep=';', encoding='ISO-8859-1')

# Mover e renomear o arquivo para o destino final
df.to_csv(destino_final_renomeado, index=False, sep=';', encoding='ISO-8859-1')

# Remover o arquivo da pasta Downloads
if os.path.exists(caminho_arquivo):
    os.remove(caminho_arquivo)
    print(f"Arquivo original removido de {download_destino}")

print(f"Arquivo CSV renomeado para {novo_nome} e movido para {diretorio_destino_final}")

driver.quit()
print("Processo concluído!")