import pyodbc
import time
import os
from datetime import datetime, timedelta

# Caminhos
query_path = r"LOCAL DA QUERY"
output_dir = r"LOCAL FINAL DE SALVAMENTO"
data_ontem = (datetime.now() - timedelta(days=1)).strftime("%d_%m_%Y")
output_file = os.path.join(output_dir, f"LOGS_CRUZEIRO_{data_ontem}.csv")

# Função para tentar a execução da consulta
def execute_query(cursor, query, retries=3, delay=2):
    for attempt in range(retries):
        try:
            cursor.execute(query)
            return cursor.fetchall(), [column[0] for column in cursor.description]
        except pyodbc.Error as e:
            if '1205' in str(e):  # Deadlock error code
                print(f"Deadlock detected. Retrying... (Attempt {attempt + 1}/{retries})")
                time.sleep(delay)
            else:
                raise e  # Se o erro não for deadlock, re-raise a exceção
    raise Exception("Max retries reached. Deadlock unresolved.")

# Estabelecendo a conexão com o banco de dados
conn = pyodbc.connect('DRIVER=INFO BD'
                      'SERVER=INFO BD'
                      'DATABASE=INFO BD'
                      'UID=INFO BD'
                      'PWD=INFO BD')

# Criando o cursor
cursor = conn.cursor()

# Lendo a query do arquivo
with open(query_path, 'r', encoding='utf-8') as file:
    query = file.read()

# Executando a query e obtendo resultados
data, columns = execute_query(cursor, query)

# Escrevendo os dados no CSV
with open(output_file, 'w', encoding='utf-8', newline='') as f:
    f.write(";".join(columns) + "\n")  # Cabeçalho
    for row in data:
        f.write(";".join(str(item) for item in row) + "\n")

# Fechar a conexão
cursor.close()
conn.close()

print(f"Arquivo salvo em: {output_file}")
