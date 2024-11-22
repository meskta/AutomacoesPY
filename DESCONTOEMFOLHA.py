import os
import pandas as pd
import mysql.connector
from mysql.connector import Error
from datetime import datetime
import logging

# Configuração de logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def create_db_connection(host_name, port, user_name, user_password, db_name):
    """
    Cria uma conexão com o banco de dados MySQL.

    Args:
        host_name (str): Nome do host do banco de dados.
        port (int): Porta do banco de dados.
        user_name (str): Nome do usuário do banco de dados.
        user_password (str): Senha do usuário do banco de dados.
        db_name (str): Nome do banco de dados.

    Returns:
        connection: Objeto de conexão ao banco de dados ou None em caso de falha.
    """
    connection = None
    try:
        connection = mysql.connector.connect(
            host=host_name,
            port=port,
            user=user_name,
            password=user_password,
            database=db_name
        )
        logging.info("Conexão MySQL bem-sucedida")
    except Error as err:
        logging.error(f"Erro de conexão: '{err}'")
    return connection

def execute_query(connection, query):
    """
    Executa uma query no banco de dados.

    Args:
        connection: Conexão ao banco de dados.
        query (str): Query SQL a ser executada.

    Returns:
        tuple: Número de linhas afetadas e um booleano indicando sucesso ou falha.
    """
    cursor = connection.cursor(buffered=True)
    try:
        cursor.execute(query)
        connection.commit()
        return cursor.rowcount, True
    except Error as err:
        logging.error(f"Erro na execução da query: '{err}'")
        return 0, False

def process_excel_file(file_path, connection):
    """
    Processa um arquivo Excel e atualiza registros no banco de dados.

    Args:
        file_path (str): Caminho para o arquivo Excel.
        connection: Conexão ao banco de dados.

    Returns:
        tuple: Total de linhas processadas, atualizações bem-sucedidas, falhas de atualização e lista de CPFs não atualizados.
    """
    try:
        df = pd.read_excel(file_path)
        logging.info(f"Planilha lida com sucesso. Total de linhas: {len(df)}")
        logging.info(f"Colunas na planilha: {df.columns.tolist()}")

        total_linhas = 0
        atualizacoes_bem_sucedidas = 0
        atualizacoes_falhas = 0
        cpfs_nao_atualizados = []

        for index, row in df.iterrows():
            total_linhas += 1
            try:
                cpf = str(row['CPF'])
                matricula = str(int(row['Matrícula']))
                mes_competencia = f"{int(row['Mês Competência']):02d}"
                ano_competencia = str(int(row['Ano Competência']))
                valor_pago = f"{int(row['Valor'] * 100):03d}"
                convenio = f"{int(row['Logo']):03d}"

                update_query = f"""
                UPDATE tb_monetario t
                JOIN (SELECT M.id
                      FROM tb_monetario M
                      INNER JOIN tb_propostas P ON P.id=M.id_proposta
                      INNER JOIN tb_clientes C ON C.id=P.id_cliente
                      JOIN tb_convenios CO ON CO.id=P.id_convenio
                      JOIN tb_logos L ON L.id=CO.id_logo
                      WHERE C.cpf='{cpf}'
                      AND P.matricula LIKE '%{matricula}'
                      AND M.mes_competencia='{mes_competencia}'
                      AND M.ano_competencia='{ano_competencia}'
                      AND M.situacao='A' AND L.cod_logo='{convenio}'
                      LIMIT 1) AS sub ON t.id = sub.id
                SET t.valor_descontado='{valor_pago}';
                """

                rows_affected, success = execute_query(connection, update_query)
                if success and rows_affected > 0:
                    atualizacoes_bem_sucedidas += 1
                    logging.info(f"Linha {index}: Atualização bem-sucedida. Linhas afetadas: {rows_affected}")
                else:
                    atualizacoes_falhas += 1
                    cpfs_nao_atualizados.append(cpf)
                    logging.warning(f"Linha {index}: Falha na atualização.")

            except Exception as e:
                logging.error(f"Erro ao processar linha {index}: {e}")
                atualizacoes_falhas += 1
                cpfs_nao_atualizados.append(cpf)

        return total_linhas, atualizacoes_bem_sucedidas, atualizacoes_falhas, cpfs_nao_atualizados

    except Exception as e:
        logging.error(f"Erro ao ler a planilha: {e}")
        return 0, 0, 0, []

def main():
    """
    Função principal que coordena o fluxo de execução do script: conexão ao banco de dados,
    processamento do arquivo Excel e atualização dos registros com base nos dados fornecidos.
    """
    # Configurações do banco de dados (substitua pelos seus valores de configuração)
    db_config = {
        "host_name": "HOST_EXAMPLE",
        "port": 3306,
        "user_name": "USER_EXAMPLE",
        "user_password": "PASSWORD_EXAMPLE",
        "db_name": "DB_EXAMPLE"
    }

    # Caminho base para os arquivos (substitua pelo seu caminho)
    base_path = r'CAMINHO\PARA\ARQUIVOS'

    # Data atual
    hoje = datetime.now()
    ano = hoje.strftime('%Y')
    mes = hoje.strftime('%m')
    dia = hoje.strftime('%d')

    # Monta o caminho completo
    caminho_diretorio = os.path.join(base_path, ano, mes, f"{ano}.{mes}.{dia}")
    logging.info(f"Tentando acessar o diretório: {caminho_diretorio}")

    # Verifica se o diretório existe
    if not os.path.exists(caminho_diretorio):
        logging.error(f"O diretório não existe: {caminho_diretorio}")
        return

    # Procura pelo arquivo mais recente que começa com "Baixas Orbitall" e termina com ".xlsx"
    arquivos = [f for f in os.listdir(caminho_diretorio) if f.startswith("Baixas Orbitall") and f.endswith(".xlsx")]
    if not arquivos:
        logging.error("Nenhum arquivo adequado encontrado.")
        return

    arquivo_mais_recente = max(arquivos, key=lambda f: os.path.getmtime(os.path.join(caminho_diretorio, f)))
    caminho_completo = os.path.join(caminho_diretorio, arquivo_mais_recente)
    logging.info(f"Arquivo encontrado: {arquivo_mais_recente}")

    # Conecta ao banco de dados
    connection = create_db_connection(**db_config)
    if not connection:
        logging.error("Não foi possível conectar ao banco de dados. Encerrando o script.")
        return

    try:
        total_linhas, atualizacoes_bem_sucedidas, atualizacoes_falhas, cpfs_nao_atualizados = process_excel_file(caminho_completo, connection)

        # Relatório final
        print("\nRelatório de Importação:")
        print(f"Total de linhas processadas: {total_linhas}")
        print(f"Atualizações bem-sucedidas: {atualizacoes_bem_sucedidas}")
        print(f"Atualizações com falha: {atualizacoes_falhas}")

        if atualizacoes_bem_sucedidas > 0:
            print("Importação parcial ou total realizada com sucesso.")
        elif total_linhas > 0:
            print("Falha na importação. Nenhum registro foi atualizado.")
        else:
            print("Nenhum dado foi processado.")

        print("\nCPFs não atualizados:")
        for cpf in cpfs_nao_atualizados:
            print(cpf)

    finally:
        if connection:
            connection.close()
            logging.info("Conexão MySQL fechada")

    print("Processamento concluído.")

if __name__ == "__main__":
    main()
