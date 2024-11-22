import os
import pandas as pd
import mysql.connector
from mysql.connector import Error
import logging

# Configuração de logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def create_db_connection(host_name, port, user_name, user_password, db_name):
    """
    Estabelece uma conexão com o banco de dados MySQL.

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
            database=db_name,
            connect_timeout=20
        )
        logging.info("Conexão MySQL bem-sucedida")
    except Error as err:
        logging.error(f"Erro de conexão: '{err}'")
    return connection

def execute_query(connection, query):
    """
    Executa uma query de atualização no banco de dados.

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

def execute_select(connection, query):
    """
    Executa uma query de seleção no banco de dados.

    Args:
        connection: Conexão ao banco de dados.
        query (str): Query SQL a ser executada.

    Returns:
        list: Resultado da query ou None em caso de falha.
    """
    cursor = connection.cursor(buffered=True)
    try:
        cursor.execute(query)
        result = cursor.fetchall()
        return result
    except Error as err:
        logging.error(f"Erro na execução do SELECT: '{err}'")
        return None

def log_audit(connection, numero_conta, matricula, cpf_cliente, logo, action):
    """
    Registra uma ação de auditoria no banco de dados.

    Args:
        connection: Conexão ao banco de dados.
        numero_conta (str): Número da conta.
        matricula (str): Matrícula associada.
        cpf_cliente (str): CPF do cliente.
        logo (str): Código do logo.
        action (str): Ação realizada.
    """
    audit_query = f"""
    INSERT INTO tb_propostas_audit (numero_conta, matricula, cpf_cliente, logo, action)
    VALUES ('{numero_conta}', '{matricula}', '{cpf_cliente}', '{logo}', '{action}');
    """
    execute_query(connection, audit_query)

def update_aluno_cpf(connection, numero_conta, new_cpf):
    """
    Atualiza o CPF de um aluno no banco de dados.

    Args:
        connection: Conexão ao banco de dados.
        numero_conta (str): Número da conta.
        new_cpf (str): Novo CPF a ser atualizado.
    """
    select_query = f"""
    SELECT matricula, cpf_cliente, logo FROM tb_propostas
    WHERE numero_conta='{numero_conta}';
    """
    result = execute_select(connection, select_query)

    if result:
        matricula, old_cpf, logo = result[0]
        update_query = f"""
        UPDATE tb_propostas
        SET cpf_cliente='{new_cpf}'
        WHERE numero_conta='{numero_conta}';
        """
        rows_affected, success = execute_query(connection, update_query)

        if success and rows_affected > 0:
            log_audit(connection, numero_conta, matricula, old_cpf, logo, 'update')
            logging.info(f"Atualização bem-sucedida para número de conta {numero_conta}. CPF alterado de {old_cpf} para {new_cpf}")
        else:
            logging.warning(f"Falha na atualização para número de conta {numero_conta}.")
    else:
        logging.warning(f"Registro não encontrado para número de conta {numero_conta}.")

def revert_aluno_cpf(connection, numero_conta, version_steps_back):
    """
    Reverte o CPF de um aluno para uma versão anterior.

    Args:
        connection: Conexão ao banco de dados.
        numero_conta (str): Número da conta.
        version_steps_back (int): Número de versões a reverter.
    """
    audit_query = f"""
    SELECT matricula, cpf_cliente, logo FROM tb_propostas_audit
    WHERE numero_conta='{numero_conta}'
    ORDER BY timestamp DESC
    LIMIT 1 OFFSET {version_steps_back};
    """
    result = execute_select(connection, audit_query)

    if result:
        matricula, target_cpf, logo = result[0]
        update_query = f"""
        UPDATE tb_propostas
        SET cpf_cliente='{target_cpf}'
        WHERE numero_conta='{numero_conta}';
        """
        rows_affected, success = execute_query(connection, update_query)

        if success and rows_affected > 0:
            log_audit(connection, numero_conta, matricula, target_cpf, logo, 'revert')
            logging.info(f"Reversão bem-sucedida para número de conta {numero_conta}. CPF revertido para {target_cpf}")
        else:
            logging.warning(f"Falha na reversão para número de conta {numero_conta}.")
    else:
        logging.warning(f"Não foi encontrada uma versão anterior suficiente para reverter para o número de conta {numero_conta}.")

def format_account_number(account_number):
    """
    Formata o número da conta para um formato padrão.

    Args:
        account_number (str): Número da conta a ser formatado.

    Returns:
        str: Número da conta formatado.
    """
    account_str = str(int(float(account_number)))
    return account_str.zfill(19)

def format_matricula(matricula):
    """
    Formata a matrícula para um formato padrão.

    Args:
        matricula (str): Matrícula a ser formatada.

    Returns:
        str: Matrícula formatada.
    """
    return str(int(float(matricula)))

def process_excel_file(file_path, connection):
    """
    Processa um arquivo Excel e atualiza registros no banco de dados.

    Args:
        file_path (str): Caminho para o arquivo Excel.
        connection: Conexão ao banco de dados.

    Returns:
        tuple: Total de linhas processadas, atualizações bem-sucedidas, falhas de atualização, CPFs não atualizados e CPFs duplicados.
    """
    try:
        df = pd.read_excel(file_path)
        logging.info(f"Planilha lida com sucesso. Total de linhas: {len(df)}")
        logging.info(f"Colunas na planilha: {df.columns.tolist()}")

        total_linhas = 0
        atualizacoes_bem_sucedidas = 0
        atualizacoes_falhas = 0
        cpfs_nao_atualizados = []
        cpfs_duplicados = []

        for index, row in df.iterrows():
            total_linhas += 1
            try:
                matricula = format_matricula(row['MATRICULA'])
                numero_conta = format_account_number(row['NUM_CONTA'])
                documento_cpf = str(row['CPF_CLIENTE'])
                logo = str(row['LOGO'])

                # Verifica duplicidade de CPF com matrículas diferentes no mesmo LOGO
                duplicidade_query = f"""
                SELECT CPF_CLIENTE
                FROM tb_propostas
                WHERE LOGO='{logo}'
                GROUP BY CPF_CLIENTE
                HAVING COUNT(DISTINCT matricula) > 1;
                """
                duplicidades = execute_select(connection, duplicidade_query)

                if duplicidades:
                    logging.warning(f"Linha {index}: CPF duplicado com matrículas diferentes encontrado para o LOGO {logo}")
                    cpfs_duplicados.append(documento_cpf)
                    continue

                # Verifica se o registro existe
                select_query = f"""
                SELECT matricula FROM tb_propostas
                WHERE numero_conta='{numero_conta}';
                """
                result = execute_select(connection, select_query)

                if not result:
                    logging.warning(f"Linha {index}: Registro não encontrado para número de conta {numero_conta}")
                    atualizacoes_falhas += 1
                    cpfs_nao_atualizados.append(documento_cpf)
                    continue

                update_aluno_cpf(connection, numero_conta, documento_cpf)

            except Exception as e:
                logging.error(f"Erro ao processar linha {index}: {e}")
                atualizacoes_falhas += 1
                cpfs_nao_atualizados.append(documento_cpf)

        return total_linhas, atualizacoes_bem_sucedidas, atualizacoes_falhas, cpfs_nao_atualizados, cpfs_duplicados

    except Exception as e:
        logging.error(f"Erro ao ler a planilha: {e}")
        return 0, 0, 0, [], []

def main():
    """
    Função principal que coordena a execução do script: verifica o arquivo Excel, conecta ao banco de dados,
    processa o arquivo, atualiza registros e gera um relatório de operações realizadas.
    """
    # Configurações do banco de dados (substitua pelos seus valores de configuração)
    db_config = {
        "host_name": "HOST_EXAMPLE",
        "port": 3306,
        "user_name": "USER_EXAMPLE",
        "user_password": "PASSWORD_EXAMPLE",
        "db_name": "DB_EXAMPLE"
    }

    # Caminho direto para o arquivo (substitua pelo seu caminho)
    caminho_arquivo = r"CAMINHO\PARA\ARQUIVO\EXCEL.xlsx"

    # Verifica se o arquivo existe
    if not os.path.exists(caminho_arquivo):
        logging.error(f"O arquivo não existe: {caminho_arquivo}")
        return

    logging.info(f"Arquivo encontrado: {caminho_arquivo}")

    # Conecta ao banco de dados
    connection = create_db_connection(**db_config)
    if not connection:
        logging.error("Não foi possível conectar ao banco de dados. Encerrando o script.")
        return

    try:
        total_linhas, atualizacoes_bem_sucedidas, atualizacoes_falhas, cpfs_nao_atualizados, cpfs_duplicados = process_excel_file(
            caminho_arquivo, connection)

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

        print("\nCPFs com duplicidade de matrículas:")
        for cpf in cpfs_duplicados:
            print(cpf)

    finally:
        if connection:
            connection.close()
            logging.info("Conexão MySQL fechada")

    print("Processamento concluído.")

if __name__ == "__main__":
    main()
