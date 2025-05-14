import os
from datetime import datetime, timedelta
import pandas as pd
from openpyxl import load_workbook
import shutil
import smtplib
import imaplib
import email
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import getpass  # Para obter a senha de forma segura

# Configurações de diretórios - substitua pelos caminhos reais em produção
INPUT_DIR = r"path/to/input/directory"
OUTPUT_DIR = r"path/to/output/directory"
HISTORICO_DIR = r"path/to/history/directory"
CONTROL_FILE = r"path/to/control/file.xlsx"

MOVER_ARQUIVO_ORIGINAL = False
# Configurações de Email - substitua pelos dados reais
SERVIDOR_SMTP = "smtp.example.com"
PORTA_SMTP = 465
SERVIDOR_IMAP = "imap.example.com"
PORTA_IMAP = 993

# Configurações de Email padrão - substitua pelos dados reais
EMAIL_PADRAO = "notification@example.com"
SENHA_PADRAO = "your_password_here"  # Não é recomendado armazenar senhas diretamente no código

data_atual = datetime.now()


def verificar_acesso_rede():
    diretorios = [INPUT_DIR, OUTPUT_DIR, HISTORICO_DIR]
    for diretorio in diretorios:
        try:
            if not os.path.exists(diretorio):
                print(f"Aviso: Diretório não encontrado: {diretorio}")
                if diretorio == HISTORICO_DIR:
                    try:
                        os.makedirs(diretorio, exist_ok=True)
                        print(f"Diretório de histórico criado: {diretorio}")
                    except Exception as e:
                        print(f"Erro ao criar diretório de histórico: {e}")
            else:
                # Verificar se podemos listar o diretório
                os.listdir(diretorio)
                print(f"Acesso confirmado ao diretório: {diretorio}")
        except Exception as e:
            raise Exception(f"Erro ao acessar o diretório {diretorio}: {e}")


def encontrar_arquivos_txn(diretorio):
    """
    Função especializada para encontrar arquivos TXN no diretório especificado.
    Implementa várias estratégias de busca.
    """
    print(f"Iniciando busca por arquivos TXN em: {diretorio}")
    
    # Listar todos os arquivos no diretório
    try:
        todos_arquivos = os.listdir(diretorio)
        arquivos_excel = [f for f in todos_arquivos if f.lower().endswith(('.xlsx', '.xls'))]
        
        print(f"Total de arquivos Excel encontrados: {len(arquivos_excel)}")
        if len(arquivos_excel) > 0:
            print("Primeiros 5 arquivos Excel encontrados:")
            for i, arquivo in enumerate(arquivos_excel[:5]):
                print(f"  {i+1}. {arquivo}")
        
        # Estratégia 1: Buscar pelo padrão exato TXN_AAAAMMDD
        data_hoje = data_atual.strftime('%Y%m%d')
        padrao_hoje = f"TXN_{data_hoje}"
        
        arquivos_txn_hoje = [f for f in arquivos_excel if padrao_hoje.upper() in f.upper()]
        if arquivos_txn_hoje:
            print(f"Encontrados {len(arquivos_txn_hoje)} arquivos com o padrão TXN_{data_hoje}:")
            for arquivo in arquivos_txn_hoje:
                print(f"  - {arquivo}")
            return arquivos_txn_hoje
        
        # Estratégia 2: Buscar por arquivos TXN de datas recentes (últimos 7 dias)
        arquivos_txn_recentes = []
        for i in range(7):  # Verificar últimos 7 dias
            data_verificar = data_atual - timedelta(days=i)
            data_str = data_verificar.strftime('%Y%m%d')
            padrao = f"TXN_{data_str}"
            
            for arquivo in arquivos_excel:
                if padrao.upper() in arquivo.upper():
                    arquivos_txn_recentes.append(arquivo)
        
        if arquivos_txn_recentes:
            print(f"Encontrados {len(arquivos_txn_recentes)} arquivos TXN de datas recentes:")
            for arquivo in arquivos_txn_recentes:
                print(f"  - {arquivo}")
            return arquivos_txn_recentes
        
        # Estratégia 3: Buscar qualquer arquivo com "TXN" no nome
        arquivos_txn_geral = [f for f in arquivos_excel if "TXN" in f.upper()]
        if arquivos_txn_geral:
            print(f"Encontrados {len(arquivos_txn_geral)} arquivos com 'TXN' no nome:")
            for arquivo in arquivos_txn_geral:
                print(f"  - {arquivo}")
            return arquivos_txn_geral
        
        # Estratégia 4: Buscar por padrão usando expressão regular
        padrao_regex = re.compile(r'TXN[_\s-]?\d{8}', re.IGNORECASE)
        arquivos_txn_regex = [f for f in arquivos_excel if padrao_regex.search(f)]
        if arquivos_txn_regex:
            print(f"Encontrados {len(arquivos_txn_regex)} arquivos com padrão regex TXN_AAAAMMDD:")
            for arquivo in arquivos_txn_regex:
                print(f"  - {arquivo}")
            return arquivos_txn_regex
        
        # Se nenhum arquivo for encontrado, retornar lista vazia
        print("Nenhum arquivo TXN encontrado após todas as estratégias de busca.")
        return []
        
    except Exception as e:
        print(f"Erro ao listar arquivos no diretório {diretorio}: {e}")
        return []


def get_next_batch_number():
    try:
        wb = load_workbook(CONTROL_FILE)
        ws = wb.active
        last_row = ws.max_row
        last_batch = ws.cell(row=last_row, column=2).value
        return int(last_batch) + 1
    except Exception as e:
        print(
            f"Erro ao ler o número do lote: {e}. Verifique se o arquivo de controle está acessível.")
        # Valor padrão para fallback - ajuste conforme necessário
        return 100


def update_control_file(batch_number, successful_records):
    try:
        wb = load_workbook(CONTROL_FILE)
        ws = wb.active
        new_row = [datetime.now().strftime('%d/%m/%Y'),
                   f'{batch_number:06d}', successful_records]
        ws.append(new_row)
        wb.save(CONTROL_FILE)
        print(
            f"Planilha de controle atualizada com sucesso. Novo lote: {batch_number:06d}, Registros inseridos: {successful_records}")
    except Exception as e:
        print(
            f"Erro ao atualizar a planilha de controle: {e}. Verifique se o arquivo está acessível.")


def create_header(batch_number: int) -> str:
    current_datetime = datetime.now().strftime('%Y%m%d%H%M%S')
    return f"000000{batch_number:03d}0000000{' ' * 40}A{current_datetime}M00000000{' ' * 412}00000002"


def create_detail_record(card_number: str, txn_code: str, value: float, date: str, sequence: int) -> str:
    formatted_card = f"{int(card_number):016d}"
    formatted_txn = f"{int(txn_code):04d}"
    formatted_value = f"{int(value * 100):017d}"
    formatted_date = datetime.strptime(
        date, '%Y-%m-%d').strftime('%Y%m%d') + "163000"

    return (
        f"1{'0' * 26}"
        f"{formatted_card}"
        f"{'0' * 7}"
        f"{formatted_txn}{formatted_txn}986"
        f"{formatted_value}"
        f"2{'0' * 17}6"
        f"{formatted_date}"
        f"{' ' * 21}"
        f"{'0' * 5}"
        f"{' ' * 79}"
        f"{'0' * 6}"
        f"{' ' * 16}"
        f"{'0' * 14}"
        f"{' ' * 23}"
        f"{'0' * 37}2"
        f"{'0' * 17}2"
        f"{'0' * 17}2"
        f"{'0' * 12}2 "
        f"{'0' * 71}"
        f"{' ' * 15}"
        f"{'0' * 8}"
        f"{' ' * 26}"
        f"{'0' * 5}2   "
        f"{sequence:08d}"
    )


def create_trailer(batch_number: int, total_records: int, total_value: int) -> str:
    current_datetime = datetime.now().strftime('%Y%m%d%H%M%S')
    return (
        f"9{'0' * 5}2{'0' * 11}"
        f"{' ' * 40}"
        f"A{current_datetime}M{current_datetime}"
        f"{total_records:08d}"
        f"{total_value:017d}"
        f"2{' ' * 378}"
        f"{'0' * 7}2"
    )


def generate_file(df: pd.DataFrame, output_path: str, batch_number: int) -> int:
    try:
        total_value = 0
        successful_records = 0
        with open(output_path, 'w') as f:
            header = create_header(batch_number)
            f.write(header + '\n')

            for i, row in df.iterrows():
                try:
                    card_number = row['NUMERO CARTÃO']
                    txn_code = row['TXN']
                    value = row['VALOR']
                    date = row['DATA DE ENVIO'].strftime('%Y-%m-%d')

                    detail_record = create_detail_record(
                        str(card_number), str(txn_code), value, date, i + 1)
                    f.write(detail_record + '\n')
                    total_value += int(value * 100)
                    successful_records += 1
                except Exception as e:
                    print(f"Erro ao processar registro {i + 1}: {e}")

            trailer = create_trailer(
                batch_number, successful_records, total_value)
            f.write(trailer + '\n')

        print(f"Arquivo gerado com sucesso: {output_path}")
        return successful_records
    except Exception as e:
        raise IOError(f"Erro ao gerar o arquivo: {e}")


def enviar_email_primario(remetente, destinatarios, assunto, corpo, senha, anexos=None):
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = ", ".join(destinatarios)
    mensagem['Subject'] = assunto
    
    mensagem.attach(MIMEText(corpo, 'html'))
    
    if anexos:
        for arquivo in anexos:
            if os.path.isfile(arquivo):
                with open(arquivo, 'rb') as f:
                    parte = MIMEApplication(f.read(), Name=os.path.basename(arquivo))
                parte['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo)}"'
                mensagem.attach(parte)
    
    try:
        # Usando servidor SMTP primário
        servidor = smtplib.SMTP("smtp.office365.com", 587)
        servidor.ehlo()
        servidor.starttls()  # Habilitando criptografia
        servidor.login(remetente, senha)
        
        # Enviando email
        texto = mensagem.as_string()
        servidor.sendmail(remetente, destinatarios, texto)
        print("Email enviado com sucesso pelo servidor primário!")
        
    except Exception as e:
        print(f"Erro ao enviar email: {e}")
        raise
    
    finally:
        if 'servidor' in locals():
            servidor.quit()


def enviar_email_secundario(remetente, destinatarios, assunto, corpo, senha, anexos=None):
    try:
        with smtplib.SMTP_SSL(SERVIDOR_SMTP, PORTA_SMTP) as servidor:
            servidor.login(remetente, senha)
            mensagem = MIMEMultipart()
            mensagem['From'] = remetente
            mensagem['To'] = ", ".join(destinatarios)
            mensagem['Subject'] = assunto

            mensagem.attach(MIMEText(corpo, 'html'))

            if anexos:
                for arquivo in anexos:
                    if os.path.isfile(arquivo):
                        with open(arquivo, 'rb') as f:
                            parte = MIMEApplication(
                                f.read(), Name=os.path.basename(arquivo))
                        parte['Content-Disposition'] = f'attachment; filename="{os.path.basename(arquivo)}"'
                        mensagem.attach(parte)

            texto = mensagem.as_string()
            servidor.sendmail(remetente, destinatarios, texto)
            print("Email enviado com sucesso pelo servidor secundário!")

    except Exception as e:
        print(f"Erro ao enviar email: {e}")
        raise


def ler_emails(usuario, senha, pasta='INBOX', quantidade=5):
    try:
        # Conectando ao servidor IMAP
        with imaplib.IMAP4_SSL(SERVIDOR_IMAP, PORTA_IMAP) as mail:
            mail.login(usuario, senha)
            mail.select(pasta)

            status, mensagens = mail.search(None, 'ALL')
            ids_mensagens = mensagens[0].split()
            ids_para_ler = ids_mensagens[-quantidade:] if len(
                ids_mensagens) > quantidade else ids_mensagens

            print(
                f"Lendo os {len(ids_para_ler)} emails mais recentes da pasta {pasta}:")

            for num in reversed(ids_para_ler):
                status, dados = mail.fetch(num, '(RFC822)')
                email_bruto = dados[0][1]
                mensagem = email.message_from_bytes(email_bruto)

                de = mensagem['From']
                assunto = mensagem['Subject']
                data = mensagem['Date']

                print(f"\nDe: {de}")
                print(f"Assunto: {assunto}")
                print(f"Data: {data}")

                if mensagem.is_multipart():
                    for parte in mensagem.walk():
                        tipo_conteudo = parte.get_content_type()
                        if tipo_conteudo == 'text/plain':
                            corpo = parte.get_payload(decode=True).decode()
                            print(f"Corpo: {corpo[:100]}...")
                            break
                else:
                    corpo = mensagem.get_payload(decode=True).decode()
                    print(f"Corpo: {corpo[:100]}...")

    except Exception as e:
        print(f"Erro ao ler emails: {e}")
        raise


def main():
    batch_number = 0  # Inicializar a variável para evitar erro no bloco de exceção
    try:
        print(f"Iniciando processamento em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        
        print("Verificando acesso aos diretórios de rede...")
        verificar_acesso_rede()
        
        if not os.path.exists(HISTORICO_DIR):
            os.makedirs(HISTORICO_DIR, exist_ok=True)
            print(f"Diretório de histórico criado: {HISTORICO_DIR}")

        print(f"Buscando arquivos em {INPUT_DIR}...")
        
        # Usar a função especializada para encontrar arquivos TXN
        arquivos_entrada = encontrar_arquivos_txn(INPUT_DIR)

        if not arquivos_entrada:
            mensagem = "Nenhum arquivo de entrada TXN encontrado."
            print(mensagem)
            
            # Diagnóstico adicional
            print("\nDiagnóstico de diretório:")
            todos_arquivos = os.listdir(INPUT_DIR)
            print(f"Total de arquivos no diretório: {len(todos_arquivos)}")
            if len(todos_arquivos) > 0:
                print("Primeiros 10 arquivos no diretório:")
                for i, arquivo in enumerate(todos_arquivos[:10]):
                    print(f"  {i+1}. {arquivo}")
            
            remetente = EMAIL_PADRAO
            senha = SENHA_PADRAO
            destinatarios = [EMAIL_PADRAO]
            assunto = "Informação - Processamento de Transações OFFLINE"
            
            corpo = f"""
            <html>
            <body>
                <h2>Informação de Processamento</h2>
                <p>{mensagem}</p>
                <p>Data e hora da verificação: {data_atual.strftime('%d/%m/%Y %H:%M:%S')}</p>
                <p>Este é um email automático, por favor não responda.</p>
            </body>
            </html>
            """
            
            try:
                enviar_email_primario(remetente, destinatarios, assunto, corpo, senha)
            except Exception as email_error:
                print(f"Erro ao enviar email via servidor primário: {email_error}")
                try:
                    enviar_email_secundario(remetente, destinatarios, assunto, corpo, senha)
                except Exception as secondary_error:
                    print(f"Erro ao enviar email via servidor secundário: {secondary_error}")
            return

        print(f"Encontrados {len(arquivos_entrada)} arquivos para processamento.")

        for arquivo in arquivos_entrada:
            INPUT_FILE = os.path.join(INPUT_DIR, arquivo)
            print(f"Processando arquivo: {INPUT_FILE}")

            if not os.path.exists(INPUT_FILE):
                raise FileNotFoundError(
                    f"Arquivo de entrada não encontrado: {INPUT_FILE}")

            df = pd.read_excel(INPUT_FILE)
            print(f"Planilha carregada com {len(df)} registros.")

            # Converte a coluna 'DATA DE ENVIO' para datetime
            df['DATA DE ENVIO'] = pd.to_datetime(
                df['DATA DE ENVIO'], errors='coerce')

            required_columns = ['NUMERO CARTÃO',
                                'TXN', 'VALOR', 'DATA DE ENVIO']
            if not all(col in df.columns for col in required_columns):
                raise ValueError(
                    "Colunas necessárias não encontradas na planilha.")

            batch_number = get_next_batch_number()
            print(f"Número do lote: {batch_number}")

            current_datetime = datetime.now()
            output_filename = f"COMPANY_OFFLINE_{current_datetime.strftime('%d%m%Y_%H%M')}.txt"
            output_path = os.path.join(OUTPUT_DIR, output_filename)

            successful_records = generate_file(df, output_path, batch_number)

            update_control_file(batch_number, successful_records)

            remetente = EMAIL_PADRAO
            senha = SENHA_PADRAO
            destinatarios = [EMAIL_PADRAO]
            assunto = f"Processamento de Transações OFFLINE_{batch_number:06d}"
            
            corpo = f"""
            <html>
            <body>
                <h2>Processamento de Transações Concluído</h2>
                <p>O processamento do lote OFFLINE_<b>{batch_number:06d}</b> foi concluído com sucesso.</p>
                <p><b>Detalhes do processamento:</b></p>
                <ul>
                    <li>Arquivo processado: {arquivo}</li>
                    <li>Arquivo gerado: {output_filename}</li>
                    <li>Total de registros processados: {successful_records}</li>
                    <li>Total de registros na planilha original: {len(df)}</li>
                    <li>Data e hora do processamento: {current_datetime.strftime('%d/%m/%Y %H:%M:%S')}</li>
                </ul>
                <p>Este é um email automático, por favor não responda.</p>
            </body>
            </html>
            """
            
            anexos = [output_path]
            
            try:
                enviar_email_primario(remetente, destinatarios, assunto, corpo, senha, anexos)
            except Exception as email_error:
                print(f"Erro ao enviar email via servidor primário: {email_error}")
                # Tentar enviar via servidor secundário como fallback
                try:
                    print("Tentando enviar email via servidor secundário como alternativa...")
                    enviar_email_secundario(remetente, destinatarios, assunto, corpo, senha, anexos)
                except Exception as secondary_error:
                    print(f"Erro ao enviar email via servidor secundário: {secondary_error}")

            print(
                f"Total de registros processados com sucesso: {successful_records}")
            print(f"Total de registros na planilha original: {len(df)}")
            if successful_records < len(df):
                print(
                    f"Atenção: {len(df) - successful_records} registros não foram processados.")

            # Mover ou copiar o arquivo para o histórico
            arquivo_destino = os.path.join(HISTORICO_DIR, arquivo)
            if MOVER_ARQUIVO_ORIGINAL:
                shutil.move(INPUT_FILE, arquivo_destino)
                print(
                    f"Arquivo original movido para o histórico: {arquivo_destino}")
            else:
                shutil.copy2(INPUT_FILE, arquivo_destino)
                print(f"Arquivo original mantido em: {INPUT_FILE}")
                print(f"Cópia do arquivo salva em: {arquivo_destino}")

        print(f"Processamento concluído em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    except Exception as e:
        erro_msg = f"Erro: {e}"
        print(erro_msg)

        # Enviar email de erro
        try:
            remetente = EMAIL_PADRAO
            senha = SENHA_PADRAO
            destinatarios = [EMAIL_PADRAO]
            assunto = "ERRO - Processamento de Transações OFFLINE"

            corpo = f"""
            <html>
            <body>
                <h2>Log de processamento - OFFLINE_{batch_number:06d}</h2>
                <p>Resultado do processamento:</p>
                <p style="color: red; font-weight: bold;">{str(e)}</p>
                <p>Data e hora do erro: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
                <p>Este é um email automático, por favor não responda.</p>
            </body>
            </html>
            """

            try:
                enviar_email_primario(remetente, destinatarios, assunto, corpo, senha)
            except Exception as email_error:
                print(f"Erro ao enviar email de erro via servidor primário: {email_error}")
                # Tentar enviar via servidor secundário como fallback
                try:
                    print("Tentando enviar email de erro via servidor secundário como alternativa...")
                    enviar_email_secundario(remetente, destinatarios, assunto, corpo, senha)
                except Exception as secondary_error:
                    print(f"Erro ao enviar email de erro via servidor secundário: {secondary_error}")
        except Exception as email_error:
            print(
                f"Erro ao enviar email de notificação de erro: {email_error}")


if __name__ == "__main__":
    main()
