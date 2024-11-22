
import json
import hashlib
import os
import subprocess
import threading
from datetime import datetime
from pytz import timezone
from flask import Flask, redirect, request

app = Flask(__name__)

# Função para registrar o log da requisição
def log_request():
    # Captura o endereço IP real do cliente, considerando cabeçalhos de proxy
    if request.headers.getlist("X-Forwarded-For"):
        ip = request.headers.getlist("X-Forwarded-For")[0]
    else:
        ip = request.remote_addr

    # Criar um hash do User-Agent para anonimizar o usuário
    user_agent = request.headers.get('User-Agent', 'unknown')
    hash_object = hashlib.md5(user_agent.encode())
    user_hash = hash_object.hexdigest()

    # Formatar timestamp para registro de log
    timestamp = datetime.now(
        timezone('America/Sao_Paulo')).strftime("%Y%m%d_%H%M%S")

    # Construir o caminho do arquivo de log (ajuste o diretório conforme necessário)
    log_file_path = os.path.join(
        'logs',  # Diretório onde os logs serão armazenados
        f'log_{user_hash}_{timestamp}.txt'
    )

    # Criar o registro de log com informações relevantes da requisição
    log = {
        "ip": ip,
        "port": request.environ.get('REMOTE_PORT', 'unknown'),
        "user-agent": user_agent,
        "scheme": request.scheme,
        "protocol": request.environ.get('SERVER_PROTOCOL', 'unknown'),
        "method": request.method,
        "query": request.query_string.decode('utf-8'),
        "accessed_at": timestamp
    }

    # Escrever o log no arquivo
    try:
        os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
        with open(log_file_path, 'w') as fp:
            fp.write(json.dumps(log, indent=4) + "\n")
    except Exception as e:
        app.logger.error(f"Failed to write log: {e}")

@app.route('/')
def index():
    # Registrar as informações de acesso
    log_request()

    # Redirecionamento imediato e forçado para a URL configurada
    return redirect("https://example.com", code=302)  # Substitua pela URL desejada

def run_ngrok():
    # Inicia o ngrok em segundo plano (ajuste o comando conforme necessário)
    ngrok_command = "ngrok start --all"  # Comando genérico para iniciar o ngrok
    subprocess.Popen(ngrok_command, shell=True)

if __name__ == "__main__":
    # Iniciar o ngrok em uma thread separada para não bloquear o Flask
    threading.Thread(target=run_ngrok).start()

    # Iniciar o servidor Flask escutando em todas as interfaces na porta 3030
    app.run(host='0.0.0.0', port=3030, debug=True)
