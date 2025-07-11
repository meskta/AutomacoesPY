import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime
import os
import sys
import json
import threading
import time
import subprocess
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd
from pathlib import Path
import schedule
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import traceback

# Variáveis de configuração da API IBM Control Center
ICC_API_BASE_URL = "https://SEU_SERVER_IBM_ICC:PORTA/api/v1"
ICC_USERNAME = "SEU_USUARIO_API"
ICC_PASSWORD = "SUA_SENHA_API"

try:
    import pyodbc
    import sqlalchemy
    from sqlalchemy import create_engine, Column, Integer, String, DateTime, Text, Float, ForeignKey, Boolean
    from sqlalchemy.orm import declarative_base
    from sqlalchemy.orm import sessionmaker, relationship
    from sqlalchemy.sql import func
    SQLSERVER_AVAILABLE = True
    print("✅ Dependências SQL Server carregadas com sucesso!")
except ImportError as e:
    print(f"❌ Erro ao importar dependências SQL Server: {e}")
    print("📥 Execute: pip install pyodbc sqlalchemy")
    SQLSERVER_AVAILABLE = False

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    print("⚠️ psutil não instalado. Monitoramento do sistema desabilitado.")
    print("Execute: pip install psutil")
    PSUTIL_AVAILABLE = False

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    print("⚠️ requests não instalado. Algumas funcionalidades podem estar limitadas.")
    REQUESTS_AVAILABLE = False


class DatabaseConfig:
    DB_SERVER = '10.2.0.200,1434'
    DB_NAME = 'dbaverbadora'
    DB_DRIVER = 'ODBC Driver 17 for SQL Server'

    # Carregar credenciais das variáveis de ambiente
    # Assume que DB_USER virá de uma variável de ambiente chamada 'DB_USER'
    _db_user_env = os.getenv('DB_USER')
    # A senha virá da variável de ambiente que você especificou: 'BANCO_USERAVERBACAO'
    _db_password_env = os.getenv('BANCO_USERAVERBACAO') 

    # Validação e atribuição das credenciais (será executado na importação do script)
    if not _db_user_env:
        print("\n" + "=" * 60)
        print("❌ ERRO: Variável de ambiente DB_USER (usuário do banco) não configurada.")
        print("Certifique-se de que 'DB_USER' está definida no ambiente do sistema.")
        print('Exemplo no Windows CMD (como Administrador): setx DB_USER "seu_usuario" /M') #  exemplo
        print("=" * 60 + "\n")
        raise ValueError("DB_USER não configurado")
    else:
        DB_USER = _db_user_env

    if not _db_password_env:
        print("\n" + "=" * 60)
        print("❌ ERRO: Variável de ambiente BANCO_USERAVERBACAO (senha do banco) não configurada.")
        print("Certifique-se de que 'BANCO_USERAVERBACAO' está definida no ambiente do sistema.")
        print('Exemplo no Windows CMD (como Administrador): setx BANCO_USERAVERBACAO "sua_senha" /M')
        print("=" * 60 + "\n")
        raise ValueError("BANCO_USERAVERBACAO (senha) não configurado")
    else:
        DB_PASSWORD = _db_password_env

    @classmethod
    def get_connection_string(cls):
        # Esta função já usa cls.DB_USER e cls.DB_PASSWORD, que agora virão das variáveis de ambiente
        return f"mssql+pyodbc://{cls.DB_USER}:{cls.DB_PASSWORD}@{cls.DB_SERVER}/{cls.DB_NAME}?driver={cls.DB_DRIVER.replace(' ', '+')}&TrustServerCertificate=yes"

    @classmethod
    def test_connection(cls):
        try:
            if not SQLSERVER_AVAILABLE:
                return False
            conn_str = (
                f"DRIVER={{{cls.DB_DRIVER}}};"
                f"SERVER={cls.DB_SERVER};"
                f"DATABASE={cls.DB_NAME};"
                f"UID={cls.DB_USER};" # Já usa cls.DB_USER
                f"PWD={cls.DB_PASSWORD};" # Já usa cls.DB_PASSWORD
                "TrustServerCertificate=yes;"
            )
            conn = pyodbc.connect(conn_str, timeout=10)
            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            result = cursor.fetchone()
            conn.close()
            return result[0] == 1
        except Exception as e:
            print(f"❌ Erro no teste de conexão: {e}")
            return False

if SQLSERVER_AVAILABLE:
    Base = declarative_base()

    class Tarefa(Base):
        __tablename__ = 'painel_tarefas'
        id = Column(Integer, primary_key=True, autoincrement=True)
        titulo = Column(String(200), nullable=False)
        descricao = Column(Text)
        status = Column(String(50), default='Configurada')
        prioridade = Column(String(50), default='Média')
        arquivo_BAT = Column(String(500))
        data_criacao = Column(DateTime, default=func.now())
        data_ultima_execucao = Column(DateTime)
        total_execucoes = Column(Integer, default=0)
        
        # --- NOVOS CAMPOS PARA INTEGRAR COM CONNECT:DIRECT ---
        tipo_execucao = Column(String(50), default='BAT') # 'BAT', 'ConnectDirect'
        cd_source_path = Column(String(500)) # Caminho do arquivo de origem no nó local do CD
        cd_destination_node = Column(String(200)) # Nome do nó remoto do CD
        cd_destination_path = Column(String(500)) # Caminho do arquivo de destino no nó remoto do CD
        cd_process_name = Column(String(200)) # Nome de um processo CD pré-configurado (opcional)
        # --- FIM NOVOS CAMPOS ---

        execucoes = relationship("Execucao", back_populates="tarefa", cascade="all, delete-orphan")
        agendamentos = relationship("Agendamento", back_populates="tarefa", cascade="all, delete-orphan")

    class Execucao(Base):
        __tablename__ = 'painel_execucoes'
        id = Column(Integer, primary_key=True, autoincrement=True)
        tarefa_id = Column(Integer, ForeignKey('painel_tarefas.id'))
        data_execucao = Column(DateTime, default=func.now())
        status = Column(String(50))
        codigo_retorno = Column(Integer)
        duracao_segundos = Column(Float)
        log_output = Column(Text)
        executado_por_agendador = Column(Boolean, default=False)
        tarefa = relationship("Tarefa", back_populates="execucoes")

    class Agendamento(Base):
        __tablename__ = 'painel_agendamentos'
        id = Column(Integer, primary_key=True, autoincrement=True)
        tarefa_id = Column(Integer, ForeignKey('painel_tarefas.id'))
        nome = Column(String(200), nullable=False)
        tipo_agendamento = Column(String(50))  # diario, semanal, mensal, unico
        horario = Column(String(10))  # HH:MM
        dias_semana = Column(String(20))  # 0,1,2,3,4,5,6 (Dom-Sab)
        dia_mes = Column(Integer)  # Para agendamento mensal
        data_especifica = Column(DateTime)  # Para agendamento único
        ativo = Column(Boolean, default=True)
        retry_count = Column(Integer, default=0)
        max_retries = Column(Integer, default=3)
        notificar_sucesso = Column(Boolean, default=False)
        notificar_erro = Column(Boolean, default=True)
        data_criacao = Column(DateTime, default=func.now())
        proxima_execucao = Column(DateTime)
        tarefa = relationship("Tarefa", back_populates="agendamentos")

    class Alerta(Base):
        __tablename__ = 'painel_alertas'
        id = Column(Integer, primary_key=True, autoincrement=True)
        tipo = Column(String(50))  # erro, sucesso, timeout, retry
        titulo = Column(String(200))
        mensagem = Column(Text)
        tarefa_id = Column(Integer, ForeignKey('painel_tarefas.id'))
        agendamento_id = Column(Integer, ForeignKey('painel_agendamentos.id'))
        data_criacao = Column(DateTime, default=func.now())
        resolvido = Column(Boolean, default=False)
        data_resolucao = Column(DateTime)
        canal_notificacao = Column(String(50))  # email, whatsapp, telegram
        enviado = Column(Boolean, default=False)

    class ConfiguracaoNotificacao(Base):
        __tablename__ = 'painel_config_notificacao'
        id = Column(Integer, primary_key=True, autoincrement=True)
        canal = Column(String(50))  # email, whatsapp, telegram
        ativo = Column(Boolean, default=False)
        configuracao = Column(Text)  # JSON com configurações específicas
        data_atualizacao = Column(DateTime, default=func.now())

    class SistemaMonitor(Base):
        __tablename__ = 'painel_sistema_monitor'
        id = Column(Integer, primary_key=True, autoincrement=True)
        timestamp = Column(DateTime, default=func.now())
        cpu_percent = Column(Float)
        memoria_percent = Column(Float)
        disco_percent = Column(Float)
        rede_enviado = Column(Integer)
        rede_recebido = Column(Integer)

    class Configuracao(Base):
        __tablename__ = 'painel_configuracoes'
        chave = Column(String(100), primary_key=True)
        valor = Column(Text)

# --- FUNÇÃO AUXILIAR PARA CHAMAR A API DO IBM ICC ---
def call_icc_api(
    method: str, # GET, POST, PUT, DELETE
    endpoint_path: str, # Ex: "/transfers", "/processes"
    data: dict = None, # Dados para o corpo da requisição (POST, PUT, PATCH)
    params: dict = None, # Parâmetros de query (GET)
    auth_type: str = "basic" # Ou "bearer", dependendo da API
) -> tuple:
    """
    Função genérica para chamar a API do IBM Control Center.
    Os detalhes exatos de endpoints, corpo e autenticação virão da documentação da IBM.
    """
    if not REQUESTS_AVAILABLE:
        return False, {"error": "Biblioteca 'requests' não instalada para chamar API."}

    full_url = f"{ICC_API_BASE_URL}{endpoint_path}"
    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    # --- Autenticação ---
    req_auth = None
    if auth_type == "basic":
        req_auth = (ICC_USERNAME, ICC_PASSWORD)
    elif auth_type == "bearer":
        # Se a API usar Bearer Token, aqui você chamaria uma função para obter/renovar o token
        # Ex: token = get_bearer_token(ICC_USERNAME, ICC_PASSWORD)
        # if token:
        #     headers["Authorization"] = f"Bearer {token}"
        # else:
        #     print("Erro: Não foi possível obter o token de autenticação.")
        #     return False, {"error": "Authentication failed - token not obtained."}
        print("Autenticação Bearer Token não implementada neste exemplo.")
        return False, {"error": "Tipo de autenticação Bearer não implementado."}
    
    try:
        response = None
        # Timeout de 30 segundos para requisições da API
        # ATENÇÃO: verify=False desabilita a verificação de certificado SSL.
        # Em produção, você DEVE configurar a verificação de certificado corretamente.
        if method.upper() == "GET":
            response = requests.get(full_url, params=params, headers=headers, auth=req_auth, timeout=30, verify=False) 
        elif method.upper() == "POST":
            response = requests.post(full_url, json=data, headers=headers, auth=req_auth, timeout=30, verify=False) 
        elif method.upper() == "PUT":
            response = requests.put(full_url, json=data, headers=headers, auth=req_auth, timeout=30, verify=False) 
        elif method.upper() == "DELETE":
            response = requests.delete(full_url, headers=headers, auth=req_auth, timeout=30, verify=False) 
        else:
            raise ValueError(f"Método HTTP '{method}' não suportado por esta função.")

        response.raise_for_status() # Levanta um erro para status 4xx/5xx

        return True, response.json()

    except requests.exceptions.HTTPError as http_err:
        print(f"❌ Erro HTTP ao chamar API: {http_err}")
        print(f"   Status: {http_err.response.status_code}")
        print(f"   Resposta: {http_err.response.text}")
        return False, {"error": str(http_err), "response_text": http_err.response.text, "status_code": http_err.response.status_code}
    except requests.exceptions.ConnectionError as conn_err:
        print(f"❌ Erro de conexão à API: {conn_err}")
        return False, {"error": str(conn_err), "message": "Não foi possível conectar ao servidor da API. Verifique a URL e a conectividade."}
    except requests.exceptions.Timeout as timeout_err:
        print(f"❌ Tempo limite excedido ao chamar API: {timeout_err}")
        return False, {"error": str(timeout_err), "message": "A requisição da API excedeu o tempo limite."}
    except json.JSONDecodeError:
        print(f"❌ Erro ao decodificar JSON da resposta. Resposta bruta: {response.text if response else 'N/A'}")
        return False, {"error": "Resposta JSON inválida", "response_text": response.text if response else 'N/A'}
    except Exception as e:
        print(f"❌ Erro inesperado ao chamar API: {e}")
        import traceback
        traceback.print_exc()
        return False, {"error": str(e), "message": "Erro interno ao processar a requisição da API."}
# --- FIM DA FUNÇÃO AUXILIAR ---

class TaskDialog:
    def __init__(self, parent, title, task_data=None):
        self.result = None
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("600x600") # Aumentado para acomodar os novos campos
        self.top.transient(parent)
        self.top.grab_set()
        self.top.configure(bg='#1e3a3a')
        
        # Centralizar janela
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.top.winfo_screenheight() // 2) - (600 // 2)
        self.top.geometry(f"600x600+{x}+{y}")

        # Frame principal
        main_frame = ttk.Frame(self.top, padding="20")
        main_frame.pack(fill='both', expand=True)

        # Título
        ttk.Label(main_frame, text="Título da Tarefa:").grid(row=0, column=0, sticky='w', pady=5)
        self.titulo_entry = ttk.Entry(main_frame, width=50)
        self.titulo_entry.grid(row=0, column=1, padx=(10, 0), pady=5, sticky='ew')

        # Descrição
        ttk.Label(main_frame, text="Descrição:").grid(row=1, column=0, sticky='w', pady=5)
        self.descricao_entry = ttk.Entry(main_frame, width=50)
        self.descricao_entry.grid(row=1, column=1, padx=(10, 0), pady=5, sticky='ew')

        # Prioridade
        ttk.Label(main_frame, text="Prioridade:").grid(row=2, column=0, sticky='w', pady=5)
        self.prioridade_var = tk.StringVar(value="Média")
        prioridade_options = ["Baixa", "Média", "Alta"]
        self.prioridade_menu = ttk.Combobox(main_frame, textvariable=self.prioridade_var, 
                                           values=prioridade_options, state='readonly', width=47)
        self.prioridade_menu.grid(row=2, column=1, padx=(10, 0), pady=5, sticky='ew')
        
        # Status
        ttk.Label(main_frame, text="Status:").grid(row=3, column=0, sticky='w', pady=5)
        self.status_var = tk.StringVar(value="Configurada")
        self.status_label = ttk.Label(main_frame, text="Configurada", foreground='#4a9b9b')
        self.status_label.grid(row=3, column=1, padx=(10, 0), pady=5, sticky='w')

        # --- TIPO DE EXECUÇÃO (NOVO CAMPO) ---
        ttk.Label(main_frame, text="Tipo de Execução:").grid(row=4, column=0, sticky='w', pady=5)
        self.tipo_execucao_var = tk.StringVar(value="BAT")
        tipo_execucao_options = ["BAT", "ConnectDirect"]
        self.tipo_execucao_combo = ttk.Combobox(main_frame, textvariable=self.tipo_execucao_var,
                                               values=tipo_execucao_options, state='readonly', width=47)
        self.tipo_execucao_combo.grid(row=4, column=1, padx=(10, 0), pady=5, sticky='ew')
        self.tipo_execucao_combo.bind('<<ComboboxSelected>>', self.on_tipo_execucao_change)

        # --- FRAME PARA CAMPOS BAT ---
        self.BAT_config_frame = ttk.LabelFrame(main_frame, text="Configuração BAT", padding="10")
        self.BAT_config_frame.grid(row=5, column=0, columnspan=2, padx=0, pady=5, sticky='ew')
        ttk.Label(self.BAT_config_frame, text="Arquivo BAT (.bat):").grid(row=0, column=0, sticky='w', pady=5)
        
        BAT_entry_frame = ttk.Frame(self.BAT_config_frame)
        BAT_entry_frame.grid(row=0, column=1, padx=(10, 0), pady=5, sticky='ew')
        
        self.BAT_entry = ttk.Entry(BAT_entry_frame, width=40)
        self.BAT_entry.pack(side='left', fill='x', expand=True)
        ttk.Button(BAT_entry_frame, text="📁 Procurar", 
                  command=self.select_BAT_file).pack(side='right', padx=(5, 0))
        self.BAT_config_frame.columnconfigure(1, weight=1)
        BAT_entry_frame.columnconfigure(0, weight=1)

        # --- FRAME PARA CAMPOS CONNECT:DIRECT ---
        self.cd_config_frame = ttk.LabelFrame(main_frame, text="Configuração Connect:Direct", padding="10")
        self.cd_config_frame.grid(row=6, column=0, columnspan=2, padx=0, pady=5, sticky='ew')

        ttk.Label(self.cd_config_frame, text="Caminho Origem (CD):").grid(row=0, column=0, sticky='w', pady=2)
        self.cd_source_path_entry = ttk.Entry(self.cd_config_frame, width=50)
        self.cd_source_path_entry.grid(row=0, column=1, padx=(10, 0), pady=2, sticky='ew')

        ttk.Label(self.cd_config_frame, text="Nó Destino (CD):").grid(row=1, column=0, sticky='w', pady=2)
        self.cd_destination_node_entry = ttk.Entry(self.cd_config_frame, width=50)
        self.cd_destination_node_entry.grid(row=1, column=1, padx=(10, 0), pady=2, sticky='ew')

        ttk.Label(self.cd_config_frame, text="Caminho Destino (CD):").grid(row=2, column=0, sticky='w', pady=2)
        self.cd_destination_path_entry = ttk.Entry(self.cd_config_frame, width=50)
        self.cd_destination_path_entry.grid(row=2, column=1, padx=(10, 0), pady=2, sticky='ew')

        ttk.Label(self.cd_config_frame, text="Processo CD (Opcional):").grid(row=3, column=0, sticky='w', pady=2)
        self.cd_process_name_entry = ttk.Entry(self.cd_config_frame, width=50)
        self.cd_process_name_entry.grid(row=3, column=1, padx=(10, 0), pady=2, sticky='ew')

        self.cd_config_frame.columnconfigure(1, weight=1)
        # --- FIM NOVOS CAMPOS ---

        # Preencher campos se for edição
        if task_data:
            self.titulo_entry.insert(0, task_data.titulo or "")
            self.descricao_entry.insert(0, task_data.descricao or "")
            self.prioridade_var.set(task_data.prioridade or "Média")
            self.tipo_execucao_var.set(getattr(task_data, 'tipo_execucao', "BAT")) # Novo campo
            
            # Preencher campos específicos de acordo com o tipo de execução
            if self.tipo_execucao_var.get() == "BAT":
                self.BAT_entry.insert(0, task_data.arquivo_BAT or "")
            elif self.tipo_execucao_var.get() == "ConnectDirect":
                self.cd_source_path_entry.insert(0, getattr(task_data, 'cd_source_path', "") or "")
                self.cd_destination_node_entry.insert(0, getattr(task_data, 'cd_destination_node', "") or "")
                self.cd_destination_path_entry.insert(0, getattr(task_data, 'cd_destination_path', "") or "")
                self.cd_process_name_entry.insert(0, getattr(task_data, 'cd_process_name', "") or "")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=7, column=0, columnspan=2, pady=20) # Linha ajustada
        
        ttk.Button(btn_frame, text="💾 Salvar", width=15, command=self.on_ok).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="❌ Cancelar", width=15, command=self.top.destroy).pack(side='left', padx=5)

        main_frame.columnconfigure(1, weight=1)
        
        self.top.bind("<Return>", lambda event: self.on_ok())
        self.top.bind("<Escape>", lambda event: self.top.destroy())
        
        self.on_tipo_execucao_change() # Chamar para configurar a visibilidade inicial
        self.titulo_entry.focus()
        
        self.top.wait_window(self.top)

    def select_BAT_file(self):
        """Seleciona arquivo BAT"""
        filename = filedialog.askopenfilename(
            title="Selecionar arquivo BAT",
            filetypes=[("BAT files", "*.bat"), ("All files", "*.*")]
        )
        if filename:
            self.BAT_entry.delete(0, tk.END)
            self.BAT_entry.insert(0, filename)

    def on_tipo_execucao_change(self, event=None):
        """Atualiza a visibilidade dos campos de configuração de acordo com o tipo de execução."""
        selected_type = self.tipo_execucao_var.get()
        if selected_type == "BAT":
            self.BAT_config_frame.grid(row=5, column=0, columnspan=2, padx=0, pady=5, sticky='ew')
            self.cd_config_frame.grid_forget()
        elif selected_type == "ConnectDirect":
            self.BAT_config_frame.grid_forget()
            self.cd_config_frame.grid(row=5, column=0, columnspan=2, padx=0, pady=5, sticky='ew') # Colocado na mesma linha que antes
            # Também ajusta a row do btn_frame, mas o grid já fará isso automaticamente.

    def on_ok(self):
        titulo = self.titulo_entry.get().strip()
        descricao = self.descricao_entry.get().strip()
        prioridade = self.prioridade_var.get()
        tipo_execucao = self.tipo_execucao_var.get()
        arquivo_BAT = None
        cd_source_path = None
        cd_destination_node = None
        cd_destination_path = None
        cd_process_name = None
        
        if not titulo:
            messagebox.showwarning("Aviso", "O título é obrigatório.")
            return
            
        if tipo_execucao == "BAT":
            arquivo_BAT = self.BAT_entry.get().strip()
            if not arquivo_BAT:
                messagebox.showwarning("Aviso", "Selecione um arquivo BAT.")
                return
            if not os.path.exists(arquivo_BAT):
                messagebox.showwarning("Aviso", "O arquivo BAT selecionado não existe.")
                return
        elif tipo_execucao == "ConnectDirect":
            cd_source_path = self.cd_source_path_entry.get().strip()
            cd_destination_node = self.cd_destination_node_entry.get().strip()
            cd_destination_path = self.cd_destination_path_entry.get().strip()
            cd_process_name = self.cd_process_name_entry.get().strip() # Pode ser opcional, mas vamos validar se vazio
            
            if not all([cd_source_path, cd_destination_node, cd_destination_path]):
                messagebox.showwarning("Aviso", "Para Connect:Direct, os campos 'Caminho Origem', 'Nó Destino' e 'Caminho Destino' são obrigatórios.")
                return
            
        self.result = {
            'titulo': titulo,
            'descricao': descricao,
            'prioridade': prioridade,
            'tipo_execucao': tipo_execucao, # Novo campo
            'arquivo_BAT': arquivo_BAT,
            'cd_source_path': cd_source_path, # Novos campos
            'cd_destination_node': cd_destination_node,
            'cd_destination_path': cd_destination_path,
            'cd_process_name': cd_process_name if cd_process_name else None # Salva None se vazio
        }
        self.top.destroy()


class ScheduleDialog:
    def __init__(self, parent, title, task_id, schedule_data=None):
        self.result = None
        self.task_id = task_id
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("600x500")
        self.top.transient(parent)
        self.top.grab_set()
        self.top.configure(bg='#1e3a3a')
        
        # Centralizar janela
        self.top.update_idletasks()
        x = (self.top.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.top.winfo_screenheight() // 2) - (500 // 2)
        self.top.geometry(f"600x500+{x}+{y}")

        # Frame principal
        main_frame = ttk.Frame(self.top, padding="20")
        main_frame.pack(fill='both', expand=True)

        # Nome do agendamento
        ttk.Label(main_frame, text="Nome do Agendamento:").grid(row=0, column=0, sticky='w', pady=5)
        self.nome_entry = ttk.Entry(main_frame, width=50)
        self.nome_entry.grid(row=0, column=1, padx=(10, 0), pady=5, sticky='ew')

        # Tipo de agendamento
        ttk.Label(main_frame, text="Tipo:").grid(row=1, column=0, sticky='w', pady=5)
        self.tipo_var = tk.StringVar(value="diario")
        tipo_options = ["diario", "semanal", "mensal", "unico"]
        self.tipo_combo = ttk.Combobox(main_frame, textvariable=self.tipo_var, 
                                      values=tipo_options, state='readonly', width=47)
        self.tipo_combo.grid(row=1, column=1, padx=(10, 0), pady=5, sticky='ew')
        self.tipo_combo.bind('<<ComboboxSelected>>', self.on_tipo_change)

        # Horário
        ttk.Label(main_frame, text="Horário (HH:MM):").grid(row=2, column=0, sticky='w', pady=5)
        self.horario_entry = ttk.Entry(main_frame, width=50)
        self.horario_entry.grid(row=2, column=1, padx=(10, 0), pady=5, sticky='ew')
        self.horario_entry.insert(0, "08:00")

        # Frame para configurações específicas
        self.config_frame = ttk.LabelFrame(main_frame, text="Configurações Específicas", padding=10)
        self.config_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky='ew')

        # Dias da semana (para agendamento semanal)
        self.dias_frame = ttk.Frame(self.config_frame)
        ttk.Label(self.dias_frame, text="Dias da Semana:").pack(anchor='w')
        
        self.dias_vars = {}
        dias_nomes = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado']
        dias_frame_checks = ttk.Frame(self.dias_frame)
        dias_frame_checks.pack(fill='x', pady=5)
        
        for i, dia in enumerate(dias_nomes):
            var = tk.BooleanVar()
            self.dias_vars[i] = var
            ttk.Checkbutton(dias_frame_checks, text=dia, variable=var).pack(side='left', padx=5)

        # Dia do mês (para agendamento mensal)
        self.dia_mes_frame = ttk.Frame(self.config_frame)
        ttk.Label(self.dia_mes_frame, text="Dia do Mês:").pack(anchor='w')
        self.dia_mes_var = tk.IntVar(value=1)
        ttk.Spinbox(self.dia_mes_frame, from_=1, to=31, textvariable=self.dia_mes_var, width=10).pack(anchor='w', pady=5)

        # Data específica (para agendamento único)
        self.data_frame = ttk.Frame(self.config_frame)
        ttk.Label(self.data_frame, text="Data (DD/MM/AAAA):").pack(anchor='w')
        self.data_entry = ttk.Entry(self.data_frame, width=20)
        self.data_entry.pack(anchor='w', pady=5)
        self.data_entry.insert(0, datetime.datetime.now().strftime("%d/%m/%Y"))

        # Configurações de retry
        retry_frame = ttk.LabelFrame(main_frame, text="Configurações de Retry", padding=10)
        retry_frame.grid(row=4, column=0, columnspan=2, pady=10, sticky='ew')

        ttk.Label(retry_frame, text="Máximo de Tentativas:").grid(row=0, column=0, sticky='w', pady=5)
        self.max_retries_var = tk.IntVar(value=3)
        ttk.Spinbox(retry_frame, from_=0, to=10, textvariable=self.max_retries_var, width=10).grid(row=0, column=1, padx=(10, 0), pady=5)

        # Configurações de notificação
        notif_frame = ttk.LabelFrame(main_frame, text="Notificações", padding=10)
        notif_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky='ew')

        self.notif_sucesso_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(notif_frame, text="Notificar em caso de sucesso", variable=self.notif_sucesso_var).pack(anchor='w')

        self.notif_erro_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(notif_frame, text="Notificar em caso de erro", variable=self.notif_erro_var).pack(anchor='w')

        # Botões
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="💾 Salvar", width=15, command=self.on_ok).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="❌ Cancelar", width=15, command=self.top.destroy).pack(side='left', padx=5)

        # Configurar grid
        main_frame.columnconfigure(1, weight=1)
        self.config_frame.columnconfigure(0, weight=1)

        # Configurar visibilidade inicial
        self.on_tipo_change()

        # Preencher campos se for edição
        if schedule_data:
            self.nome_entry.insert(0, schedule_data.nome or "")
            self.tipo_var.set(schedule_data.tipo_agendamento or "diario")
            self.horario_entry.delete(0, tk.END)
            self.horario_entry.insert(0, schedule_data.horario or "08:00")
            self.max_retries_var.set(schedule_data.max_retries or 3)
            self.notif_sucesso_var.set(schedule_data.notificar_sucesso or False)
            self.notif_erro_var.set(schedule_data.notificar_erro or True)
            
            if schedule_data.dias_semana:
                dias_list = [int(d) for d in schedule_data.dias_semana.split(',') if d.strip()]
                for dia in dias_list:
                    if dia in self.dias_vars:
                        self.dias_vars[dia].set(True)
            
            if schedule_data.dia_mes:
                self.dia_mes_var.set(schedule_data.dia_mes)
            
            if schedule_data.data_especifica:
                self.data_entry.delete(0, tk.END)
                self.data_entry.insert(0, schedule_data.data_especifica.strftime("%d/%m/%Y"))

        # Eventos
        self.top.bind("<Return>", lambda event: self.on_ok())
        self.top.bind("<Escape>", lambda event: self.top.destroy())
        
        # Focar no primeiro campo
        self.nome_entry.focus()
        
        self.top.wait_window(self.top)

    def on_tipo_change(self, event=None):
        """Atualiza visibilidade dos campos baseado no tipo"""
        tipo = self.tipo_var.get()
        
        # Esconder todos os frames
        self.dias_frame.pack_forget()
        self.dia_mes_frame.pack_forget()
        self.data_frame.pack_forget()
        
        # Mostrar frame apropriado
        if tipo == "semanal":
            self.dias_frame.pack(fill='x', pady=5)
        elif tipo == "mensal":
            self.dia_mes_frame.pack(fill='x', pady=5)
        elif tipo == "unico":
            self.data_frame.pack(fill='x', pady=5)

    def on_ok(self):
        nome = self.nome_entry.get().strip()
        tipo = self.tipo_var.get()
        horario = self.horario_entry.get().strip()
        
        if not nome:
            messagebox.showwarning("Aviso", "O nome do agendamento é obrigatório.")
            return
        
        # Validar horário
        try:
            datetime.datetime.strptime(horario, "%H:%M")
        except ValueError:
            messagebox.showwarning("Aviso", "Formato de horário inválido. Use HH:MM.")
            return
        
        # Preparar dados específicos
        dias_semana = None
        dia_mes = None
        data_especifica = None
        
        if tipo == "semanal":
            dias_selecionados = [str(dia) for dia, var in self.dias_vars.items() if var.get()]
            if not dias_selecionados:
                messagebox.showwarning("Aviso", "Selecione pelo menos um dia da semana.")
                return
            dias_semana = ",".join(dias_selecionados)
        
        elif tipo == "mensal":
            dia_mes = self.dia_mes_var.get()
        
        elif tipo == "unico":
            try:
                data_especifica = datetime.datetime.strptime(self.data_entry.get(), "%d/%m/%Y")
                if data_especifica < datetime.datetime.now():
                    messagebox.showwarning("Aviso", "A data deve ser futura.")
                    return
            except ValueError:
                messagebox.showwarning("Aviso", "Formato de data inválido. Use DD/MM/AAAA.")
                return
        
        self.result = {
            'nome': nome,
            'tipo_agendamento': tipo,
            'horario': horario,
            'dias_semana': dias_semana,
            'dia_mes': dia_mes,
            'data_especifica': data_especifica,
            'max_retries': self.max_retries_var.get(),
            'notificar_sucesso': self.notif_sucesso_var.get(),
            'notificar_erro': self.notif_erro_var.get()
        }
        self.top.destroy()

# ==================== CLASSE PRINCIPAL DO PAINEL ====================

class PainelDesktop:
    def __init__(self, root):
        self.root = root
        self.root.title("🤖 Painel de Controle de Rotinas Automatizadas")
        self.root.geometry("1200x800")
        self.root.configure(bg='#1e3a3a')
        
        # Tentar maximizar janela
        try:
            self.root.state('zoomed')  # Windows
        except:
            try:
                self.root.attributes('-zoomed', True)  # Linux
            except:
                pass  # Mac ou outros
        
        # Inicializar variáveis
        self.engine = None
        self.SessionLocal = None
        self.using_sqlite = False
        
        # Configuração do banco de dados
        self.init_database()
        
        # Variáveis de controle
        self.scheduler_running = False
        self.current_theme = "teal"
        self.notification_configs = {}
        
        # Configuração da interface
        self.setup_styles()
        self.create_widgets()
        self.load_data()
        
        # Iniciar agendador
        self.start_scheduler()

    def init_database(self):
        """Inicializa o banco de dados"""
        if SQLSERVER_AVAILABLE:
            self.init_sqlserver_database()
        else:
            self.init_sqlite_database()

    def init_sqlserver_database(self):
        """Inicializa o banco de dados SQL Server"""
        try:
            print("🔄 Conectando ao SQL Server...")
            print(f"📍 Servidor: {DatabaseConfig.DB_SERVER}")
            print(f"🗄️ Banco: {DatabaseConfig.DB_NAME}")
            print(f"👤 Usuário: {DatabaseConfig.DB_USER}")
            
            # Testar conexão primeiro
            if not DatabaseConfig.test_connection():
                raise Exception("Falha no teste de conexão")
            
            # Criar engine
            connection_string = DatabaseConfig.get_connection_string()
            self.engine = create_engine(
                connection_string,
                echo=False,
                pool_pre_ping=True,
                pool_recycle=3600,
                pool_timeout=30,
                max_overflow=10,
                connect_args={
                    "TrustServerCertificate": "yes",
                    "timeout": 30
                }
            )
            
            # Criar sessão
            self.SessionLocal = sessionmaker(
                autocommit=False,
                autoflush=False,
                bind=self.engine
            )
            
            # Testar engine
            with self.engine.connect() as connection:
                result = connection.execute(sqlalchemy.text("SELECT 1 as test"))
                test_result = result.fetchone()
                if test_result[0] != 1:
                    raise Exception("Teste de engine falhou")
            
            # Criar tabelas se não existirem
            print("📋 Criando/verificando tabelas...")
            Base.metadata.create_all(bind=self.engine)
            
            print("✅ Conexão com SQL Server configurada com sucesso!")
            print("🎉 Sistema pronto para uso!")
            
            # Verificar se existem dados
            self.check_initial_data()
            
        except Exception as e:
            error_msg = f"❌ Erro ao conectar com SQL Server: {e}"
            print(error_msg)
            
            # Perguntar se quer usar SQLite como fallback
            use_sqlite = messagebox.askyesno(
                "Erro de Conexão SQL Server", 
                f"Não foi possível conectar ao SQL Server:\n\n{e}\n\n"
                "Deseja usar SQLite como alternativa temporária?"
            )
            
            if use_sqlite:
                self.init_sqlite_database()
            else:
                messagebox.showerror("Erro", "Aplicação será encerrada.")
                self.root.quit()

    def init_sqlite_database(self):
        """Inicializa o banco de dados SQLite como fallback"""
        try:
            print("🔄 Configurando SQLite como fallback...")
            
            self.db_path = "painel_dados.db"
            self.using_sqlite = True
            
            # Criar engine para SQLite
            self.engine = create_engine(f'sqlite:///{self.db_path}', echo=False)
            self.SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=self.engine)
            
            # Para SQLite, precisamos recriar os modelos
            if SQLSERVER_AVAILABLE: # Se SQLALCHEMY estiver disponível, use o Base.metadata
                Base.metadata.create_all(bind=self.engine)
            else: # Se SQLALCHEMY não estiver disponível, crie tabelas manualmente
                self.create_sqlite_tables()
            
            print("✅ SQLite configurado com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro Fatal", f"Erro ao configurar banco de dados: {e}")
            self.root.quit()

    def create_sqlite_tables(self):
        """Cria tabelas SQLite manualmente"""
        import sqlite3
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Tabela de tarefas - ATUALIZADA COM NOVOS CAMPOS
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS painel_tarefas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                descricao TEXT,
                status TEXT DEFAULT 'Configurada',
                prioridade TEXT DEFAULT 'Média',
                arquivo_BAT TEXT,
                data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                data_ultima_execucao TIMESTAMP,
                total_execucoes INTEGER DEFAULT 0,
                -- NOVOS CAMPOS PARA CONNECT:DIRECT
                tipo_execucao TEXT DEFAULT 'BAT',
                cd_source_path TEXT,
                cd_destination_node TEXT,
                cd_destination_path TEXT,
                cd_process_name TEXT
            )
        ''')
        
        # Tabela de execuções
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS painel_execucoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tarefa_id INTEGER,
                data_execucao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                status TEXT,
                codigo_retorno INTEGER,
                duracao_segundos REAL,
                log_output TEXT,
                executado_por_agendador BOOLEAN DEFAULT 0,
                FOREIGN KEY (tarefa_id) REFERENCES painel_tarefas (id)
            )
        ''')
        
        # Tabela de agendamentos
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS painel_agendamentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tarefa_id INTEGER,
                nome TEXT NOT NULL,
                tipo_agendamento TEXT,
                horario TEXT,
                dias_semana TEXT,
                dia_mes INTEGER,
                data_especifica TIMESTAMP,
                ativo BOOLEAN DEFAULT 1,
                retry_count INTEGER DEFAULT 0,
                max_retries INTEGER DEFAULT 3,
                notificar_sucesso BOOLEAN DEFAULT 0,
                notificar_erro BOOLEAN DEFAULT 1,
                data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                proxima_execucao TIMESTAMP,
                FOREIGN KEY (tarefa_id) REFERENCES painel_tarefas (id)
            )
        ''')
        
        # Tabela de alertas
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS painel_alertas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo TEXT,
                titulo TEXT,
                mensagem TEXT,
                tarefa_id INTEGER,
                agendamento_id INTEGER,
                data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                resolvido BOOLEAN DEFAULT 0,
                data_resolucao TIMESTAMP,
                canal_notificacao TEXT,
                enviado BOOLEAN DEFAULT 0,
                FOREIGN KEY (tarefa_id) REFERENCES painel_tarefas (id),
                FOREIGN KEY (agendamento_id) REFERENCES painel_agendamentos (id)
            )
        ''')
        
        # Tabela de configurações de notificação
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS painel_config_notificacao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                canal TEXT,
                ativo BOOLEAN DEFAULT 0,
                configuracao TEXT,
                data_atualizacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Tabela de monitoramento
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS painel_sistema_monitor (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                cpu_percent REAL,
                memoria_percent REAL,
                disco_percent REAL,
                rede_enviado INTEGER,
                rede_recebido INTEGER
            )
        ''')
        
        conn.commit()
        conn.close()

    def check_initial_data(self):
        """Verifica se há dados iniciais no banco"""
        try:
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                # Para SQLite sem SQLAlchemy
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                cursor.execute("SELECT COUNT(*) FROM painel_tarefas")
                count_tarefas = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM painel_execucoes")
                count_execucoes = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM painel_agendamentos")
                count_agendamentos = cursor.fetchone()[0]
                
                conn.close()
            else:
                # Para SQL Server com SQLAlchemy
                session = self.get_db_session()
                try:
                    count_tarefas = session.query(Tarefa).count()
                    count_execucoes = session.query(Execucao).count()
                    count_agendamentos = session.query(Agendamento).count()
                finally:
                    session.close()
            
            print(f"📊 Dados encontrados:")
            print(f"   - Tarefas: {count_tarefas}")
            print(f"   - Execuções: {count_execucoes}")
            print(f"   - Agendamentos: {count_agendamentos}")
            
        except Exception as e:
            print(f"⚠️ Erro ao verificar dados iniciais: {e}")

    def get_db_session(self):
        """Retorna nova sessão do banco"""
        if self.SessionLocal:
            return self.SessionLocal()
        else:
            raise Exception("Sessão do banco não configurada")

    def setup_styles(self):
        """Configura os estilos da interface - Tema Verde Petróleo"""
        style = ttk.Style()
        
        # Cores do tema verde petróleo
        colors = {
            'bg_dark': '#1e3a3a',
            'bg_medium': '#2d5555',
            'bg_light': '#4a9b9b',
            'fg_light': '#e8f4f4',
            'fg_dark': '#1e3a3a',
            'accent': '#66cccc',
            'success': '#4ade80',
            'error': '#ef4444',
            'warning': '#f59e0b'
        }
        
        try:
            style.theme_use('clam')
        except:
            style.theme_use('default')
        
        # Configurar Notebook
        style.configure('TNotebook', 
                       background=colors['bg_medium'], 
                       borderwidth=0,
                       tabmargins=[2, 5, 2, 0])
        
        style.configure('TNotebook.Tab', 
                       background=colors['bg_dark'], 
                       foreground=colors['fg_light'],
                       padding=[15, 8], 
                       borderwidth=1,
                       focuscolor='none')
        
        style.map('TNotebook.Tab', 
                 background=[('selected', colors['bg_light']),
                           ('active', colors['bg_medium'])],
                 foreground=[('selected', colors['fg_dark']),
                           ('active', colors['fg_light'])])
        
        # Configurar outros elementos
        style.configure('TFrame', background=colors['bg_medium'])
        style.configure('TLabel', background=colors['bg_medium'], foreground=colors['fg_light'], font=('Segoe UI', 9))
        style.configure('TButton', background=colors['bg_light'], foreground=colors['fg_dark'], borderwidth=1, focuscolor='none', font=('Segoe UI', 9))
        style.configure('TLabelFrame', background=colors['bg_medium'], foreground=colors['fg_light'], borderwidth=2, relief='groove')
        style.configure('TLabelFrame.Label', background=colors['bg_medium'], foreground=colors['accent'], font=('Segoe UI', 10, 'bold'))
        style.configure('TEntry', fieldbackground=colors['bg_light'], foreground=colors['fg_dark'], borderwidth=1, insertcolor=colors['fg_dark'])
        style.configure('TCombobox', fieldbackground=colors['bg_light'], foreground=colors['fg_dark'], arrowcolor=colors['fg_dark'])
        style.configure('Treeview', background=colors['bg_light'], foreground=colors['fg_dark'], fieldbackground=colors['bg_light'], borderwidth=1)
        style.configure('Treeview.Heading', background=colors['bg_dark'], foreground=colors['fg_light'], font=('Segoe UI', 9, 'bold'))
        style.configure('TScrollbar', background=colors['bg_medium'], troughcolor=colors['bg_dark'], arrowcolor=colors['fg_light'])

    def get_greeting(self):
        """Retorna saudação baseada no horário"""
        now = datetime.datetime.now()
        hour = now.hour
        
        if 5 <= hour < 12:
            return "🌅 Bom dia, usuário!" 
        elif 12 <= hour < 18:
            return "☀️ Boa tarde, usuário!" 
        else:
            return "🌙 Boa noite, usuário!" 

    def create_widgets(self):
        """Cria os widgets da interface"""
        # Menu principal
        self.create_menu()
        
        # Notebook para abas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Abas
        self.create_dashboard_tab()
        self.create_tasks_tab()
        self.create_control_center_tab()
        self.create_settings_tab()

    def create_menu(self):
        """Cria o menu principal"""
        menubar = tk.Menu(self.root, bg='#1e3a3a', fg='#e8f4f4', activebackground='#4a9b9b', activeforeground='#1e3a3a')
        self.root.config(menu=menubar)
        
        # Menu Arquivo
        file_menu = tk.Menu(menubar, tearoff=0, bg='#2d5555', fg='#e8f4f4', activebackground='#4a9b9b', activeforeground='#1e3a3a')
        menubar.add_cascade(label="Arquivo", menu=file_menu)
        file_menu.add_command(label="Exportar Dados", command=self.export_data)
        file_menu.add_command(label="Importar Dados", command=self.import_data)
        file_menu.add_separator()
        file_menu.add_command(label="Sair", command=self.on_closing)
        
        # Menu Ferramentas
        tools_menu = tk.Menu(menubar, tearoff=0, bg='#2d5555', fg='#e8f4f4', activebackground='#4a9b9b', activeforeground='#1e3a3a')
        menubar.add_cascade(label="Ferramentas", menu=tools_menu)
        tools_menu.add_command(label="Backup Banco", command=self.backup_database)
        tools_menu.add_command(label="Limpar Dados", command=self.clear_data)
        tools_menu.add_command(label="Teste de Notificação", command=self.test_notification)
        
        # Menu Ajuda
        help_menu = tk.Menu(menubar, tearoff=0, bg='#2d5555', fg='#e8f4f4', activebackground='#4a9b9b', activeforeground='#1e3a3a')
        menubar.add_cascade(label="Ajuda", menu=help_menu)
        help_menu.add_command(label="Sobre", command=self.show_about)

    def create_dashboard_tab(self):
        """Cria a aba do dashboard"""
        dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(dashboard_frame, text="📊 Dashboard")
        
        # Saudação
        greeting_frame = ttk.Frame(dashboard_frame)
        greeting_frame.pack(fill='x', padx=10, pady=10)
        
        greeting_label = ttk.Label(greeting_frame, text=self.get_greeting(), 
                                  font=('Segoe UI', 16, 'bold'))
        greeting_label.pack(side='left')
        
        # Indicador de banco
        db_type = "SQL Server" if not self.using_sqlite else "SQLite"
        db_indicator = ttk.Label(greeting_frame, text=f"💾 {db_type}", 
                               font=('Segoe UI', 10), foreground='#66cccc')
        db_indicator.pack(side='right')
        
        welcome_label = ttk.Label(greeting_frame, 
                                 text="Bem-vindo ao painel de controle das suas rotinas automatizadas", 
                                 font=('Segoe UI', 12))
        welcome_label.pack(side='left', padx=(10, 0))
        
        # Frame superior com estatísticas
        stats_frame = ttk.LabelFrame(dashboard_frame, text="📈 Estatísticas Gerais", padding=10)
        stats_frame.pack(fill='x', padx=10, pady=5)
        
        # Labels de estatísticas
        self.stats_labels = {}
        stats_data = [
            ("Tarefas Configuradas", "tasks_total"),
            ("Agendamentos Ativos", "schedules_active"),
            ("Execuções Hoje", "executions_today"),
            ("Alertas Pendentes", "alerts_pending"),
            ("Última Execução", "last_execution"),
            ("Taxa de Sucesso", "success_rate")
        ]
        
        # Organizar em duas linhas
        for i, (label, key) in enumerate(stats_data):
            row = i // 3
            col = i % 3
            
            frame = ttk.Frame(stats_frame)
            frame.grid(row=row, column=col, padx=10, pady=5, sticky='ew')
            
            ttk.Label(frame, text=label, font=('Segoe UI', 10, 'bold')).pack()
            self.stats_labels[key] = ttk.Label(frame, text="0", font=('Segoe UI', 14, 'bold'))
            self.stats_labels[key].pack()
        
        # Configurar grid
        for i in range(3):
            stats_frame.columnconfigure(i, weight=1)
        
                # Frame para gráficos
        charts_frame = ttk.LabelFrame(dashboard_frame, text="📈 Histórico de Execuções", padding=10)
        charts_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Criar gráfico de execuções
        self.create_execution_chart(charts_frame)

    def create_execution_chart(self, parent):
        """Cria gráfico de execuções por data"""
        try:
            # Configurar matplotlib para tema escuro
            plt.style.use('dark_background')
            
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4))
            fig.patch.set_facecolor('#1e3a3a')
            
            # Buscar dados
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                # SQLite sem SQLAlchemy
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                
                # Execuções por data
                df_executions = pd.read_sql_query("""
                    SELECT DATE(data_execucao) as data, COUNT(*) as count,
                           SUM(CASE WHEN status = 'sucesso' THEN 1 ELSE 0 END) as sucessos
                    FROM painel_execucoes 
                    WHERE data_execucao >= datetime('now', '-7 days')
                    GROUP BY DATE(data_execucao)
                    ORDER BY data
                """, conn)
                
                # Tarefas por prioridade
                df_tasks = pd.read_sql_query("""
                    SELECT prioridade, COUNT(*) as count 
                    FROM painel_tarefas 
                    GROUP BY prioridade
                """, conn)
                
                conn.close()
            else:
                # SQL Server com SQLAlchemy
                from sqlalchemy import text
                
                with self.engine.connect() as conn:
                    # Execuções por data
                    df_executions = pd.read_sql_query("""
                        SELECT CAST(data_execucao AS DATE) as data, COUNT(*) as count,
                               SUM(CASE WHEN status = 'sucesso' THEN 1 ELSE 0 END) as sucessos
                        FROM painel_execucoes 
                        WHERE data_execucao >= DATEADD(day, -7, GETDATE())
                        GROUP BY CAST(data_execucao AS DATE)
                        ORDER BY data
                    """, conn)
                    
                    # Tarefas por prioridade
                    df_tasks = pd.read_sql_query("""
                        SELECT prioridade, COUNT(*) as count 
                        FROM painel_tarefas 
                        GROUP BY prioridade
                    """, conn)
            
            # Gráfico 1: Execuções por data
            if not df_executions.empty:
                ax1.bar(df_executions['data'], df_executions['count'],
				                       color='#4a9b9b', alpha=0.7, label='Total')
                ax1.bar(df_executions['data'], df_executions['sucessos'], 
                       color='#66cccc', alpha=0.8, label='Sucessos')
                ax1.set_title('Execuções por Data (7 dias)', color='#e8f4f4', fontsize=12)
                ax1.set_ylabel('Quantidade', color='#e8f4f4')
                ax1.legend()
                ax1.tick_params(axis='x', rotation=45, colors='#e8f4f4')
                ax1.tick_params(axis='y', colors='#e8f4f4')
            else:
                ax1.text(0.5, 0.5, 'Sem execuções nos últimos 7 dias', 
                        ha='center', va='center', color='#e8f4f4', transform=ax1.transAxes)
                ax1.set_title('Execuções por Data (7 dias)', color='#e8f4f4', fontsize=12)
            
            # Gráfico 2: Tarefas por prioridade
            if not df_tasks.empty:
                colors = {'Alta': '#ef4444', 'Média': '#f59e0b', 'Baixa': '#4ade80'}
                task_colors = [colors.get(p, '#4a9b9b') for p in df_tasks['prioridade']]
                
                ax2.pie(df_tasks['count'], labels=df_tasks['prioridade'], 
                       autopct='%1.1f%%', colors=task_colors, textprops={'color': '#e8f4f4'})
                ax2.set_title('Tarefas por Prioridade', color='#e8f4f4', fontsize=12)
            else:
                ax2.text(0.5, 0.5, 'Sem tarefas cadastradas', 
                        ha='center', va='center', color='#e8f4f4', transform=ax2.transAxes)
                ax2.set_title('Tarefas por Prioridade', color='#e8f4f4', fontsize=12)
            
            # Configurar cores dos eixos
            for ax in [ax1, ax2]:
                ax.set_facecolor('#2d5555')
                ax.grid(True, alpha=0.3, color='#4a9b9b')
            
            plt.tight_layout()
            
            # Adicionar ao tkinter
            canvas = FigureCanvasTkAgg(fig, parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            
        except Exception as e:
            error_label = ttk.Label(parent, text=f"Erro ao carregar gráficos: {e}")
            error_label.pack(pady=20)

    def create_tasks_tab(self):
        """Cria a aba de tarefas"""
        tasks_frame = ttk.Frame(self.notebook)
        self.notebook.add(tasks_frame, text="🤖 Tarefas Automatizadas")
        
        # Frame superior com controles
        controls_frame = ttk.Frame(tasks_frame)
        controls_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(controls_frame, text="➕ Nova Tarefa", 
                  command=self.new_task).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="▶️ Executar Tarefa", 
                  command=self.execute_task).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="✏️ Editar Tarefa", 
                  command=self.edit_task).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="🗑️ Excluir Tarefa", 
                  command=self.delete_task).pack(side='left', padx=5)
        
        # Separador
        ttk.Separator(controls_frame, orient='vertical').pack(side='left', fill='y', padx=10)
        
        ttk.Button(controls_frame, text="📊 Ver Histórico", 
                  command=self.show_execution_history).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="🔄 Atualizar", 
                  command=self.refresh_tasks).pack(side='left', padx=5)
        
        # Frame para tabela
        table_frame = ttk.Frame(tasks_frame)
        table_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Treeview para tarefas
        # COLUNAS ATUALIZADAS PARA INCLUIR TIPO DE EXECUÇÃO
        columns = ('ID', 'Título', 'Tipo Execução', 'Prioridade', 'Detalhes Config.', 'Última Execução', 'Total Execuções')
        self.tasks_tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        # Configurar colunas
        column_widths = {
            'ID': 50, 
            'Título': 180, 
            'Tipo Execução': 100, # Nova coluna
            'Prioridade': 90, 
            'Detalhes Config.': 250, # Nova coluna para substituir 'Arquivo BAT' e incluir CD
            'Última Execução': 150, 
            'Total Execuções': 120
        }
        
        for col in columns:
            self.tasks_tree.heading(col, text=col)
            self.tasks_tree.column(col, width=column_widths.get(col, 150))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical', command=self.tasks_tree.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal', command=self.tasks_tree.xview)
        
        self.tasks_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tasks_tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Duplo clique para executar
        self.tasks_tree.bind('<Double-1>', lambda event: self.execute_task())

    def create_control_center_tab(self):
        """Cria a aba do controle central (agendador + alertas)"""
        control_frame = ttk.Frame(self.notebook)
        self.notebook.add(control_frame, text="🎛️ Controle Central")
        
        # Notebook interno para sub-abas
        control_notebook = ttk.Notebook(control_frame)
        control_notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Sub-aba: Agendador
        self.create_scheduler_subtab(control_notebook)
        
        # Sub-aba: Alertas
        self.create_alerts_subtab(control_notebook)
        
        # Sub-aba: Notificações
        self.create_notifications_subtab(control_notebook)

    def create_scheduler_subtab(self, parent):
        """Cria a sub-aba do agendador"""
        scheduler_frame = ttk.Frame(parent)
        parent.add(scheduler_frame, text="📅 Agendador")
        
        # Frame de controles
        controls_frame = ttk.Frame(scheduler_frame)
        controls_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(controls_frame, text="➕ Novo Agendamento", 
                  command=self.new_schedule).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="✏️ Editar Agendamento", 
                  command=self.edit_schedule).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="🗑️ Excluir Agendamento", 
                  command=self.delete_schedule).pack(side='left', padx=5)
        
        # Separador
        ttk.Separator(controls_frame, orient='vertical').pack(side='left', fill='y', padx=10)
        
        # Status do agendador
        self.scheduler_status_label = ttk.Label(controls_frame, text="🟢 Agendador Ativo", 
                                               font=('Segoe UI', 10, 'bold'))
        self.scheduler_status_label.pack(side='left', padx=10)
        
        ttk.Button(controls_frame, text="🔄 Atualizar", 
                  command=self.refresh_schedules).pack(side='right', padx=5)
        
        # Frame para próximas execuções
        next_frame = ttk.LabelFrame(scheduler_frame, text="🕐 Próximas Execuções", padding=10)
        next_frame.pack(fill='x', padx=10, pady=5)
        
        # Treeview para próximas execuções
        next_columns = ('Tarefa', 'Agendamento', 'Próxima Execução', 'Tipo', 'Status')
        self.next_tree = ttk.Treeview(next_frame, columns=next_columns, show='headings', height=6)
        
        for col in next_columns:
            self.next_tree.heading(col, text=col)
            self.next_tree.column(col, width=150)
        
        self.next_tree.pack(fill='x', pady=5)
        
        # Frame para agendamentos
        schedule_frame = ttk.LabelFrame(scheduler_frame, text="📋 Agendamentos Configurados", padding=10)
        schedule_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Treeview para agendamentos
        schedule_columns = ('ID', 'Nome', 'Tarefa', 'Tipo', 'Horário', 'Configuração', 'Status', 'Próxima Execução')
        self.schedule_tree = ttk.Treeview(schedule_frame, columns=schedule_columns, show='headings', height=10)
        
        schedule_widths = {'ID': 50, 'Nome': 150, 'Tarefa': 150, 'Tipo': 80, 
                          'Horário': 80, 'Configuração': 120, 'Status': 80, 'Próxima Execução': 150}
        
        for col in schedule_columns:
            self.schedule_tree.heading(col, text=col)
            self.schedule_tree.column(col, width=schedule_widths.get(col, 100))
        
        # Scrollbars para agendamentos
        schedule_v_scroll = ttk.Scrollbar(schedule_frame, orient='vertical', command=self.schedule_tree.yview)
        schedule_h_scroll = ttk.Scrollbar(schedule_frame, orient='horizontal', command=self.schedule_tree.xview)
        
        self.schedule_tree.configure(yscrollcommand=schedule_v_scroll.set, xscrollcommand=schedule_h_scroll.set)
        
        # Layout
        self.schedule_tree.pack(side='left', fill='both', expand=True)
        schedule_v_scroll.pack(side='right', fill='y')
        schedule_h_scroll.pack(side='bottom', fill='x')

    def create_alerts_subtab(self, parent):
        """Cria a sub-aba de alertas"""
        alerts_frame = ttk.Frame(parent)
        parent.add(alerts_frame, text="🚨 Alertas")
        
        # Frame de controles
        controls_frame = ttk.Frame(alerts_frame)
        controls_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(controls_frame, text="✅ Marcar como Resolvido", 
                  command=self.resolve_alert).pack(side='left', padx=5)
        ttk.Button(controls_frame, text="🗑️ Limpar Alertas Resolvidos", 
                  command=self.clear_resolved_alerts).pack(side='left', padx=5)
        
        # Separador
        ttk.Separator(controls_frame, orient='vertical').pack(side='left', fill='y', padx=10)
        
        # Filtros
        ttk.Label(controls_frame, text="Filtro:").pack(side='left', padx=5)
        self.alert_filter_var = tk.StringVar(value="pendentes")
        filter_combo = ttk.Combobox(controls_frame, textvariable=self.alert_filter_var, 
                                   values=["todos", "pendentes", "resolvidos", "erro", "sucesso"],
                                   state='readonly', width=15)
        filter_combo.pack(side='left', padx=5)
        filter_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_alerts())
        
        ttk.Button(controls_frame, text="🔄 Atualizar", 
                  command=self.refresh_alerts).pack(side='right', padx=5)
        
        # Frame para alertas pendentes
        pending_frame = ttk.LabelFrame(alerts_frame, text="⚠️ Alertas Pendentes", padding=10)
        pending_frame.pack(fill='x', padx=10, pady=5)
        
        # Labels de resumo
        self.alert_summary_frame = ttk.Frame(pending_frame)
        self.alert_summary_frame.pack(fill='x', pady=5)
        
        self.alert_labels = {}
        alert_types = [("Erros", "erro"), ("Timeouts", "timeout"), ("Sucessos", "sucesso")]
        
        for i, (label, key) in enumerate(alert_types):
            frame = ttk.Frame(self.alert_summary_frame)
            frame.pack(side='left', padx=20)
            
            ttk.Label(frame, text=label, font=('Segoe UI', 10, 'bold')).pack()
            self.alert_labels[key] = ttk.Label(frame, text="0", font=('Segoe UI', 14, 'bold'))
            self.alert_labels[key].pack()
        
        # Frame para lista de alertas
        alerts_list_frame = ttk.LabelFrame(alerts_frame, text="📋 Lista de Alertas", padding=10)
        alerts_list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Treeview para alertas
        alert_columns = ('ID', 'Tipo', 'Título', 'Tarefa', 'Data/Hora', 'Status', 'Canal')
        self.alerts_tree = ttk.Treeview(alerts_list_frame, columns=alert_columns, show='headings', height=12)
        
        alert_widths = {'ID': 50, 'Tipo': 80, 'Título': 200, 'Tarefa': 150, 
                       'Data/Hora': 150, 'Status': 80, 'Canal': 100}
        
        for col in alert_columns:
            self.alerts_tree.heading(col, text=col)
            self.alerts_tree.column(col, width=alert_widths.get(col, 100))
        
        # Scrollbars para alertas
        alerts_v_scroll = ttk.Scrollbar(alerts_list_frame, orient='vertical', command=self.alerts_tree.yview)
        alerts_h_scroll = ttk.Scrollbar(alerts_list_frame, orient='horizontal', command=self.alerts_tree.xview)
        
        self.alerts_tree.configure(yscrollcommand=alerts_v_scroll.set, xscrollcommand=alerts_h_scroll.set)
        
        # Layout
        self.alerts_tree.pack(side='left', fill='both', expand=True)
        alerts_v_scroll.pack(side='right', fill='y')
        alerts_h_scroll.pack(side='bottom', fill='x')
        
        # Duplo clique para ver detalhes
        self.alerts_tree.bind('<Double-1>', lambda event: self.show_alert_details())

    def create_notifications_subtab(self, parent):
        """Cria a sub-aba de configurações de notificação"""
        notif_frame = ttk.Frame(parent)
        parent.add(notif_frame, text="📧 Notificações")
        
        # Frame de configurações de email
        email_frame = ttk.LabelFrame(notif_frame, text="📧 Configurações de Email", padding=10)
        email_frame.pack(fill='x', padx=10, pady=5)
        
        # Ativar email
        self.email_active_var = tk.BooleanVar()
        ttk.Checkbutton(email_frame, text="Ativar notificações por email", 
                       variable=self.email_active_var).pack(anchor='w', pady=5)
        
        # Configurações de email
        email_config_frame = ttk.Frame(email_frame)
        email_config_frame.pack(fill='x', pady=5)
        
        # Servidor SMTP
        ttk.Label(email_config_frame, text="Servidor SMTP:").grid(row=0, column=0, sticky='w', pady=2)
        self.smtp_server_entry = ttk.Entry(email_config_frame, width=30)
        self.smtp_server_entry.grid(row=0, column=1, padx=(10, 0), pady=2, sticky='ew')
        self.smtp_server_entry.insert(0, "smtp.gmail.com")
        
        ttk.Label(email_config_frame, text="Porta:").grid(row=0, column=2, sticky='w', padx=(10, 0), pady=2)
        self.smtp_port_entry = ttk.Entry(email_config_frame, width=10)
        self.smtp_port_entry.grid(row=0, column=3, padx=(10, 0), pady=2)
        self.smtp_port_entry.insert(0, "587")
        
        # Email e senha
        ttk.Label(email_config_frame, text="Email:").grid(row=1, column=0, sticky='w', pady=2)
        self.email_entry = ttk.Entry(email_config_frame, width=30)
        self.email_entry.grid(row=1, column=1, padx=(10, 0), pady=2, sticky='ew')
        
        ttk.Label(email_config_frame, text="Senha:").grid(row=1, column=2, sticky='w', padx=(10, 0), pady=2)
        self.email_password_entry = ttk.Entry(email_config_frame, width=20, show="*")
        self.email_password_entry.grid(row=1, column=3, padx=(10, 0), pady=2)
        
        # Email destinatário
        ttk.Label(email_config_frame, text="Destinatário:").grid(row=2, column=0, sticky='w', pady=2)
        self.email_to_entry = ttk.Entry(email_config_frame, width=30)
        self.email_to_entry.grid(row=2, column=1, padx=(10, 0), pady=2, sticky='ew')
        
        # Configurar grid
        email_config_frame.columnconfigure(1, weight=1)
        
        # Botões de teste e salvar
        email_buttons_frame = ttk.Frame(email_frame)
        email_buttons_frame.pack(fill='x', pady=10)
        
        ttk.Button(email_buttons_frame, text="📧 Testar Email", 
                  command=self.test_email).pack(side='left', padx=5)
        ttk.Button(email_buttons_frame, text="💾 Salvar Configurações", 
                  command=self.save_notification_config).pack(side='left', padx=5)
        
        # Frame para outras notificações (futuras implementações)
        other_frame = ttk.LabelFrame(notif_frame, text="🔔 Outras Notificações", padding=10)
        other_frame.pack(fill='x', padx=10, pady=5)
        
        # WhatsApp (placeholder)
        self.whatsapp_active_var = tk.BooleanVar()
        ttk.Checkbutton(other_frame, text="WhatsApp (Em desenvolvimento)", 
                       variable=self.whatsapp_active_var, state='disabled').pack(anchor='w', pady=2)
        
        # Telegram (placeholder)
        self.telegram_active_var = tk.BooleanVar()
        ttk.Checkbutton(other_frame, text="Telegram (Em desenvolvimento)", 
                       variable=self.telegram_active_var, state='disabled').pack(anchor='w', pady=2)
        
        # Slack (placeholder)
        self.slack_active_var = tk.BooleanVar()
        ttk.Checkbutton(other_frame, text="Slack (Em desenvolvimento)", 
                       variable=self.slack_active_var, state='disabled').pack(anchor='w', pady=2)
        
        # Carregar configurações existentes
        self.load_notification_configs()

    def create_settings_tab(self):
        """Cria a aba de configurações"""
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="⚙️ Configurações")
        
        # Frame de configurações gerais
        general_frame = ttk.LabelFrame(settings_frame, text="🔧 Configurações Gerais", padding=10)
        general_frame.pack(fill='x', padx=10, pady=5)
        
        # Tema
        ttk.Label(general_frame, text="Tema:").grid(row=0, column=0, sticky='w', pady=5)
        self.theme_var = tk.StringVar(value=self.current_theme)
        theme_combo = ttk.Combobox(general_frame, textvariable=self.theme_var, 
                                  values=['teal'], state='readonly')
        theme_combo.grid(row=0, column=1, sticky='ew', padx=(10, 0), pady=5)
        
        # Configurar grid
        general_frame.columnconfigure(1, weight=1)
        
        # Frame de backup
        backup_frame = ttk.LabelFrame(settings_frame, text="💾 Backup e Restauração", padding=10)
        backup_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(backup_frame, text="📦 Fazer Backup", 
                  command=self.backup_database).pack(side='left', padx=5)
        ttk.Button(backup_frame, text="📥 Restaurar Backup", 
                  command=self.restore_database).pack(side='left', padx=5)
        
        # Frame de informações
        info_frame = ttk.LabelFrame(settings_frame, text="ℹ️ Informações do Sistema", padding=10)
        info_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        info_text = tk.Text(info_frame, height=10, wrap='word', state='disabled',
                           bg='#2d5555', fg='#e8f4f4', insertbackground='#e8f4f4')
        info_text.pack(fill='both', expand=True)
        
        # Adicionar informações do sistema
        self.update_system_info(info_text)

    def update_system_info(self, text_widget):
        """Atualiza informações do sistema"""
        try:
            text_widget.config(state='normal')
            text_widget.delete(1.0, tk.END)
            
            db_type = "SQL Server" if not self.using_sqlite else "SQLite"
            db_server = DatabaseConfig.DB_SERVER if not self.using_sqlite else "Local"
            
            info = f"""Sistema Operacional: {os.name}
Python: {sys.version.split()[0]}
Banco de Dados: {db_type}
Servidor: {db_server}

Dependências:
• matplotlib: ✅ Instalado
• pandas: ✅ Instalado
• schedule: ✅ Instalado
• pyodbc: {'✅ Instalado' if SQLSERVER_AVAILABLE else '❌ Não instalado'}
• sqlalchemy: {'✅ Instalado' if SQLSERVER_AVAILABLE else '❌ Não instalado'}
• psutil: {'✅ Instalado' if PSUTIL_AVAILABLE else '❌ Não instalado'}
• requests: {'✅ Instalado' if REQUESTS_AVAILABLE else '❌ Não instalado'}

Funcionalidades Disponíveis:
• Gerenciamento de Tarefas Automatizadas: ✅
• Execução de Scripts BAT: ✅
• Dashboard com Gráficos: ✅
• Histórico de Execuções: ✅
• Agendador de Tarefas: ✅
• Sistema de Alertas: ✅
• Notificações por Email: ✅
• Backup/Restauração: ✅
• Exportar/Importar Dados: ✅
• Conexão SQL Server: {'✅' if not self.using_sqlite else '❌ (usando SQLite)'}
• Integração Connect:Direct (via API ICC): {'✅' if REQUESTS_AVAILABLE else '❌ (requests não instalado)'}

📧 Suporte a aplicação: seu.email@exemplo.com # 
🌐 Website: seu.website.com # 
🧑‍💻 Desenvolvido por: Usuário # 

© 2025. Todos os direitos reservados. # 
"""
            
            text_widget.insert(1.0, info)
            text_widget.config(state='disabled')
            
        except Exception as e:
            print(f"Erro ao atualizar informações: {e}")

    def load_data(self):
        """Carrega dados iniciais"""
        self.refresh_tasks()
        self.refresh_schedules()
        self.refresh_alerts()
        self.update_dashboard_stats()

    def refresh_tasks(self):
        """Atualiza a lista de tarefas"""
        try:
            # Limpar treeview
            for item in self.tasks_tree.get_children():
                self.tasks_tree.delete(item)
            
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                # SQLite sem SQLAlchemy
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                # QUERY ATUALIZADA PARA NOVOS CAMPOS
                cursor.execute("""
                    SELECT id, titulo, prioridade, arquivo_BAT, data_ultima_execucao, total_execucoes,
                           tipo_execucao, cd_source_path, cd_destination_node, cd_destination_path, cd_process_name
                    FROM painel_tarefas 
                    ORDER BY 
                        CASE prioridade 
                            WHEN 'Alta' THEN 1 
                            WHEN 'Média' THEN 2 
                            WHEN 'Baixa' THEN 3 
                        END, titulo
                """,)
                
                rows = cursor.fetchall()
                conn.close()
                
                for row in rows:
                    # Formatar data da última execução
                    ultima_execucao = "Nunca"
                    if row[4]:
                        try:
                            dt = datetime.datetime.fromisoformat(row[4])
                            ultima_execucao = dt.strftime("%d/%m/%Y %H:%M")
                        except:
                            ultima_execucao = "Erro na data"
                    
                    # Detalhes de configuração (nova coluna)
                    tipo_exec = row[6] if len(row) > 6 else "BAT"
                    detalhes_config = "N/A"
                    if tipo_exec == "BAT":
                        detalhes_config = os.path.basename(row[3]) if row[3] else "N/A"
                    elif tipo_exec == "ConnectDirect":
                        detalhes_config = f"Origem: {row[7] or 'N/A'} | Destino: {row[8] or 'N/A'}:{row[9] or 'N/A'}"
                        if row[10]: # Process name
                            detalhes_config += f" (Proc: {row[10]})"

                    self.tasks_tree.insert('', 'end', values=(
                        row[0], row[1], tipo_exec.title(), row[2], detalhes_config, ultima_execucao, row[5] or 0
                    ))
            else:
                # SQL Server com SQLAlchemy
                session = self.get_db_session()
                try:
                    tarefas = session.query(Tarefa).order_by(
                        Tarefa.prioridade.desc(), 
                        Tarefa.titulo
                    ).all()
                    
                    for tarefa in tarefas:
                        # Formatar data da última execução
                        ultima_execucao = "Nunca"
                        if tarefa.data_ultima_execucao:
                            ultima_execucao = tarefa.data_ultima_execucao.strftime("%d/%m/%Y %H:%M")
                        
                        # Detalhes de configuração (nova coluna)
                        detalhes_config = "N/A"
                        if tarefa.tipo_execucao == "BAT":
                            detalhes_config = os.path.basename(tarefa.arquivo_BAT) if tarefa.arquivo_BAT else "N/A"
                        elif tarefa.tipo_execucao == "ConnectDirect":
                            detalhes_config = f"Origem: {tarefa.cd_source_path or 'N/A'} | Destino: {tarefa.cd_destination_node or 'N/A'}:{tarefa.cd_destination_path or 'N/A'}"
                            if tarefa.cd_process_name:
                                detalhes_config += f" (Proc: {tarefa.cd_process_name})"

                        self.tasks_tree.insert('', 'end', values=(
                            tarefa.id, 
                            tarefa.titulo, 
                            tarefa.tipo_execucao.title(), # Nova coluna
                            tarefa.prioridade, 
                            detalhes_config, # Nova coluna
                            ultima_execucao, 
                            tarefa.total_execucoes or 0
                        ))
                finally:
                    session.close()
            
        except Exception as e:
            print(f"❌ Erro ao carregar tarefas: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar tarefas: {e}")

    def refresh_schedules(self):
        """Atualiza a lista de agendamentos"""
        try:
            # Limpar treeviews
            for item in self.schedule_tree.get_children():
                self.schedule_tree.delete(item)
            
            for item in self.next_tree.get_children():
                self.next_tree.delete(item)
            
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                # SQLite sem SQLAlchemy
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                cursor.execute("""
                    SELECT a.id, a.nome, t.titulo, a.tipo_agendamento, a.horario, 
                           a.dias_semana, a.dia_mes, a.ativo, a.proxima_execucao
                    FROM painel_agendamentos a
                    JOIN painel_tarefas t ON a.tarefa_id = t.id
                    ORDER BY a.proxima_execucao
                """,)
                
                rows = cursor.fetchall()
                conn.close()
                
                for row in rows:
                    # Formatar configuração
                    config = ""
                    if row[3] == "diario":
                        config = "Diário"
                    elif row[3] == "semanal":
                        dias = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"]
                        if row[5]:
                            dias_selecionados = [dias[int(d)] for d in row[5].split(',') if d.strip()]
                            config = ", ".join(dias_selecionados)
                    elif row[3] == "mensal":
                        config = f"Dia {row[6]}"
                    elif row[3] == "unico":
                        config = "Execução única"
                    
                    # Status
                    status = "✅ Ativo" if row[7] else "❌ Inativo"
                    
                    # Próxima execução
                    proxima = "N/A"
                    if row[8]:
                        try:
                            dt = datetime.datetime.fromisoformat(row[8])
                            proxima = dt.strftime("%d/%m/%Y %H:%M")
                        except:
                            proxima = "Erro na data"
                    
                    self.schedule_tree.insert('', 'end', values=(
                        row[0], row[1], row[2], row[3].title(), row[4], config, status, proxima
                    ))
                    
                    # Adicionar às próximas execuções se ativo
                    if row[7] and row[8]:
                        try:
                            dt = datetime.datetime.fromisoformat(row[8])
                            if dt > datetime.datetime.now():
                                self.next_tree.insert('', 'end', values=(
                                    row[2], row[1], proxima, row[3].title(), status
                                ))
                        except:
                            pass
            else:
                # SQL Server com SQLAlchemy
                session = self.get_db_session()
                try:
                    agendamentos = session.query(Agendamento).join(Tarefa).order_by(
                        Agendamento.proxima_execucao
                    ).all()
                    
                    for agendamento in agendamentos:
                        # Formatar configuração
                        config = ""
                        if agendamento.tipo_agendamento == "diario":
                            config = "Diário"
                        elif agendamento.tipo_agendamento == "semanal":
                            dias = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"]
                            if agendamento.dias_semana:
                                dias_selecionados = [dias[int(d)] for d in agendamento.dias_semana.split(',') if d.strip()]
                                config = ", ".join(dias_selecionados)
                        elif agendamento.tipo_agendamento == "mensal":
                            config = f"Dia {agendamento.dia_mes}"
                        elif agendamento.tipo_agendamento == "unico":
                            config = "Execução única"
                        
                        # Status
                        status = "✅ Ativo" if agendamento.ativo else "❌ Inativo"
                        
                        # Próxima execução
                        proxima = "N/A"
                        if agendamento.proxima_execucao:
                            proxima = agendamento.proxima_execucao.strftime("%d/%m/%Y %H:%M")
                        
                        self.schedule_tree.insert('', 'end', values=(
                            agendamento.id, 
                            agendamento.nome, 
                            agendamento.tarefa.titulo, 
                            agendamento.tipo_agendamento.title(), 
                            agendamento.horario,
                            config, 
                            status, 
                            proxima
                        ))
                        
                        # Adicionar às próximas execuções se ativo
                        if agendamento.ativo and agendamento.proxima_execucao:
                            if agendamento.proxima_execucao > datetime.datetime.now():
                                self.next_tree.insert('', 'end', values=(
                                    agendamento.tarefa.titulo, 
                                    agendamento.nome, 
                                    proxima, 
                                    agendamento.tipo_agendamento.title(), 
                                    status
                                ))
                finally:
                    session.close()
            
        except Exception as e:
            print(f"❌ Erro ao carregar agendamentos: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar agendamentos: {e}")

    def refresh_alerts(self):
        """Atualiza a lista de alertas"""
        try:
            # Limpar treeview
            for item in self.alerts_tree.get_children():
                self.alerts_tree.delete(item)
            
            # Aplicar filtro
            filter_value = self.alert_filter_var.get()
            where_clause = ""
            
            if filter_value == "pendentes":
                where_clause = "WHERE resolvido = 0"
            elif filter_value == "resolvidos":
                where_clause = "WHERE resolvido = 1"
            elif filter_value == "erro":
                where_clause = "WHERE tipo = 'erro'"
            elif filter_value == "sucesso":
                where_clause = "WHERE tipo = 'sucesso'"
            
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                # SQLite sem SQLAlchemy
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                cursor.execute(f"""
                    SELECT a.id, a.tipo, a.titulo, t.titulo, a.data_criacao, 
                           a.resolvido, a.canal_notificacao
                    FROM painel_alertas a
                    LEFT JOIN painel_tarefas t ON a.tarefa_id = t.id
                    {where_clause}
                    ORDER BY a.data_criacao DESC
                    LIMIT 100
                """,)
                
                rows = cursor.fetchall()
                
                # Contar por tipo
                cursor.execute("SELECT tipo, COUNT(*) FROM painel_alertas WHERE resolvido = 0 GROUP BY tipo")
                counts = dict(cursor.fetchall())
                
                conn.close()
                
                for row in rows:
                    # Formatar data
                    data_formatada = "N/A"
                    if row[4]:
                        try:
                            dt = datetime.datetime.fromisoformat(row[4])
                            data_formatada = dt.strftime("%d/%m/%Y %H:%M")
                        except:
                            data_formatada = "Erro na data"
                    
                    # Status
                    status = "✅ Resolvido" if row[5] else "⚠️ Pendente"
                    
                    # Tipo com emoji
                    tipo_display = row[1]
                    if row[1] == "erro":
                        tipo_display = "❌ Erro"
                    elif row[1] == "sucesso":
                        tipo_display = "✅ Sucesso"
                    elif row[1] == "timeout":
                        tipo_display = "⏰ Timeout"
                    
                    self.alerts_tree.insert('', 'end', values=(
                        row[0], tipo_display, row[2], row[3] or "N/A", 
                        data_formatada, status, row[6] or "N/A"
                    ))
                
                # Atualizar labels de resumo
                self.alert_labels['erro'].config(text=str(counts.get('erro', 0)))
                self.alert_labels['timeout'].config(text=str(counts.get('timeout', 0)))
                self.alert_labels['sucesso'].config(text=str(counts.get('sucesso', 0)))
                
            else:
                # SQL Server com SQLAlchemy
                session = self.get_db_session()
                try:
                    query = session.query(Alerta).outerjoin(Tarefa)
                    
                    if filter_value == "pendentes":
                        query = query.filter(Alerta.resolvido == False)
                    elif filter_value == "resolvidos":
                        query = query.filter(Alerta.resolvido == True)
                    elif filter_value == "erro":
                        query = query.filter(Alerta.tipo == 'erro')
                    elif filter_value == "sucesso":
                        query = query.filter(Alerta.tipo == 'sucesso')
                    
                    alertas = query.order_by(Alerta.data_criacao.desc()).limit(100).all()
                    
                    for alerta in alertas:
                        # Formatar data
                        data_formatada = alerta.data_criacao.strftime("%d/%m/%Y %H:%M")
                        
                        # Status
                        status = "✅ Resolvido" if alerta.resolvido else "⚠️ Pendente"
                        
                        # Tipo com emoji
                        tipo_display = alerta.tipo
                        if alerta.tipo == "erro":
                            tipo_display = "❌ Erro"
                        elif alerta.tipo == "sucesso":
                            tipo_display = "✅ Sucesso"
                        elif alerta.tipo == "timeout":
                            tipo_display = "⏰ Timeout"
                        
                        # Nome da tarefa
                        tarefa_nome = "N/A"
                        if alerta.tarefa_id:
                            tarefa = session.query(Tarefa).filter(Tarefa.id == alerta.tarefa_id).first()
                            if tarefa:
                                tarefa_nome = tarefa.titulo
                        
                        self.alerts_tree.insert('', 'end', values=(
                            alerta.id, 
                            tipo_display, 
                            alerta.titulo, 
                            tarefa_nome, 
                            data_formatada, 
                            status, 
                            alerta.canal_notificacao or "N/A"
                        ))
                    
                    # Contar por tipo
                    counts = {}
                    for tipo in ['erro', 'timeout', 'sucesso']:
                        count = session.query(Alerta).filter(
                            Alerta.tipo == tipo, 
                            Alerta.resolvido == False
                        ).count()
                        counts[tipo] = count
                    
                    # Atualizar labels de resumo
                    self.alert_labels['erro'].config(text=str(counts.get('erro', 0)))
                    self.alert_labels['timeout'].config(text=str(counts.get('timeout', 0)))
                    self.alert_labels['sucesso'].config(text=str(counts.get('sucesso', 0)))
                    
                finally:
                    session.close()
            
        except Exception as e:
            print(f"❌ Erro ao carregar alertas: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar alertas: {e}")

    def update_dashboard_stats(self):
        """Atualiza as estatísticas do dashboard"""
        try:
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                # Total de tarefas configuradas
                cursor.execute("SELECT COUNT(*) FROM painel_tarefas")
                total_tasks = cursor.fetchone()[0]
                
                # Agendamentos ativos
                cursor.execute("SELECT COUNT(*) FROM painel_agendamentos WHERE ativo = 1")
                active_schedules = cursor.fetchone()[0]
                
                # Execuções hoje
                cursor.execute("""
                    SELECT COUNT(*) FROM painel_execucoes 
                    WHERE DATE(data_execucao) = DATE('now')
                """,)
                executions_today = cursor.fetchone()[0]
                
                # Alertas pendentes
                cursor.execute("SELECT COUNT(*) FROM painel_alertas WHERE resolvido = 0")
                pending_alerts = cursor.fetchone()[0]
                
                # Última execução
                cursor.execute("SELECT MAX(data_execucao) FROM painel_execucoes")
                last_execution = cursor.fetchone()[0]
                
                # Taxa de sucesso
                cursor.execute("SELECT COUNT(*) FROM painel_execucoes WHERE status = 'sucesso'")
                sucessos = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM painel_execucoes")
                total_execucoes = cursor.fetchone()[0]
                
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    # Total de tarefas configuradas
                    total_tasks = session.query(Tarefa).count()
                    
                    # Agendamentos ativos
                    active_schedules = session.query(Agendamento).filter(Agendamento.ativo == True).count()
                    
                    # Execuções hoje
                    from sqlalchemy import func, cast, Date
                    today = datetime.date.today()
                    executions_today = session.query(Execucao).filter(
                        cast(Execucao.data_execucao, Date) == today
                    ).count()
                    
                    # Alertas pendentes
                    pending_alerts = session.query(Alerta).filter(Alerta.resolvido == False).count()
                    
                    # Última execução
                    last_execution_obj = session.query(Execucao).order_by(
                        Execucao.data_execucao.desc()
                    ).first()
                    last_execution = last_execution_obj.data_execucao if last_execution_obj else None
                    
                    # Taxa de sucesso
                    sucessos = session.query(Execucao).filter(Execucao.status == 'sucesso').count()
                    total_execucoes = session.query(Execucao).count()
                    
                finally:
                    session.close()
            
            # Atualizar labels
            self.stats_labels['tasks_total'].config(text=str(total_tasks))
            self.stats_labels['schedules_active'].config(text=str(active_schedules))
            self.stats_labels['executions_today'].config(text=str(executions_today))
            self.stats_labels['alerts_pending'].config(text=str(pending_alerts))
            
            # Última execução
            if last_execution:
                try:
                    if isinstance(last_execution, str):
                        dt = datetime.datetime.fromisoformat(last_execution)
                    else:
                        dt = last_execution
                    last_exec_text = dt.strftime("%d/%m %H:%M")
                except:
                    last_exec_text = "Erro na data"
            else:
                last_exec_text = "Nunca"
            self.stats_labels['last_execution'].config(text=last_exec_text)
            
            # Taxa de sucesso
            if total_execucoes > 0:
                taxa_sucesso = (sucessos / total_execucoes) * 100
                self.stats_labels['success_rate'].config(text=f"{taxa_sucesso:.1f}%")
            else:
                self.stats_labels['success_rate'].config(text="0%")
            
        except Exception as e:
            print(f"Erro ao atualizar estatísticas: {e}")

    # ==================== FUNÇÕES DE TAREFAS ====================

    def new_task(self):
        """Cria uma nova tarefa"""
        dialog = TaskDialog(self.root, "Nova Tarefa Automatizada")
        if dialog.result:
            try:
                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                    # SQLite sem SQLAlchemy
                    import sqlite3
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    
                    # INSERT ATUALIZADO PARA NOVOS CAMPOS
                    cursor.execute("""
                        INSERT INTO painel_tarefas (titulo, descricao, prioridade, arquivo_BAT, tipo_execucao,
                                                   cd_source_path, cd_destination_node, cd_destination_path, cd_process_name) 
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (dialog.result['titulo'], dialog.result['descricao'], 
                         dialog.result['prioridade'], dialog.result['arquivo_BAT'], dialog.result['tipo_execucao'],
                         dialog.result['cd_source_path'], dialog.result['cd_destination_node'],
                         dialog.result['cd_destination_path'], dialog.result['cd_process_name']))
                    
                    conn.commit()
                    conn.close()
                else:
                    # SQL Server com SQLAlchemy
                    session = self.get_db_session()
                    try:
                        nova_tarefa = Tarefa(
                            titulo=dialog.result['titulo'],
                            descricao=dialog.result['descricao'],
                            prioridade=dialog.result['prioridade'],
                            arquivo_BAT=dialog.result['arquivo_BAT'],
                            tipo_execucao=dialog.result['tipo_execucao'], # Novo campo
                            cd_source_path=dialog.result['cd_source_path'], # Novo campo
                            cd_destination_node=dialog.result['cd_destination_node'], # Novo campo
                            cd_destination_path=dialog.result['cd_destination_path'], # Novo campo
                            cd_process_name=dialog.result['cd_process_name'] # Novo campo
                        )
                        
                        session.add(nova_tarefa)
                        session.commit()
                    finally:
                        session.close()
                
                self.refresh_tasks()
                self.update_dashboard_stats()
                messagebox.showinfo("Sucesso", "Tarefa criada com sucesso!")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao criar tarefa: {e}")

    def edit_task(self):
        """Edita a tarefa selecionada"""
        selection = self.tasks_tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma tarefa para editar.")
            return
        
        item = self.tasks_tree.item(selection[0])
        task_id = item['values'][0]
        
        try:
            # Buscar dados da tarefa
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                # SELECT ATUALIZADO PARA NOVOS CAMPOS
                cursor.execute("SELECT * FROM painel_tarefas WHERE id = ?", (task_id,))
                row = cursor.fetchone()
                conn.close()
                
                if row:
                    # Criar objeto mock para compatibilidade
                    class MockTask:
                        def __init__(self, row_data):
                            self.id = row_data[0]
                            self.titulo = row_data[1]
                            self.descricao = row_data[2]
                            self.status = row_data[3]
                            self.prioridade = row_data[4]
                            self.arquivo_BAT = row_data[5]
                            self.data_criacao = row_data[6]
                            self.data_ultima_execucao = row_data[7]
                            self.total_execucoes = row_data[8]
                            self.tipo_execucao = row_data[9] if len(row_data) > 9 else "BAT" # Nova coluna
                            self.cd_source_path = row_data[10] if len(row_data) > 10 else None # Nova coluna
                            self.cd_destination_node = row_data[11] if len(row_data) > 11 else None # Nova coluna
                            self.cd_destination_path = row_data[12] if len(row_data) > 12 else None # Nova coluna
                            self.cd_process_name = row_data[13] if len(row_data) > 13 else None # Nova coluna

                    task_data = MockTask(row)
                else:
                    messagebox.showerror("Erro", "Tarefa não encontrada.")
                    return
            else:
                session = self.get_db_session()
                try:
                    task_data = session.query(Tarefa).filter(Tarefa.id == task_id).first()
                    if not task_data:
                        messagebox.showerror("Erro", "Tarefa não encontrada.")
                        return
                finally:
                    session.close()
            
            # Abrir diálogo de edição
            dialog = TaskDialog(self.root, "Editar Tarefa", task_data)
            if dialog.result:
                # Atualizar tarefa
                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    # UPDATE ATUALIZADO PARA NOVOS CAMPOS
                    cursor.execute("""
                        UPDATE painel_tarefas 
                        SET titulo = ?, descricao = ?, prioridade = ?, arquivo_BAT = ?, tipo_execucao = ?,
                            cd_source_path = ?, cd_destination_node = ?, cd_destination_path = ?, cd_process_name = ?
                        WHERE id = ?
                    """, (dialog.result['titulo'], dialog.result['descricao'], 
                         dialog.result['prioridade'], dialog.result['arquivo_BAT'], dialog.result['tipo_execucao'],
                         dialog.result['cd_source_path'], dialog.result['cd_destination_node'],
                         dialog.result['cd_destination_path'], dialog.result['cd_process_name'], task_id))
                    conn.commit()
                    conn.close()
                else:
                    session = self.get_db_session()
                    try:
                        tarefa = session.query(Tarefa).filter(Tarefa.id == task_id).first()
                        if tarefa:
                            tarefa.titulo = dialog.result['titulo']
                            tarefa.descricao = dialog.result['descricao']
                            tarefa.prioridade = dialog.result['prioridade']
                            tarefa.arquivo_BAT = dialog.result['arquivo_BAT']
                            tarefa.tipo_execucao = dialog.result['tipo_execucao'] # Novo campo
                            tarefa.cd_source_path = dialog.result['cd_source_path'] # Novo campo
                            tarefa.cd_destination_node = dialog.result['cd_destination_node'] # Novo campo
                            tarefa.cd_destination_path = dialog.result['cd_destination_path'] # Novo campo
                            tarefa.cd_process_name = dialog.result['cd_process_name'] # Novo campo
                            session.commit()
                    finally:
                        session.close()
                
                self.refresh_tasks()
                messagebox.showinfo("Sucesso", "Tarefa atualizada com sucesso!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao editar tarefa: {e}")

    def execute_task(self):
        """Executa a tarefa selecionada"""
        selection = self.tasks_tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma tarefa para executar.")
            return
        
        item = self.tasks_tree.item(selection[0])
        task_id = item['values'][0]
        # task_title e outros detalhes serão buscados no execute_task_by_id

        self.execute_task_by_id(task_id, executado_por_agendador=False)

    def execute_task_by_id(self, task_id, executado_por_agendador=True):
        """Executa uma tarefa pelo ID"""
        start_time = time.time()
        start_datetime = datetime.datetime.now()
        status = "erro" # Default
        codigo_retorno = -1
        log_output = "N/A"
        task_title = "N/A"
        
        try:
            # Buscar dados da tarefa com os novos campos
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT titulo, arquivo_BAT, tipo_execucao, cd_source_path, cd_destination_node, cd_destination_path, cd_process_name FROM painel_tarefas WHERE id = ?", (task_id,))
                result = cursor.fetchone()
                conn.close()
                
                if not result:
                    raise Exception("Tarefa não encontrada.")
                
                task_title = result[0]
                BAT_file = result[1]
                tipo_execucao = result[2] if len(result) > 2 else "BAT" # Compatibilidade
                cd_source_path = result[3] if len(result) > 3 else None
                cd_destination_node = result[4] if len(result) > 4 else None
                cd_destination_path = result[5] if len(result) > 5 else None
                cd_process_name = result[6] if len(result) > 6 else None
                
            else: # SQL Server com SQLAlchemy
                session = self.get_db_session()
                try:
                    tarefa = session.query(Tarefa).filter(Tarefa.id == task_id).first()
                    if not tarefa:
                        raise Exception("Tarefa não encontrada.")
                    
                    task_title = tarefa.titulo
                    BAT_file = tarefa.arquivo_BAT
                    tipo_execucao = getattr(tarefa, 'tipo_execucao', "BAT") # Usa 'BAT' como default se a coluna não existir
                    cd_source_path = getattr(tarefa, 'cd_source_path', None)
                    cd_destination_node = getattr(tarefa, 'cd_destination_node', None)
                    cd_destination_path = getattr(tarefa, 'cd_destination_path', None)
                    cd_process_name = getattr(tarefa, 'cd_process_name', None)
                finally:
                    session.close()
            
            # Confirmar execução se manual (apenas a mensagem muda)
            if not executado_por_agendador:
                confirm_msg = ""
                if tipo_execucao == 'BAT':
                    confirm_msg = f"Executar a tarefa '{task_title}'?\n\nArquivo: {os.path.basename(BAT_file)}"
                elif tipo_execucao == 'ConnectDirect':
                    confirm_msg = f"Iniciar transferência Connect:Direct para '{task_title}'?\n\nOrigem: {cd_source_path}\nDestino: {cd_destination_node}:{cd_destination_path}"
                else:
                    confirm_msg = f"Executar a tarefa '{task_title}' (Tipo: {tipo_execucao})?"

                if not messagebox.askyesno("Confirmar Execução", confirm_msg):
                    return False
            
            # --- Lógica de Execução baseada no tipo ---
            if tipo_execucao == 'BAT':
                if not BAT_file or not os.path.exists(BAT_file):
                    raise FileNotFoundError(f"Arquivo BAT não existe ou não foi informado: {BAT_file}")

                result_subprocess = subprocess.run(
                    [BAT_file], 
                    capture_output=True, 
                    text=True, 
                    timeout=300,  # 5 minutos de timeout
                    cwd=os.path.dirname(BAT_file)
                )
                status = "sucesso" if result_subprocess.returncode == 0 else "erro"
                codigo_retorno = result_subprocess.returncode
                log_output = f"STDOUT:\n{result_subprocess.stdout}\n\nSTDERR:\n{result_subprocess.stderr}"

            elif tipo_execucao == 'ConnectDirect':
                if not all([cd_source_path, cd_destination_node, cd_destination_path]):
                    raise ValueError("Parâmetros do Connect:Direct (caminho de origem, nó de destino, caminho de destino) não configurados para esta tarefa.")
                
                # --- PREPARAR O PAYLOAD CONFORME A DOCUMENTAÇÃO DA API DO ICC ---
                # ESTA ESTRUTURA É UM EXEMPLO COMUM. AJUSTE EXATAMENTE CONFORME A DOC DA IBM!
                transfer_payload = {
                    "source": {
                        "node": "LOCAL_CD_NODE_NAME", # Obtenha este valor da sua configuração do ICC
                        "path": cd_source_path,
                        "fileFormat": "binary"
                    },
                    "destination": {
                        "node": cd_destination_node,
                        "path": cd_destination_path,
                        "fileFormat": "binary"
                    },
                    "processName": cd_process_name if cd_process_name else "DEFAULT_CD_PROCESS", # Nome do processo CD ou um default
                    "metadata": {
                        "originApp": "PainelDesktop",
                        "taskId": task_id,
                        "taskTitle": task_title
                    },
                    "options": {
                        "overwrite": True,
                        "compress": False # Exemplo, ajuste conforme necessidade
                    }
                }
                
                # O endpoint para iniciar transferências via ICC pode variar. Consulte a documentação.
                # Exemplo: /transfer_requests, /filetransfers
                icc_endpoint = "/filetransfers" # MOCK: Substitua pelo endpoint real da API do ICC
                
                success_api, api_response = call_icc_api(
                    method="POST", # Geralmente POST para iniciar algo
                    endpoint_path=icc_endpoint,
                    data=transfer_payload,
                    auth_type="basic" # MOCK: Substitua pelo tipo de autenticação real (basic, bearer)
                )
                
                if success_api:
                    status = "sucesso"
                    codigo_retorno = 0
                    log_output = json.dumps(api_response, indent=2, ensure_ascii=False)
                else:
                    status = "erro"
                    codigo_retorno = api_response.get("status_code", -1) # Tenta pegar o código HTTP ou -1
                    log_output = json.dumps(api_response, indent=2, ensure_ascii=False) # Armazena a resposta completa da API
                
            else:
                raise ValueError(f"Tipo de execução desconhecido: {tipo_execucao}")

        except FileNotFoundError as e:
            log_output = str(e)
            status = "erro"
        except ValueError as e:
            log_output = str(e)
            status = "erro"
        except subprocess.TimeoutExpired as e:
            log_output = f"A execução da tarefa excedeu o tempo limite de 5 minutos. STDOUT: {e.stdout}, STDERR: {e.stderr}"
            status = "timeout"
            codigo_retorno = -1
        except Exception as e:
            log_output = f"Erro inesperado durante a execução: {str(e)}"
            status = "erro"
            codigo_retorno = -1
            print(f"Erro detalhado na execução da tarefa {task_id}: {log_output}")
            import traceback
            traceback.print_exc()

        duration = time.time() - start_time

        # Registrar execução no banco
        if self.using_sqlite and not SQLSERVER_AVAILABLE:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute("""
                INSERT INTO painel_execucoes (tarefa_id, status, codigo_retorno, duracao_segundos, log_output, executado_por_agendador)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (task_id, status, codigo_retorno, duration, log_output, executado_por_agendador))
            
            cursor.execute("""
                UPDATE painel_tarefas 
                SET data_ultima_execucao = ?, total_execucoes = total_execucoes + 1
                WHERE id = ?
            """, (start_datetime.isoformat(), task_id))
            
            conn.commit()
            conn.close()
        else:
            session = self.get_db_session()
            try:
                # Registrar execução
                nova_execucao = Execucao(
                    tarefa_id=task_id,
                    status=status,
                    codigo_retorno=codigo_retorno,
                    duracao_segundos=duration,
                    log_output=log_output,
                    executado_por_agendador=executado_por_agendador
                )
                session.add(nova_execucao)
                
                # Atualizar tarefa
                tarefa_obj = session.query(Tarefa).filter(Tarefa.id == task_id).first()
                if tarefa_obj:
                    tarefa_obj.data_ultima_execucao = start_datetime
                    tarefa_obj.total_execucoes = (tarefa_obj.total_execucoes or 0) + 1
                
                session.commit()
            finally:
                session.close()
        
        # Criar alerta baseado no resultado
        if status == "sucesso":
            self.create_alert("sucesso", f"Tarefa executada com sucesso: {task_title}", 
                            f"Duração: {duration:.2f} segundos", task_id)
            
            if not executado_por_agendador:
                messagebox.showinfo("Execução Concluída", 
                                    f"Tarefa '{task_title}' executada com sucesso!\n\n"
                                    f"Duração: {duration:.2f} segundos\n"
                                    f"Código de retorno: {codigo_retorno}")
        else:
            self.create_alert(status, f"Falha na execução: {task_title}", 
                            f"Código: {codigo_retorno}\nErro: {log_output[:200]}...", task_id)
            
            if not executado_por_agendador:
                messagebox.showerror("Erro na Execução", 
                                    f"Tarefa '{task_title}' falhou!\n\n"
                                    f"Código de retorno: {codigo_retorno}\n"
                                    f"Duração: {duration:.2f} segundos\n\n"
                                    f"Erro: {log_output[:200]}...")
        
        # Atualizar interface
        if not executado_por_agendador:
            self.refresh_tasks()
            self.update_dashboard_stats()
        
        return status == "sucesso"

    def delete_task(self):
        """Exclui a tarefa selecionada"""
        selection = self.tasks_tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma tarefa para excluir.")
            return
        
        item = self.tasks_tree.item(selection[0])
        task_id = item['values'][0]
        task_title = item['values'][1]
        
        if messagebox.askyesno("Confirmar Exclusão", 
                             f"Deseja realmente excluir a tarefa '{task_title}'?\n\n"
                             "Isso também excluirá todo o histórico de execuções e agendamentos."):
            try:
                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                    import sqlite3
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    
                    # Excluir relacionados
                    cursor.execute("DELETE FROM painel_execucoes WHERE tarefa_id = ?", (task_id,))
                    cursor.execute("DELETE FROM painel_agendamentos WHERE tarefa_id = ?", (task_id,))
                    cursor.execute("DELETE FROM painel_alertas WHERE tarefa_id = ?", (task_id,))
                    
                    # Excluir tarefa
                    cursor.execute("DELETE FROM painel_tarefas WHERE id = ?", (task_id,))
                    
                    conn.commit()
                    conn.close()
                else:
                    session = self.get_db_session()
                    try:
                        # Excluir tarefa (relacionados serão excluídos automaticamente por cascade)
                        tarefa = session.query(Tarefa).filter(Tarefa.id == task_id).first()
                        if tarefa:
                            session.delete(tarefa)
                            session.commit()
                    finally:
                        session.close()
                
                self.refresh_tasks()
                self.refresh_schedules()
                self.refresh_alerts()
                self.update_dashboard_stats()
                messagebox.showinfo("Sucesso", "Tarefa excluída com sucesso!")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao excluir tarefa: {e}")

    def show_execution_history(self):
        """Mostra histórico de execuções"""
        selection = self.tasks_tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione uma tarefa para ver o histórico.")
            return
        
        item = self.tasks_tree.item(selection[0])
        task_id = item['values'][0]
        task_title = item['values'][1] # Ajustado para pegar o título da coluna correta
        
        # Criar janela de histórico
        history_window = tk.Toplevel(self.root)
        history_window.title(f"Histórico de Execuções - {task_title}")
        history_window.geometry("900x600")
        history_window.transient(self.root)
        history_window.configure(bg='#1e3a3a')
        
        # Frame principal
        main_frame = ttk.Frame(history_window, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(main_frame, text=f"Histórico de Execuções: {task_title}", 
                 font=('Segoe UI', 14, 'bold')).pack(pady=(0, 10))
        
        # Treeview para histórico
        columns = ('Data/Hora', 'Status', 'Duração', 'Código Retorno', 'Origem')
        history_tree = ttk.Treeview(main_frame, columns=columns, show='headings', height=15)
        
        column_widths = {'Data/Hora': 150, 'Status': 100, 'Duração': 100, 
                        'Código Retorno': 120, 'Origem': 150}
        
        for col in columns:
            history_tree.heading(col, text=col)
            self.tasks_tree.column(col, width=column_widths.get(col, 150)) # Mantido tasks_tree por engano no original, corrigido para history_tree
            history_tree.column(col, width=column_widths.get(col, 100)) # Corrigido
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(main_frame, orient='vertical', command=history_tree.yview)
        h_scroll = ttk.Scrollbar(main_frame, orient='horizontal', command=history_tree.xview)
        history_tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # Layout
        history_tree.pack(side='left', fill='both', expand=True)
        v_scroll.pack(side='right', fill='y')
        h_scroll.pack(side='bottom', fill='x')
        
        # Carregar dados
        try:
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT data_execucao, status, duracao_segundos, codigo_retorno, 
                           log_output, executado_por_agendador
                    FROM painel_execucoes 
                    WHERE tarefa_id = ? 
                    ORDER BY data_execucao DESC
                """, (task_id,))
                executions = cursor.fetchall()
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    executions = session.query(Execucao).filter(
                        Execucao.tarefa_id == task_id
                    ).order_by(Execucao.data_execucao.desc()).all()
                    
                    # Converter para tuplas para compatibilidade
                    executions = [(e.data_execucao, e.status, e.duracao_segundos, 
                                 e.codigo_retorno, e.log_output, e.executado_por_agendador) 
                                 for e in executions]
                finally:
                    session.close()
            
            for exec_data in executions:
                # Formatar data
                try:
                    if isinstance(exec_data[0], str):
                        dt = datetime.datetime.fromisoformat(exec_data[0])
                    else:
                        dt = exec_data[0]
                    data_formatada = dt.strftime("%d/%m/%Y %H:%M:%S")
                except:
                    data_formatada = str(exec_data[0])
                
                # Formatar duração
                duracao = f"{exec_data[2]:.2f}s" if exec_data[2] else "N/A"
                
                # Status com emoji
                status = exec_data[1]
                if status == "sucesso":
                    status_display = "✅ Sucesso"
                elif status == "erro":
                    status_display = "❌ Erro"
                elif status == "timeout":
                    status_display = "⏰ Timeout"
                else:
                    status_display = status
                
                # Origem
                origem = "🤖 Agendador" if exec_data[5] else "👤 Manual"
                
                history_tree.insert('', 'end', values=(
                    data_formatada, status_display, duracao, exec_data[3], origem
                ))
            
            if not executions:
                ttk.Label(main_frame, text="Nenhuma execução encontrada para esta tarefa.", 
                         font=('Segoe UI', 10)).pack(pady=20)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar histórico: {e}")
        
        # Botão fechar
        ttk.Button(main_frame, text="Fechar", command=history_window.destroy).pack(pady=10)

    # ==================== FUNÇÕES DE AGENDAMENTO ====================

    def new_schedule(self):
        """Cria um novo agendamento"""
        # Primeiro verificar se há tarefas disponíveis
        if self.using_sqlite and not SQLSERVER_AVAILABLE:
            import sqlite3
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT id, titulo FROM painel_tarefas ORDER BY titulo")
            tasks = cursor.fetchall()
            conn.close()
        else:
            session = self.get_db_session()
            try:
                tasks = [(t.id, t.titulo) for t in session.query(Tarefa).order_by(Tarefa.titulo).all()]
            finally:
                session.close()
        
        if not tasks:
            messagebox.showwarning("Aviso", "Não há tarefas cadastradas. Crie uma tarefa primeiro.")
            return
        
        # Mostrar seletor de tarefa
        task_window = tk.Toplevel(self.root)
        task_window.title("Selecionar Tarefa")
        task_window.geometry("400x300")
        task_window.transient(self.root)
        task_window.grab_set()
        task_window.configure(bg='#1e3a3a')
        
        # Centralizar
        task_window.update_idletasks()
        x = (task_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (task_window.winfo_screenheight() // 2) - (300 // 2)
        task_window.geometry(f"400x300+{x}+{y}")
        
        main_frame = ttk.Frame(task_window, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(main_frame, text="Selecione a tarefa para agendar:", 
                 font=('Segoe UI', 12, 'bold')).pack(pady=(0, 10))
        
        # Lista de tarefas
        task_listbox = tk.Listbox(main_frame, height=10, bg='#2d5555', fg='#e8f4f4', 
                                 selectbackground='#4a9b9b', selectforeground='#1e3a3a')
        task_listbox.pack(fill='both', expand=True, pady=(0, 10))
        
        for task_id, task_title in tasks:
            task_listbox.insert(tk.END, f"{task_id} - {task_title}")
        
        selected_task_id = None
        
        def on_select():
            nonlocal selected_task_id
            selection = task_listbox.curselection()
            if selection:
                selected_task_id = tasks[selection[0]][0]
                task_window.destroy()
        
        def on_cancel():
            task_window.destroy()
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x')
        
        ttk.Button(btn_frame, text="Selecionar", command=on_select).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=on_cancel).pack(side='left', padx=5)
        
        # Duplo clique para selecionar
        task_listbox.bind('<Double-1>', lambda e: on_select())
        
        task_window.wait_window()
        
        if selected_task_id:
            # Abrir diálogo de agendamento
            dialog = ScheduleDialog(self.root, "Novo Agendamento", selected_task_id)
            if dialog.result:
                try:
                    # Calcular próxima execução
                    proxima_execucao = self.calculate_next_execution(dialog.result)
                    
                    if self.using_sqlite and not SQLSERVER_AVAILABLE:
                        import sqlite3
                        conn = sqlite3.connect(self.db_path)
                        cursor = conn.cursor()
                        
                        cursor.execute("""
                            INSERT INTO painel_agendamentos 
                            (tarefa_id, nome, tipo_agendamento, horario, dias_semana, dia_mes, 
                             data_especifica, max_retries, notificar_sucesso, notificar_erro, proxima_execucao) 
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (selected_task_id, dialog.result['nome'], dialog.result['tipo_agendamento'],
                             dialog.result['horario'], dialog.result['dias_semana'], dialog.result['dia_mes'],
                             dialog.result['data_especifica'].isoformat() if dialog.result['data_especifica'] else None,
                             dialog.result['max_retries'], dialog.result['notificar_sucesso'], 
                             dialog.result['notificar_erro'], proxima_execucao.isoformat() if proxima_execucao else None))
                        
                        conn.commit()
                        conn.close()
                    else:
                        session = self.get_db_session()
                        try:
                            novo_agendamento = Agendamento(
                                tarefa_id=selected_task_id,
                                nome=dialog.result['nome'],
                                tipo_agendamento=dialog.result['tipo_agendamento'],
                                horario=dialog.result['horario'],
                                dias_semana=dialog.result['dias_semana'],
                                dia_mes=dialog.result['dia_mes'],
                                data_especifica=dialog.result['data_especifica'],
                                max_retries=dialog.result['max_retries'],
                                notificar_sucesso=dialog.result['notificar_sucesso'],
                                notificar_erro=dialog.result['notificar_erro'],
                                proxima_execucao=proxima_execucao
                            )
                            
                            session.add(novo_agendamento)
                            session.commit()
                        finally:
                            session.close()
                    
                    self.refresh_schedules()
                    self.update_dashboard_stats()
                    messagebox.showinfo("Sucesso", "Agendamento criado com sucesso!")
                    
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao criar agendamento: {e}")

    def calculate_next_execution(self, schedule_data):
        """Calcula a próxima execução baseada no agendamento"""
        try:
            now = datetime.datetime.now()
            hora, minuto = map(int, schedule_data['horario'].split(':'))
            
            if schedule_data['tipo_agendamento'] == 'diario':
                # Próxima execução diária
                next_exec = now.replace(hour=hora, minute=minuto, second=0, microsecond=0)
                if next_exec <= now:
                    next_exec += datetime.timedelta(days=1)
                return next_exec
            
            elif schedule_data['tipo_agendamento'] == 'semanal':
                # Próxima execução semanal
                dias_semana = [int(d) for d in schedule_data['dias_semana'].split(',')]
                
                # Encontrar o próximo dia da semana
                current_weekday = now.weekday()
                # Converter para domingo = 0
                current_weekday = (current_weekday + 1) % 7
                
                next_days = []
                for dia in dias_semana:
                    days_ahead = dia - current_weekday
                    if days_ahead < 0:  # Já passou esta semana
                        days_ahead += 7
                    elif days_ahead == 0:  # Hoje
                        exec_today = now.replace(hour=hora, minute=minuto, second=0, microsecond=0)
                        if exec_today > now:
                            days_ahead = 0
                        else:
                            days_ahead = 7
                    next_days.append(days_ahead)
                
                days_ahead = min(next_days)
                next_exec = now.replace(hour=hora, minute=minuto, second=0, microsecond=0)
                next_exec += datetime.timedelta(days=days_ahead)
                return next_exec
            
            elif schedule_data['tipo_agendamento'] == 'mensal':
                # Próxima execução mensal
                dia_mes = schedule_data['dia_mes']
                
                # Tentar no mês atual
                try:
                    next_exec = now.replace(day=dia_mes, hour=hora, minute=minuto, second=0, microsecond=0)
                    if next_exec <= now:
                        # Próximo mês
                        if now.month == 12:
                            next_exec = next_exec.replace(year=now.year + 1, month=1)
                        else:
                            next_exec = next_exec.replace(month=now.month + 1)
                except ValueError:
                    # Dia não existe no mês (ex: 31 de fevereiro)
                    # Ir para o próximo mês
                    import calendar
                    if now.month == 12:
                        next_month = now.replace(year=now.year + 1, month=1, day=1)
                    else:
                        next_month = now.replace(month=now.month + 1, day=1)
                    
                    # Encontrar o último dia do mês
                    last_day = calendar.monthrange(next_month.year, next_month.month)[1]
                    dia_mes = min(dia_mes, last_day)
                    
                    next_exec = next_month.replace(day=dia_mes, hour=hora, minute=minuto, second=0, microsecond=0)
                
                return next_exec
            
            elif schedule_data['tipo_agendamento'] == 'unico':
                # Execução única
                data_especifica = schedule_data['data_especifica']
                next_exec = data_especifica.replace(hour=hora, minute=minuto, second=0, microsecond=0)
                return next_exec
            
        except Exception as e:
            print(f"Erro ao calcular próxima execução: {e}")
            return None

    def edit_schedule(self):
        """Edita o agendamento selecionado"""
        selection = self.schedule_tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um agendamento para editar.")
            return
        
        item = self.schedule_tree.item(selection[0])
        schedule_id = item['values'][0]
        
        try:
            # Buscar dados do agendamento
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM painel_agendamentos WHERE id = ?", (schedule_id,))
                row = cursor.fetchone()
                conn.close()
                
                if row:
                    # Criar objeto mock para compatibilidade
                    class MockSchedule:
                        def __init__(self, row_data):
                            self.id = row_data[0]
                            self.tarefa_id = row_data[1]
                            self.nome = row_data[2]
                            self.tipo_agendamento = row_data[3]
                            self.horario = row_data[4]
                            self.dias_semana = row_data[5]
                            self.dia_mes = row_data[6]
                            self.data_especifica = datetime.datetime.fromisoformat(row_data[7]) if row_data[7] else None
                            self.ativo = row_data[8]
                            self.retry_count = row_data[9]
                            self.max_retries = row_data[10]
                            self.notificar_sucesso = row_data[11]
                            self.notificar_erro = row_data[12]
                            self.data_criacao = row_data[13]
                            self.proxima_execucao = row_data[14]

                    schedule_data = MockSchedule(row)
                    task_id = schedule_data.tarefa_id
                else:
                    messagebox.showerror("Erro", "Agendamento não encontrado.")
                    return
            else:
                session = self.get_db_session()
                try:
                    schedule_data = session.query(Agendamento).filter(Agendamento.id == schedule_id).first()
                    if not schedule_data:
                        messagebox.showerror("Erro", "Agendamento não encontrado.")
                        return
                    task_id = schedule_data.tarefa_id
                finally:
                    session.close()
            
            # Abrir diálogo de edição
            dialog = ScheduleDialog(self.root, "Editar Agendamento", task_id, schedule_data)
            if dialog.result:
                # Calcular nova próxima execução
                proxima_execucao = self.calculate_next_execution(dialog.result)
                
                # Atualizar agendamento
                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    cursor.execute("""
                        UPDATE painel_agendamentos 
                        SET nome = ?, tipo_agendamento = ?, horario = ?, dias_semana = ?, 
                            dia_mes = ?, data_especifica = ?, max_retries = ?, 
                            notificar_sucesso = ?, notificar_erro = ?, proxima_execucao = ?
                        WHERE id = ?
                    """, (dialog.result['nome'], dialog.result['tipo_agendamento'], dialog.result['horario'],
                         dialog.result['dias_semana'], dialog.result['dia_mes'],
                         dialog.result['data_especifica'].isoformat() if dialog.result['data_especifica'] else None,
                         dialog.result['max_retries'], dialog.result['notificar_sucesso'], 
                         dialog.result['notificar_erro'], 
                         proxima_execucao.isoformat() if proxima_execucao else None, schedule_id))
                    conn.commit()
                    conn.close()
                else:
                    session = self.get_db_session()
                    try:
                        agendamento = session.query(Agendamento).filter(Agendamento.id == schedule_id).first()
                        if agendamento:
                            agendamento.nome = dialog.result['nome']
                            agendamento.tipo_agendamento = dialog.result['tipo_agendamento']
                            agendamento.horario = dialog.result['horario']
                            agendamento.dias_semana = dialog.result['dias_semana']
                            agendamento.dia_mes = dialog.result['dia_mes']
                            agendamento.data_especifica = dialog.result['data_especifica']
                            agendamento.max_retries = dialog.result['max_retries']
                            agendamento.notificar_sucesso = dialog.result['notificar_sucesso']
                            agendamento.notificar_erro = dialog.result['notificar_erro']
                            agendamento.proxima_execucao = proxima_execucao
                            session.commit()
                    finally:
                        session.close()
                
                self.refresh_schedules()
                messagebox.showinfo("Sucesso", "Agendamento atualizado com sucesso!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao editar agendamento: {e}")

    def delete_schedule(self):
        """Exclui o agendamento selecionado"""
        selection = self.schedule_tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um agendamento para excluir.")
            return
        
        item = self.schedule_tree.item(selection[0])
        schedule_id = item['values'][0]
        schedule_name = item['values'][1]
        
        if messagebox.askyesno("Confirmar Exclusão", 
                             f"Deseja realmente excluir o agendamento '{schedule_name}'?"):
            try:
                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                    import sqlite3
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    
                    # Excluir agendamento
                    cursor.execute("DELETE FROM painel_agendamentos WHERE id = ?", (schedule_id,))
                    
                    conn.commit()
                    conn.close()
                else:
                    session = self.get_db_session()
                    try:
                        agendamento = session.query(Agendamento).filter(Agendamento.id == schedule_id).first()
                        if agendamento:
                            session.delete(agendamento)
                            session.commit()
                    finally:
                        session.close()
                
                self.refresh_schedules()
                self.update_dashboard_stats()
                messagebox.showinfo("Sucesso", "Agendamento excluído com sucesso!")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao excluir agendamento: {e}")

    # ==================== SISTEMA DE AGENDAMENTO ====================

    def start_scheduler(self):
        """Inicia o sistema de agendamento"""
        self.scheduler_running = True
        self.scheduler_thread = threading.Thread(target=self.scheduler_loop, daemon=True)
        self.scheduler_thread.start()
        print("🕐 Sistema de agendamento iniciado")

    def scheduler_loop(self):
        """Loop principal do agendador"""
        while self.scheduler_running:
            try:
                self.check_scheduled_tasks()
                time.sleep(60)  # Verificar a cada minuto
            except Exception as e:
                print(f"Erro no agendador: {e}")
                time.sleep(60)

    def check_scheduled_tasks(self):
        """Verifica e executa tarefas agendadas"""
        try:
            now = datetime.datetime.now()
            
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                cursor.execute("""
                    SELECT a.id, a.tarefa_id, a.nome, a.proxima_execucao, a.max_retries, 
                           a.retry_count, a.notificar_sucesso, a.notificar_erro, t.titulo
                    FROM painel_agendamentos a
                    JOIN painel_tarefas t ON a.tarefa_id = t.id
                    WHERE a.ativo = 1 AND a.proxima_execucao <= ?
                """, (now.isoformat(),))
                
                scheduled_tasks = cursor.fetchall()
                conn.close()
                
                for task in scheduled_tasks:
                    schedule_id, task_id, schedule_name, _, max_retries, retry_count, notif_success, notif_error, task_title = task
                    
                    print(f"🕐 Executando agendamento: {schedule_name} ({task_title})")
                    
                    # Executar tarefa
                    success = self.execute_task_by_id(task_id, executado_por_agendador=True) # task_title é buscado dentro da função
                    
                    if success:
                        # Sucesso - calcular próxima execução
                        conn = sqlite3.connect(self.db_path)
                        cursor = conn.cursor()
                        
                        # Buscar dados do agendamento para recalcular
                        cursor.execute("SELECT * FROM painel_agendamentos WHERE id = ?", (schedule_id,))
                        schedule_row = cursor.fetchone()
                        
                        if schedule_row:
                            schedule_data = {
                                'tipo_agendamento': schedule_row[3],
                                'horario': schedule_row[4],
                                'dias_semana': schedule_row[5],
                                'dia_mes': schedule_row[6],
                                'data_especifica': datetime.datetime.fromisoformat(schedule_row[7]) if schedule_row[7] else None
                            }
                            
                            next_exec = self.calculate_next_execution(schedule_data)
                            
                            # Atualizar agendamento
                            cursor.execute("""
                                UPDATE painel_agendamentos 
                                SET proxima_execucao = ?, retry_count = 0
                                WHERE id = ?
                            """, (next_exec.isoformat() if next_exec else None, schedule_id))
                            
                            # Se for execução única, desativar
                            if schedule_data['tipo_agendamento'] == 'unico':
                                cursor.execute("UPDATE painel_agendamentos SET ativo = 0 WHERE id = ?", (schedule_id,))
                        
                        conn.commit()
                        conn.close()
                        
                        # Notificar sucesso se configurado
                        if notif_success:
                                                self.send_notification("sucesso", f"Agendamento executado com sucesso: {schedule_name}",
                                         f"Tarefa: {task_title}\nHorário: {now.strftime('%d/%m/%Y %H:%M')}",
                                         tarefa_id=task_id, agendamento_id=schedule_id) # IDs adicionados aqui
                    # NOVO: ALERTA POP-UP DE SUCESSO PARA AGENDAMENTO
                        self.root.after(0, lambda: messagebox.showinfo("Agendamento Executado com Sucesso",
                                        f"✅ Agendamento '{schedule_name}' para a tarefa '{task_title}' foi executado com sucesso!\n\n"
                                        f"Verifique o histórico para mais detalhes."))
                    
                    else:
                        # Falha - verificar retry
                        if retry_count < max_retries:
                            # Tentar novamente em 5 minutos
                            retry_time = now + datetime.timedelta(minutes=5)
                            
                            conn = sqlite3.connect(self.db_path)
                            cursor = conn.cursor()
                            cursor.execute("""
                                UPDATE painel_agendamentos 
                                SET proxima_execucao = ?, retry_count = retry_count + 1
                                WHERE id = ?
                            """, (retry_time.isoformat(), schedule_id))
                            conn.commit()
                            conn.close()
                            
                            print(f"⚠️ Reagendando para retry em 5 minutos: {schedule_name}")
                        else:
                            # Esgotou tentativas - calcular próxima execução normal
                            conn = sqlite3.connect(self.db_path)
                            cursor = conn.cursor()
                            
                            cursor.execute("SELECT * FROM painel_agendamentos WHERE id = ?", (schedule_id,))
                            schedule_row = cursor.fetchone()
                            
                            if schedule_row:
                                schedule_data = {
                                    'tipo_agendamento': schedule_row[3],
                                    'horario': schedule_row[4],
                                    'dias_semana': schedule_row[5],
                                    'dia_mes': schedule_row[6],
                                    'data_especifica': datetime.datetime.fromisoformat(schedule_row[7]) if schedule_row[7] else None
                                }
                                
                                next_exec = self.calculate_next_execution(schedule_data)
                                
                                cursor.execute("""
                                    UPDATE painel_agendamentos 
                                    SET proxima_execucao = ?, retry_count = 0
                                    WHERE id = ?
                                """, (next_exec.isoformat() if next_exec else None, schedule_id))
                                # Se for execução única, desativar
                                if schedule_data['tipo_agendamento'] == 'unico':
                                    cursor.execute("UPDATE painel_agendamentos SET ativo = 0 WHERE id = ?", (schedule_id,))
                            
                            conn.commit()
                            conn.close()
                            
                        # Notificar falha se configurado
                        self.send_notification("erro", f"Falha no agendamento: {schedule_name}",
                                             f"Tarefa: {task_title}\nTentativas esgotadas: {max_retries}",
                                             tarefa_id=task_id, agendamento_id=schedule_id) # IDs adicionados aqui
                        # NOVO: ALERTA POP-UP DE FALHA PARA AGENDAMENTO
                        self.root.after(0, lambda: messagebox.showerror("Falha no Agendamento",
                                            f"❌ Agendamento '{schedule_name}' para a tarefa '{task_title}' falhou!\n\n"
                                            f"Motivo: Esgotadas {max_retries} tentativas.\n"
                                            f"Verifique o histórico e alertas para mais detalhes."))
            
            else:
                # SQL Server com SQLAlchemy
                session = self.get_db_session()
                try:
                    scheduled_tasks = session.query(Agendamento).join(Tarefa).filter(
                        Agendamento.ativo == True,
                        Agendamento.proxima_execucao <= now
                    ).all()
                    
                    for agendamento in scheduled_tasks:
                        print(f"🕐 Executando agendamento: {agendamento.nome} ({agendamento.tarefa.titulo})")
                        
                        # Executar tarefa
                        success = self.execute_task_by_id(agendamento.tarefa_id, executado_por_agendador=True)
                        
                        if success:
                            # Sucesso - calcular próxima execução
                            schedule_data = {
                                'tipo_agendamento': agendamento.tipo_agendamento,
                                'horario': agendamento.horario,
                                'dias_semana': agendamento.dias_semana,
                                'dia_mes': agendamento.dia_mes,
                                'data_especifica': agendamento.data_especifica
                            }
                            
                            next_exec = self.calculate_next_execution(schedule_data)
                            agendamento.proxima_execucao = next_exec
                            agendamento.retry_count = 0
                            
                            # Se for execução única, desativar
                            if agendamento.tipo_agendamento == 'unico':
                                agendamento.ativo = False
                            
                            session.commit()
                            
                            # Notificar sucesso se configurado
                        self.send_notification("sucesso", f"Agendamento executado com sucesso: {agendamento.nome}",
                                             f"Tarefa: {agendamento.tarefa.titulo}\nHorário: {now.strftime('%d/%m/%Y %H:%M')}",
                                             tarefa_id=agendamento.tarefa_id, agendamento_id=agendamento.id) # IDs adicionados aqui
                        # NOVO: ALERTA POP-UP DE SUCESSO PARA AGENDAMENTO
                        self.root.after(0, lambda: messagebox.showinfo("Agendamento Executado com Sucesso",
                                            f"✅ Agendamento '{agendamento.nome}' para a tarefa '{agendamento.tarefa.titulo}' foi executado com sucesso!\n\n"
                                            f"Verifique o histórico para mais detalhes."))
                        
                    else:
                            # Falha - verificar retry
                            if agendamento.retry_count < agendamento.max_retries:
                                # Tentar novamente em 5 minutos
                                retry_time = now + datetime.timedelta(minutes=5)
                                agendamento.proxima_execucao = retry_time
                                agendamento.retry_count += 1
                                session.commit()
                                
                                print(f"⚠️ Reagendando para retry em 5 minutos: {agendamento.nome}")
                            else:
                                # Esgotou tentativas - calcular próxima execução normal
                                schedule_data = {
                                    'tipo_agendamento': agendamento.tipo_agendamento,
                                    'horario': agendamento.horario,
                                    'dias_semana': agendamento.dias_semana,
                                    'dia_mes': agendamento.dia_mes,
                                    'data_especifica': agendamento.data_especifica
                                }
                                
                                next_exec = self.calculate_next_execution(schedule_data)
                                agendamento.proxima_execucao = next_exec
                                agendamento.retry_count = 0
                                
                                # Se for execução única, desativar
                                if agendamento.tipo_agendamento == 'unico':
                                    agendamento.ativo = False
                                
                                session.commit()
                                
                                # Notificar falha se configurado
                            self.send_notification("erro", f"Falha no agendamento: {agendamento.nome}",
                                                 f"Tarefa: {agendamento.tarefa.titulo}\nTentativas esgotadas: {agendamento.max_retries}",
                                                 tarefa_id=agendamento.tarefa_id, agendamento_id=agendamento.id) # IDs adicionados aqui
                            # NOVO: ALERTA POP-UP DE FALHA PARA AGENDAMENTO
                            self.root.after(0, lambda: messagebox.showerror("Falha no Agendamento",
                                                f"❌ Agendamento '{agendamento.nome}' para a tarefa '{agendamento.tarefa.titulo}' falhou!\n\n"
                                                f"Motivo: Esgotadas {agendamento.max_retries} tentativas.\n"
                                                f"Verifique o histórico e alertas para mais detalhes."))
                
                finally:
                    session.close()
            
            # Atualizar interface se necessário
            if scheduled_tasks:
                self.root.after(0, self.refresh_schedules)
                self.root.after(0, self.update_dashboard_stats)
                
        except Exception as e:
            print(f"Erro ao verificar tarefas agendadas: {e}")

    # ==================== SISTEMA DE ALERTAS ====================

    def create_alert(self, tipo, titulo, mensagem, tarefa_id=None, agendamento_id=None):
        """Cria um novo alerta"""
        try:
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                cursor.execute("""
                    INSERT INTO painel_alertas (tipo, titulo, mensagem, tarefa_id, agendamento_id)
                    VALUES (?, ?, ?, ?, ?)
                """, (tipo, titulo, mensagem, tarefa_id, agendamento_id))
                
                alert_id = cursor.lastrowid
                conn.commit()
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    novo_alerta = Alerta(
                        tipo=tipo,
                        titulo=titulo,
                        mensagem=mensagem,
                        tarefa_id=tarefa_id,
                        agendamento_id=agendamento_id
                    )
                    
                    session.add(novo_alerta)
                    session.commit()
                    alert_id = novo_alerta.id
                finally:
                    session.close()
            
            # Enviar notificação se configurado
            if tipo in ['erro', 'timeout']:
                self.send_notification(tipo, titulo, mensagem)
            elif tipo == 'sucesso':
                # Só enviar notificação de sucesso se especificamente configurado
                pass
            
            print(f"🚨 Alerta criado: {titulo}")
            
        except Exception as e:
            print(f"Erro ao criar alerta: {e}")

    def resolve_alert(self):
        """Marca alerta como resolvido"""
        selection = self.alerts_tree.selection()
        if not selection:
            messagebox.showwarning("Aviso", "Selecione um alerta para marcar como resolvido.")
            return
        
        item = self.alerts_tree.item(selection[0])
        alert_id = item['values'][0]
        
        try:
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                cursor.execute("""
                    UPDATE painel_alertas 
                    SET resolvido = 1, data_resolucao = ?
                    WHERE id = ?
                """, (datetime.datetime.now().isoformat(), alert_id))
                
                conn.commit()
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    alerta = session.query(Alerta).filter(Alerta.id == alert_id).first()
                    if alerta:
                        alerta.resolvido = True
                        alerta.data_resolucao = datetime.datetime.now()
                        session.commit()
                finally:
                    session.close()
            
            self.refresh_alerts()
            self.update_dashboard_stats()
            messagebox.showinfo("Sucesso", "Alerta marcado como resolvido!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao resolver alerta: {e}")

    def clear_resolved_alerts(self):
        """Limpa alertas resolvidos"""
        if messagebox.askyesno("Confirmar", "Deseja limpar todos os alertas resolvidos?"):
            try:
                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                    import sqlite3
                    conn = sqlite3.connect(self.db_path)
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM painel_alertas WHERE resolvido = 1")
                    deleted_count = cursor.rowcount
                    conn.commit()
                    conn.close()
                else:
                    session = self.get_db_session()
                    try:
                        deleted_count = session.query(Alerta).filter(Alerta.resolvido == True).delete()
                        session.commit()
                    finally:
                        session.close()
                
                self.refresh_alerts()
                self.update_dashboard_stats()
                messagebox.showinfo("Sucesso", f"{deleted_count} alertas resolvidos foram removidos!")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao limpar alertas: {e}")

    def show_alert_details(self):
        """Mostra detalhes do alerta selecionado"""
        selection = self.alerts_tree.selection()
        if not selection:
            return
        
        item = self.alerts_tree.item(selection[0])
        alert_id = item['values'][0]
        
        try:
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT a.*, t.titulo as tarefa_titulo
                    FROM painel_alertas a
                    LEFT JOIN painel_tarefas t ON a.tarefa_id = t.id
                    WHERE a.id = ?
                """, (alert_id,))
                alert_data = cursor.fetchone()
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    alerta = session.query(Alerta).filter(Alerta.id == alert_id).first()
                    if alerta:
                        tarefa_titulo = ""
                        if alerta.tarefa_id:
                            tarefa = session.query(Tarefa).filter(Tarefa.id == alerta.tarefa_id).first()
                            if tarefa:
                                tarefa_titulo = tarefa.titulo
                        
                        alert_data = (alerta.id, alerta.tipo, alerta.titulo, alerta.mensagem,
                                    alerta.tarefa_id, alerta.agendamento_id, alerta.data_criacao,
                                    alerta.resolvido, alerta.data_resolucao, alerta.canal_notificacao,
                                    alerta.enviado, tarefa_titulo)
                    else:
                        alert_data = None
                finally:
                    session.close()
            
            if not alert_data:
                messagebox.showerror("Erro", "Alerta não encontrado.")
                return
            
            # Criar janela de detalhes
            details_window = tk.Toplevel(self.root)
            details_window.title("Detalhes do Alerta")
            details_window.geometry("600x400")
            details_window.transient(self.root)
            details_window.configure(bg='#1e3a3a')
            
            # Frame principal
            main_frame = ttk.Frame(details_window, padding="20")
            main_frame.pack(fill='both', expand=True)
            
            # Título
            ttk.Label(main_frame, text="Detalhes do Alerta", 
                     font=('Segoe UI', 14, 'bold')).pack(pady=(0, 10))
            
            # Informações
            info_frame = ttk.LabelFrame(main_frame, text="Informações", padding=10)
            info_frame.pack(fill='x', pady=5)
            
            ttk.Label(info_frame, text=f"ID: {alert_data[0]}").pack(anchor='w')
            ttk.Label(info_frame, text=f"Tipo: {alert_data[1]}").pack(anchor='w')
            ttk.Label(info_frame, text=f"Título: {alert_data[2]}").pack(anchor='w')
            ttk.Label(info_frame, text=f"Tarefa: {alert_data[11] or 'N/A'}").pack(anchor='w')
            
            # Data de criação
            data_criacao = "N/A"
            if alert_data[6]:
                try:
                    if isinstance(alert_data[6], str):
                        dt = datetime.datetime.fromisoformat(alert_data[6])
                    else:
                        dt = alert_data[6]
                    data_criacao = dt.strftime("%d/%m/%Y %H:%M:%S")
                except:
                    data_criacao = str(alert_data[6])
            
            ttk.Label(info_frame, text=f"Data/Hora: {data_criacao}").pack(anchor='w')
            ttk.Label(info_frame, text=f"Status: {'✅ Resolvido' if alert_data[7] else '⚠️ Pendente'}").pack(anchor='w')
            
            # Mensagem
            msg_frame = ttk.LabelFrame(main_frame, text="Mensagem", padding=10)
            msg_frame.pack(fill='both', expand=True, pady=5)
            
            msg_text = tk.Text(msg_frame, height=10, wrap='word', state='disabled',
                              bg='#2d5555', fg='#e8f4f4', insertbackground='#e8f4f4')
            msg_text.pack(fill='both', expand=True)
            
            msg_text.config(state='normal')
            msg_text.insert(1.0, alert_data[3] or "Sem mensagem")
            msg_text.config(state='disabled')
            
            # Botão fechar
            ttk.Button(main_frame, text="Fechar", command=details_window.destroy).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar detalhes: {e}")

    # ==================== SISTEMA DE NOTIFICAÇÕES ====================

    def load_notification_configs(self):
        """Carrega configurações de notificação"""
        try:
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT canal, ativo, configuracao FROM painel_config_notificacao")
                configs = cursor.fetchall()
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    configs = [(c.canal, c.ativo, c.configuracao) for c in session.query(ConfiguracaoNotificacao).all()]
                finally:
                    session.close()
            
            # Aplicar configurações na interface
            for canal, ativo, config_json in configs:
                if canal == "email":
                    self.email_active_var.set(ativo)
                    if config_json:
                        try:
                            config = json.loads(config_json)
                            self.smtp_server_entry.delete(0, tk.END)
                            self.smtp_server_entry.insert(0, config.get('smtp_server', 'smtp.gmail.com'))
                            self.smtp_port_entry.delete(0, tk.END)
                            self.smtp_port_entry.insert(0, config.get('smtp_port', '587'))
                            self.email_entry.delete(0, tk.END)
                            self.email_entry.insert(0, config.get('email', ''))
                            self.email_password_entry.delete(0, tk.END)
                            self.email_password_entry.insert(0, config.get('password', ''))
                            self.email_to_entry.delete(0, tk.END)
                            self.email_to_entry.insert(0, config.get('to_email', ''))
                        except:
                            pass
            
        except Exception as e:
            print(f"Erro ao carregar configurações de notificação: {e}")

    def save_notification_config(self):
        """Salva configurações de notificação"""
        try:
            # Configuração de email
            email_config = {
                'smtp_server': self.smtp_server_entry.get(),
                'smtp_port': self.smtp_port_entry.get(),
                'email': self.email_entry.get(),
                'password': self.email_password_entry.get(),
                'to_email': self.email_to_entry.get()
            }
            
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                # Deletar configuração existente
                cursor.execute("DELETE FROM painel_config_notificacao WHERE canal = 'email'")
                
                # Inserir nova configuração
                cursor.execute("""
                    INSERT INTO painel_config_notificacao (canal, ativo, configuracao)
                    VALUES (?, ?, ?)
                """, ('email', self.email_active_var.get(), json.dumps(email_config)))
                
                conn.commit()
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    # Deletar configuração existente
                    session.query(ConfiguracaoNotificacao).filter(
                        ConfiguracaoNotificacao.canal == 'email'
                    ).delete()
                    
                    # Inserir nova configuração
                    nova_config = ConfiguracaoNotificacao(
                        canal='email',
                        ativo=self.email_active_var.get(),
                        configuracao=json.dumps(email_config)
                    )
                    
                    session.add(nova_config)
                    session.commit()
                finally:
                    session.close()
            
            messagebox.showinfo("Sucesso", "Configurações de notificação salvas com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar configurações: {e}")

    def test_email(self):
        """Testa configuração de email"""
        try:
            smtp_server = self.smtp_server_entry.get()
            smtp_port = int(self.smtp_port_entry.get())
            email = self.email_entry.get()
            password = self.email_password_entry.get()
            to_email = self.email_to_entry.get()
            
            if not all([smtp_server, smtp_port, email, password, to_email]):
                messagebox.showwarning("Aviso", "Preencha todos os campos de email.")
                return
            
            # Criar mensagem de teste
            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = to_email
            msg['Subject'] = "Teste de Notificação - Painel de Controle"
            
            body = """
            Este é um teste de notificação do Painel de Controle.
            
            Se você recebeu este email, as configurações estão corretas!
            
            Data/Hora: {}
            Sistema: Painel de Automação
            """.format(datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Enviar email
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(email, password)
            server.send_message(msg)
            server.quit()
            
            messagebox.showinfo("Sucesso", "Email de teste enviado com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar email de teste: {e}")

    def send_notification(self, tipo, titulo, mensagem):
        """Envia notificação pelos canais configurados"""
        try:
            # Verificar se email está ativo
            if self.email_active_var.get():
                self.send_email_notification(tipo, titulo, mensagem)
            
            # Outros canais podem ser implementados aqui
            
        except Exception as e:
            print(f"Erro ao enviar notificação: {e}")

    def send_email_notification(self, tipo, titulo, mensagem):
        """Envia notificação por email"""
        try:
            # Buscar configurações
            if self.using_sqlite and not SQLSERVER_AVAILABLE:
                import sqlite3
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                cursor.execute("SELECT configuracao FROM painel_config_notificacao WHERE canal = 'email' AND ativo = 1")
                result = cursor.fetchone()
                conn.close()
            else:
                session = self.get_db_session()
                try:
                    config_obj = session.query(ConfiguracaoNotificacao).filter(
                        ConfiguracaoNotificacao.canal == 'email',
                        ConfiguracaoNotificacao.ativo == True
                    ).first()
                    result = (config_obj.configuracao,) if config_obj else None
                finally:
                    session.close()
            
            if not result:
                return
            
            config = json.loads(result[0])
            
            # Emoji baseado no tipo
            emoji = "❌" if tipo == "erro" else "⏰" if tipo == "timeout" else "✅"
            
            # Criar mensagem
            msg = MIMEMultipart()
            msg['From'] = config['email']
            msg['To'] = config['to_email']
            msg['Subject'] = f"{emoji} {titulo} - Painel de Controle"
            
            body = f"""
            {titulo}
            
            {mensagem}
            
            Data/Hora: {datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
            Sistema: Painel de Controle
            Usuário: Usuário # 
            
            ---
            Esta é uma notificação automática do sistema de automação.
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Enviar email
            server = smtplib.SMTP(config['smtp_server'], int(config['smtp_port']))
            server.starttls()
            server.login(config['email'], config['password'])
            server.send_message(msg)
            server.quit()
            
            print(f"📧 Email enviado: {titulo}")
            
        except Exception as e:
            print(f"Erro ao enviar email: {e}")

    def test_notification(self):
        """Testa sistema de notificação"""
        self.create_alert("teste", "Teste de Notificação", 
                         "Este é um teste do sistema de alertas e notificações.")
        messagebox.showinfo("Teste", "Alerta de teste criado! Verifique a aba de Alertas.")

    # ==================== FUNÇÕES DE BACKUP E DADOS ====================

    def export_data(self):
        """Exporta dados para arquivo JSON"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                title="Exportar Dados"
            )
            
            if filename:
                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                    import sqlite3
                    conn = sqlite3.connect(self.db_path)
                    
                    # Exportar todas as tabelas
                    tables = ['painel_tarefas', 'painel_execucoes', 'painel_agendamentos', 
                             'painel_alertas', 'painel_config_notificacao']
                    
                    export_data = {}
                    for table in tables:
                        df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
                        export_data[table] = df.to_dict('records')
                    
                    conn.close()
                else:
                    with self.engine.connect() as conn:
                        # Exportar todas as tabelas
                        tables = ['painel_tarefas', 'painel_execucoes', 'painel_agendamentos', 
                                 'painel_alertas', 'painel_config_notificacao']
                        
                        export_data = {}
                        for table in tables:
                            df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
                            export_data[table] = df.to_dict('records')
                
                # Criar estrutura de exportação
                final_export = {
                    'export_date': datetime.datetime.now().isoformat(),
                    'version': '3.0',
                    'usuario': 'Usuário', # 
                    'database_type': 'SQLite' if self.using_sqlite else 'SQL Server',
                    'data': export_data
                }
                
                # Salvar arquivo
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(final_export, f, ensure_ascii=False, indent=2, default=str)
                
                messagebox.showinfo("Sucesso", f"Dados exportados para {filename}")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar dados: {e}")

    def import_data(self):
        """Importa dados de arquivo JSON"""
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                title="Importar Dados"
            )
            
            if filename:
                with open(filename, 'r', encoding='utf-8') as f:
                    import_data = json.load(f)
                
                imported_count = 0
                
                # Importar dados baseado na estrutura
                if 'data' in import_data:
                    # Nova estrutura (v3.0)
                    data = import_data['data']
                    
                    # Importar tarefas
                    if 'painel_tarefas' in data:
                        for task in data['painel_tarefas']:
                            try:
                                if self.using_sqlite and not SQLSERVER_AVAILABLE:
                                    import sqlite3
                                    conn = sqlite3.connect(self.db_path)
                                    cursor = conn.cursor()
                                    # INSERT ATUALIZADO PARA NOVOS CAMPOS
                                    cursor.execute("""
                                        INSERT INTO painel_tarefas (titulo, descricao, prioridade, arquivo_BAT, tipo_execucao,
                                                                   cd_source_path, cd_destination_node, cd_destination_path, cd_process_name) 
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                    """, (task.get('titulo'), task.get('descricao'), 
                                         task.get('prioridade'), task.get('arquivo_BAT'), 
                                         task.get('tipo_execucao', 'BAT'), # Valor padrão para compatibilidade
                                         task.get('cd_source_path'), task.get('cd_destination_node'),
                                         task.get('cd_destination_path'), task.get('cd_process_name')))
                                    conn.commit()
                                    conn.close()
                                else:
                                    session = self.get_db_session()
                                    try:
                                        nova_tarefa = Tarefa(
                                            titulo=task.get('titulo'),
                                            descricao=task.get('descricao'),
                                            prioridade=task.get('prioridade'),
                                            arquivo_BAT=task.get('arquivo_BAT'),
                                            tipo_execucao=task.get('tipo_execucao', 'BAT'),
                                            cd_source_path=task.get('cd_source_path'),
                                            cd_destination_node=task.get('cd_destination_node'),
                                            cd_destination_path=task.get('cd_destination_path'),
                                            cd_process_name=task.get('cd_process_name')
                                        )
                                        session.add(nova_tarefa)
                                        session.commit()
                                    finally:
                                        session.close()
                                
                                imported_count += 1
                            except Exception: # Ignora tarefas que não podem ser importadas (ex: chaves duplicadas)
                                continue
                
                # Atualizar interface
                self.load_data()
                messagebox.showinfo("Sucesso", 
                    f"Dados importados com sucesso!\n"
                    f"Itens importados: {imported_count}")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar dados: {e}")

    def backup_database(self):
        """Faz backup do banco de dados"""
        try:
            backup_filename = f"backup_painel_controle_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            if self.using_sqlite:
                backup_filename += ".db"
                backup_path = filedialog.asksaveasfilename(
                    initialname=backup_filename,
                    defaultextension=".db",
                    filetypes=[("Database files", "*.db"), ("All files", "*.*日子")],
                    title="Salvar Backup SQLite"
                )
                
                if backup_path:
                    import shutil
                    shutil.copy2(self.db_path, backup_path)
                    messagebox.showinfo("Sucesso", f"Backup SQLite criado: {backup_path}")
            else:
                backup_filename += ".json"
                backup_path = filedialog.asksaveasfilename(
                    initialname=backup_filename,
                    defaultextension=".json",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*日子")],
                    title="Salvar Backup SQL Server"
                )
                
                if backup_path:
                    # Para SQL Server, fazer backup via export
                    self.export_data_to_file(backup_path)
                    messagebox.showinfo("Sucesso", f"Backup SQL Server criado: {backup_path}")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao fazer backup: {e}")

    def export_data_to_file(self, filename):
        """Exporta dados para arquivo específico"""
        with self.engine.connect() as conn:
            # Exportar todas as tabelas
            tables = ['painel_tarefas', 'painel_execucoes', 'painel_agendamentos', 
                     'painel_alertas', 'painel_config_notificacao']
            
            export_data = {}
            for table in tables:
                df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
                export_data[table] = df.to_dict('records')
        
        # Criar estrutura de backup
        backup_data = {
            'backup_date': datetime.datetime.now().isoformat(),
            'version': '3.0',
            'usuario': 'Usuário', # 
            'database_type': 'SQL Server',
            'server': DatabaseConfig.DB_SERVER,
            'database': DatabaseConfig.DB_NAME,
            'data': export_data
        }
        
        # Salvar arquivo
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(backup_data, f, ensure_ascii=False, indent=2, default=str)

    def restore_database(self):
        """Restaura backup do banco de dados"""
        if messagebox.askyesno("Confirmar", "Isso substituirá todos os dados atuais. Continuar?"):
            try:
                if self.using_sqlite:
                    backup_path = filedialog.askopenfilename(
                        filetypes=[("Database files", "*.db"), ("All files", "*.*日子")],
                        title="Selecionar Backup SQLite"
                    )
                    
                    if backup_path:
                        # Parar agendador
                        self.scheduler_running = False
                        time.sleep(2)
                        
                        # Restaurar arquivo
                        import shutil
                        shutil.copy2(self.db_path, backup_path)
                        
                        # Recarregar dados
                        self.load_data()
                        
                        # Reiniciar agendador
                        self.start_scheduler()
                        
                        messagebox.showinfo("Sucesso", "Backup SQLite restaurado com sucesso!")
                else:
                    backup_path = filedialog.askopenfilename(
                        filetypes=[("JSON files", "*.json"), ("All files", "*.*日子")],
                        title="Selecionar Backup SQL Server"
                    )
                    
                    if backup_path:
                        # Importar dados do backup
                        self.import_data_from_file(backup_path)
                        messagebox.showinfo("Sucesso", "Backup SQL Server restaurado com sucesso!")
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao restaurar backup: {e}")

    def import_data_from_file(self, filename):
        """Importa dados de arquivo específico"""
        with open(filename, 'r', encoding='utf-8') as f:
            backup_data = json.load(f)
        
        # Limpar dados existentes
        session = self.get_db_session()
        try:
            # Excluir dados em ordem inversa de dependência
            session.query(Alerta).delete()
            session.query(Agendamento).delete()
            session.query(Execucao).delete()
            # session.query(SistemaMonitor).delete() # Não incluído no export/import atual
            session.query(ConfiguracaoNotificacao).delete()
            session.query(Tarefa).delete() # Tarefas por último
            session.commit()
            
            # Importar dados
            if 'data' in backup_data:
                data = backup_data['data']
                
                # Importar tarefas
                if 'painel_tarefas' in data:
                    for task in data['painel_tarefas']:
                        nova_tarefa = Tarefa(
                            id=task.get('id'), # Tentar manter IDs para relacionamentos, se houver
                            titulo=task.get('titulo'),
                            descricao=task.get('descricao'),
                            status=task.get('status'),
                            prioridade=task.get('prioridade'),
                            arquivo_BAT=task.get('arquivo_BAT'),
                            data_criacao=datetime.datetime.fromisoformat(task.get('data_criacao')) if task.get('data_criacao') else None,
                            data_ultima_execucao=datetime.datetime.fromisoformat(task.get('data_ultima_execucao')) if task.get('data_ultima_execucao') else None,
                            total_execucoes=task.get('total_execucoes'),
                            tipo_execucao=task.get('tipo_execucao', 'BAT'),
                            cd_source_path=task.get('cd_source_path'),
                            cd_destination_node=task.get('cd_destination_node'),
                            cd_destination_path=task.get('cd_destination_path'),
                            cd_process_name=task.get('cd_process_name')
                        )
                        session.add(nova_tarefa)
                        # Comitar em lote pode ser mais eficiente, mas para backup pequeno, linha a linha ok
                        session.flush() # Para gerar IDs se autoincrement
                
                # Importar configurações de notificação
                if 'painel_config_notificacao' in data:
                    for config in data['painel_config_notificacao']:
                        nova_config = ConfiguracaoNotificacao(
                            id=config.get('id'),
                            canal=config.get('canal'),
                            ativo=config.get('ativo'),
                            configuracao=config.get('configuracao'),
                            data_atualizacao=datetime.datetime.fromisoformat(config.get('data_atualizacao')) if config.get('data_atualizacao') else None
                        )
                        session.add(nova_config)
                        session.flush()

                # Importar agendamentos, execuções e alertas (assumindo IDs de tarefa e agendamento correspondem)
                # IMPORTANTE: A ordem importa devido às chaves estrangeiras. Tarefas primeiro, depois dependentes.
                if 'painel_agendamentos' in data:
                    for sched in data['painel_agendamentos']:
                        novo_agendamento = Agendamento(
                            id=sched.get('id'),
                            tarefa_id=sched.get('tarefa_id'),
                            nome=sched.get('nome'),
                            tipo_agendamento=sched.get('tipo_agendamento'),
                            horario=sched.get('horario'),
                            dias_semana=sched.get('dias_semana'),
                            dia_mes=sched.get('dia_mes'),
                            data_especifica=datetime.datetime.fromisoformat(sched.get('data_especifica')) if sched.get('data_especifica') else None,
                            ativo=sched.get('ativo'),
                            retry_count=sched.get('retry_count'),
                            max_retries=sched.get('max_retries'),
                            notificar_sucesso=sched.get('notificar_sucesso'),
                            notificar_erro=sched.get('notificar_erro'),
                            data_criacao=datetime.datetime.fromisoformat(sched.get('data_criacao')) if sched.get('data_criacao') else None,
                            proxima_execucao=datetime.datetime.fromisoformat(sched.get('proxima_execucao')) if sched.get('proxima_execucao') else None
                        )
                        session.add(novo_agendamento)
                        session.flush()

                if 'painel_execucoes' in data:
                    for exec_data in data['painel_execucoes']:
                        nova_execucao = Execucao(
                            id=exec_data.get('id'),
                            tarefa_id=exec_data.get('tarefa_id'),
                            data_execucao=datetime.datetime.fromisoformat(exec_data.get('data_execucao')) if exec_data.get('data_execucao') else None,
                            status=exec_data.get('status'),
                            codigo_retorno=exec_data.get('codigo_retorno'),
                            duracao_segundos=exec_data.get('duracao_segundos'),
                            log_output=exec_data.get('log_output'),
                            executado_por_agendador=exec_data.get('executado_por_agendador')
                        )
                        session.add(nova_execucao)
                        session.flush()

                if 'painel_alertas' in data:
                    for alert_data in data['painel_alertas']:
                        novo_alerta = Alerta(
                            id=alert_data.get('id'),
                            tipo=alert_data.get('tipo'),
                            titulo=alert_data.get('titulo'),
                            mensagem=alert_data.get('mensagem'),
                            tarefa_id=alert_data.get('tarefa_id'),
                            agendamento_id=alert_data.get('agendamento_id'),
                            data_criacao=datetime.datetime.fromisoformat(alert_data.get('data_criacao')) if alert_data.get('data_criacao') else None,
                            resolvido=alert_data.get('resolvido'),
                            data_resolucao=datetime.datetime.fromisoformat(alert_data.get('data_resolucao')) if alert_data.get('data_resolucao') else None,
                            canal_notificacao=alert_data.get('canal_notificacao'),
                            enviado=alert_data.get('enviado')
                        )
                        session.add(novo_alerta)
                        session.flush()
            
            session.commit()
            
        finally:
            session.close()
        
        # Recarregar dados
        self.load_data()

    def clear_data(self):
        """Limpa todos os dados"""
        if messagebox.askyesno("Confirmar", "Isso excluirá TODOS os dados. Continuar?"):
            if messagebox.askyesno("Confirmar Novamente", "Tem certeza absoluta? Esta ação não pode ser desfeita!"):
                try:
                    if self.using_sqlite and not SQLSERVER_AVAILABLE:
                        import sqlite3
                        conn = sqlite3.connect(self.db_path)
                        cursor = conn.cursor()
                        
                        # Limpar todas as tabelas em ordem inversa de dependência
                        tables = ['painel_alertas', 'painel_agendamentos', 'painel_execucoes', 
                                 'painel_sistema_monitor', 'painel_config_notificacao', 'painel_tarefas']
                        
                        for table in tables:
                            cursor.execute(f"DELETE FROM {table}")
                        
                        conn.commit()
                        conn.close()
                    else:
                        session = self.get_db_session()
                        try:
                            # Excluir dados em ordem inversa de dependência
                            session.query(Alerta).delete()
                            session.query(Agendamento).delete()
                            session.query(Execucao).delete()
                            session.query(SistemaMonitor).delete()
                            session.query(ConfiguracaoNotificacao).delete()
                            session.query(Tarefa).delete() # Tarefas por último
                            session.commit()
                        finally:
                            session.close()
                    
                    self.load_data()
                    messagebox.showinfo("Sucesso", "Todos os dados foram limpos!")
                    
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao limpar dados: {e}")

    def show_about(self):
        """Mostra informações sobre o sistema"""
        db_info = f"SQL Server ({DatabaseConfig.DB_SERVER})" if not self.using_sqlite else "SQLite (Local)"
        
        about_text = f"""🤖 Painel de Controle
Rotinas Automatizadas com Agendador
Versão: 3.0.0 - Controle Central Edition
Desenvolvido por: Usuário # 

🎯 Funcionalidades:
• Gerenciamento de Tarefas Automatizadas (BAT e Connect:Direct via API)
• Agendador Inteligente (Diário, Semanal, Mensal, Único)
• Sistema de Alertas e Notificações
• Notificações por Email
• Retry Automático em Falhas
• Dashboard com Estatísticas Avançadas
• Histórico Detalhado de Execuções
• Backup e Restauração Completa
• Exportar/Importar Dados
• Conexão SQL Server {'✅' if not self.using_sqlite else '❌ (usando SQLite)'}
• Integração Connect:Direct (via API IBM Control Center)

🛠️ Tecnologias:
• Python {sys.version.split()[0]}
• Tkinter (Interface)
• {db_info} (Banco de Dados)
• Matplotlib (Gráficos)
• Pandas (Manipulação de Dados)
• Schedule (Agendamento)
• SMTP (Email)
• Requests (Chamadas API HTTP/REST)
{'• SQLAlchemy (ORM)' if SQLSERVER_AVAILABLE else ''}
{'• pyodbc (Driver SQL Server)' if SQLSERVER_AVAILABLE else ''}

💾 Configuração do Banco:
• Servidor: {DatabaseConfig.DB_SERVER if not self.using_sqlite else 'Local'}
• Banco: {DatabaseConfig.DB_NAME if not self.using_sqlite else 'painel_dados.db'}
• Usuário: {DatabaseConfig.DB_USER if not self.using_sqlite else 'N/A'}

👨‍💻 Usuário: Usuário # 
📧 Suporte a aplicação: seu.email@exemplo.com # 
🌐 Website: seu.website.com # 

© 2025. Todos os direitos reservados. # 
"""
        
        # Criar janela de sobre
        about_window = tk.Toplevel(self.root)
        about_window.title("Sobre o Painel de Controle")
        about_window.geometry("600x600")
        about_window.transient(self.root)
        about_window.grab_set()
        about_window.configure(bg='#1e3a3a')
        
        # Centralizar janela
        about_window.update_idletasks()
        x = (about_window.winfo_screenwidth() // 2) - (600 // 2)
        y = (about_window.winfo_screenheight() // 2) - (600 // 2)
        about_window.geometry(f"600x600+{x}+{y}")
        
        # Frame principal
        main_frame = ttk.Frame(about_window, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        # Texto sobre
        text_widget = tk.Text(main_frame, wrap='word', state='disabled', height=30,
                             bg='#2d5555', fg='#e8f4f4', insertbackground='#e8f4f4')
        text_widget.pack(fill='both', expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        
        # Inserir texto
        text_widget.config(state='normal')
        text_widget.insert(1.0, about_text)
        text_widget.config(state='disabled')
        
        # Botão fechar
        ttk.Button(main_frame, text="Fechar", command=about_window.destroy).pack(pady=10)

    def on_closing(self):
        """Manipula o fechamento da aplicação"""
        if messagebox.askokcancel("Sair", "Deseja realmente sair do Painel de Controle?"):
            # Parar agendador
            self.scheduler_running = False
            
            # Fechar aplicação
            self.root.destroy()

# ==================== FUNÇÃO PRINCIPAL ====================

def main():
    """Função principal da aplicação"""
    print("🚀 Iniciando Painel de Controle..")
    print("=" * 60)
    
    # Verificar dependências obrigatórias
    missing_deps = []
    
    try:
        import matplotlib
        import pandas
        import schedule
        print("✅ matplotlib: Instalado")
        print("✅ pandas: Instalado")
        print("✅ schedule: Instalado")
    except ImportError as e:
        missing_deps.append(f"matplotlib/pandas/schedule: {e}")
    
    # Verificar dependências SQL Server
    if SQLSERVER_AVAILABLE:
        print("✅ pyodbc: Instalado")
        print("✅ sqlalchemy: Instalado")
    else:
        print("⚠️ pyodbc/sqlalchemy: Não instalado - usando SQLite")
    
    # Verificar dependências opcionais
    if PSUTIL_AVAILABLE:
        print("✅ psutil: Instalado")
    else:
        print("⚠️ psutil: Não instalado")
    
    if REQUESTS_AVAILABLE:
        print("✅ requests: Instalado")
    else:
        print("⚠️ requests: Não instalado")
    
    if missing_deps:
        print("\n❌ Dependências obrigatórias não encontradas:")
        for dep in missing_deps:
            print(f"   - {dep}")
        print("\nInstale as dependências com:")
        print("pip install matplotlib pandas schedule")
        if not SQLSERVER_AVAILABLE: # A mensagem já está no corpo, mas reiterar não faz mal
            print("pip install pyodbc sqlalchemy")
        return
    
    print("\n" + "=" * 60)
    
    # Testar conexão SQL Server se disponível
    if SQLSERVER_AVAILABLE:
        print("🔍 Testando conexão SQL Server...")
        if DatabaseConfig.test_connection():
            print("✅ Conexão SQL Server: OK")
        else:
            print("⚠️ Conexão SQL Server: Falhou - usando SQLite como fallback")
    
    print("\n" + "=" * 60)
    
    # Criar e executar aplicação
    root = tk.Tk()
    
    try:
        app = PainelDesktop(root)
        
        # Configurar evento de fechamento
        root.protocol("WM_DELETE_WINDOW", app.on_closing)
        
        print("✅ Painel de Controle iniciado com sucesso!")
        print("📊 Interface carregada")
        print("💾 Banco de dados configurado")
        print("🤖 Sistema de automação ativo")
        print("🕐 Agendador de tarefas ativo")
        print("🚨 Sistema de alertas ativo")
        print("📧 Sistema de notificações configurado")
        
        if not app.using_sqlite:
            print(f"🗄️ Conectado ao SQL Server: {DatabaseConfig.DB_SERVER}")
        else:
            print("🗄️ Usando SQLite local")
        
        print("\n🎉 Sistema pronto para uso!")
        print("=" * 60)
        
        # Executar loop principal
        root.mainloop()
        
    except KeyboardInterrupt:
        print("\n⚠️ Aplicação interrompida pelo usuário")
    except Exception as e:
        print(f"❌ Erro fatal na aplicação: {e}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Erro Fatal", f"Erro na aplicação: {str(e)}")
    finally:
        print("👋 Painel de Controle encerrado")

# ==================== PONTO DE ENTRADA ====================

if __name__ == "__main__":
    main()
