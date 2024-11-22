import os
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Optional

class Config:
    # Configurações de diretórios e arquivos
    DIRETORIO_BASE = r"CAMINHO\PARA\IMPORTACAO_CADASTRAL"
    CONTROL_FILE = r'CAMINHO\PARA\Controle_Alteracao.xlsx'
    DIRETORIO_SAIDA = r"CAMINHO\PARA\SAIDA"
    DIRETORIO_LOG = r"CAMINHO\PARA\LOGS"

    @staticmethod
    def obter_arquivo_origem() -> str:
        # Gera o caminho para o arquivo de entrada com base na data atual
        data_hoje = datetime.now().strftime('%d%m%Y')
        return os.path.join(Config.DIRETORIO_BASE, f"AMC_{data_hoje}.xlsm")

class Logger:
    @staticmethod
    def log(mensagem: str, nivel: str = 'INFO'):
        # Registra mensagens de log em um arquivo de log diário
        try:
            os.makedirs(Config.DIRETORIO_LOG, exist_ok=True)
            data_atual = datetime.now().strftime('%d-%m-%Y')
            hora_atual = datetime.now().strftime('%H:%M:%S')
            nome_arquivo_log = os.path.join(
                Config.DIRETORIO_LOG, f"mancad_log_{data_atual}.txt")
            with open(nome_arquivo_log, 'a', encoding='utf-8') as arquivo_log:
                arquivo_log.write(
                    f"{data_atual} {hora_atual} - [{nivel}] {mensagem}\n")
            print(f"{hora_atual} - [{nivel}] {mensagem}")
        except Exception as e:
            print(f"Erro ao registrar log: {e}")
            print(f"{hora_atual} - [ERRO] {mensagem}")

class FormatadorMancad:
    @staticmethod
    def formatar_registro(registro: Dict[str, Any]) -> str:
        # Formata registros com base no tipo especificado
        tipo = registro.get('TIPO', '')
        try:
            if tipo == 4:
                return FormatadorMancad._formatar_registro_04(registro)
            elif tipo == 5:
                return FormatadorMancad._formatar_registro_05(registro)
            elif tipo == 13:
                return FormatadorMancad._formatar_registro_13(registro)
            elif tipo == 28:
                return FormatadorMancad._formatar_registro_28(registro)
            elif tipo == 29:
                return FormatadorMancad._formatar_registro_29(registro)
            elif tipo == 45:
                return FormatadorMancad._formatar_registro_45(registro)
            elif tipo == 46:
                return FormatadorMancad._formatar_registro_46(registro)
            else:
                Logger.log(f"Tipo de registro {tipo} não suportado.", 'ALERTA')
                return ""
        except Exception as e:
            Logger.log(f"Erro ao formatar registro tipo {tipo}: {e}", 'ERRO')
            return ""

    # Métodos estáticos para formatar tipos específicos de registros
    @staticmethod
    def _formatar_registro_04(registro: Dict[str, Any]) -> str:
        numero_conta = str(registro.get('NUMERO CONTA', '')).zfill(16)
        logo = str(registro.get('LOGO', '')).zfill(3)
        limite = str(registro.get('VALOR_LIMITE', '0')).zfill(8)
        registro_mancad = f"041{numero_conta}{logo}907{limite}"
        return registro_mancad.ljust(378) + "00"

    @staticmethod
    def _formatar_registro_05(registro: Dict[str, Any]) -> str:
        mot = registro.get('MOT', '')
        if pd.isna(mot) or mot == '':
            return ""
        numero_conta = str(registro.get('NUMERO CONTA', '')).zfill(16)
        logo = str(registro.get('LOGO', '')).zfill(3)
        mot = str(mot).zfill(2)
        registro_mancad = f"051{numero_conta}{logo}{mot}"
        return registro_mancad.ljust(378) + "00"

    @staticmethod
    def _formatar_registro_13(registro: Dict[str, Any]) -> str:
        numero_cartao = str(registro.get('NUMERO CARTAO', '')).zfill(16)
        logo = str(registro.get('LOGO', '')).zfill(3)
        cod_bloqueio = str(registro.get('BLOCK CODE', '')).ljust(1)
        registro_mancad = f"131{numero_cartao}{logo}{cod_bloqueio}"
        return registro_mancad.ljust(378) + "00"

    @staticmethod
    def _formatar_registro_28(registro: Dict[str, Any]) -> str:
        numero_cartao = str(registro.get('NUMERO CARTAO', '')).zfill(16)
        logo = str(registro.get('LOGO', '')).zfill(3)
        indicador = str(registro.get('INDICADOR', '')).ljust(1)
        registro_mancad = f"281{numero_cartao}{logo}{indicador}"
        return registro_mancad.ljust(378) + "00"

    @staticmethod
    def _formatar_registro_29(registro: Dict[str, Any]) -> str:
        numero_conta = str(registro.get('NUMERO CONTA', '')).zfill(16)
        logo = str(registro.get('LOGO', '')).zfill(3)
        rmc = str(registro.get('VALOR_RMC', '0')).zfill(17)
        registro_mancad = f"291{numero_conta}{logo}{rmc}"
        return registro_mancad.ljust(378) + "00"

    @staticmethod
    def _formatar_registro_45(registro: Dict[str, Any]) -> str:
        numero_conta = str(registro.get('NUMERO CONTA', '')).zfill(16)
        logo = str(registro.get('LOGO', '')).zfill(3)
        block_code = str(registro.get('BLOCK CODE', '')).ljust(1)
        registro_mancad = f"451{numero_conta}{logo}{block_code}"
        return registro_mancad.ljust(378) + "00"

    @staticmethod
    def _formatar_registro_46(registro: Dict[str, Any]) -> str:
        numero_conta = str(registro.get('NUMERO CONTA', '')).zfill(16)
        logo = str(registro.get('LOGO', '')).zfill(3)
        matricula = str(registro.get('MATRICULA', '')).zfill(12)
        contrato = str(registro.get('CONTRATO', '')).ljust(20)
        registro_mancad = f"461{numero_conta}{logo}{matricula}{contrato}"
        return registro_mancad.ljust(378) + "00"

class ProcessadorArquivoExcel:
    def __init__(self, caminho_arquivo: str = Config.obter_arquivo_origem()):
        self.caminho_arquivo = caminho_arquivo

    def ler_arquivo(self) -> List[Dict[str, Any]]:
        # Lê o arquivo Excel e retorna uma lista de registros
        try:
            if not os.path.exists(self.caminho_arquivo):
                Logger.log(
                    f"ERRO: Arquivo não encontrado - {self.caminho_arquivo}", 'ERRO')
                return []

            sheets = pd.ExcelFile(self.caminho_arquivo).sheet_names
            registros = []

            for sheet in sheets:
                df = pd.read_excel(self.caminho_arquivo, sheet_name=sheet)
                registros.extend(df.to_dict('records'))

            Logger.log(f"Total de registros encontrados: {len(registros)}")
            return registros

        except Exception as e:
            Logger.log(f"Erro ao ler arquivo Excel: {e}", 'ERRO')
            return []

class ProcessadorMancad:
    def __init__(self, data: datetime = datetime.now()):
        self.data = data

    def obter_e_incrementar_valor_coluna_c(self) -> str:
        # Obtém e incrementa o valor na coluna C do arquivo de controle
        try:
            df_controle = pd.read_excel(Config.CONTROL_FILE)

            if df_controle.empty:
                valor_coluna_c = 3700  # Valor inicial se a planilha estiver vazia
            else:
                # Incrementa o último valor
                valor_coluna_c = df_controle.iloc[-1, 2] + 1

            # Atualiza o DataFrame com o novo valor na coluna C
            # Atualiza a última linha
            df_controle.loc[df_controle.index[-1],
                            'Sequencial'] = valor_coluna_c
            df_controle.to_excel(Config.CONTROL_FILE, index=False)

            # Preenche com zeros à esquerda para garantir 4 dígitos
            return str(valor_coluna_c).zfill(4)
        except Exception as e:
            Logger.log(f"Erro ao atualizar valor da coluna C: {e}", 'ERRO')
            return "0000"

    def criar_header(self, numero_lote: int) -> str:
        # Cria o cabeçalho do arquivo de saída
        header = (
            f"00"
            f"{self.data.strftime('%d%m%y')}"
            f"00"
            f"{str(numero_lote).zfill(4)}"
            f"{self.data.strftime('%d%m%Y')}"
        )
        return header.ljust(378) + "00"

    def criar_trailer(self) -> str:
        # Cria o trailer do arquivo de saída
        return "99".ljust(378) + "00"

    def gerar_arquivo_mancad(self) -> Optional[str]:
        # Gera o arquivo de saída formatado
        try:
            processador_excel = ProcessadorArquivoExcel()
            registros = processador_excel.ler_arquivo()

            if not registros:
                Logger.log("Nenhum registro encontrado no arquivo.", 'ALERTA')
                return None

            valor_coluna_c = self.obter_e_incrementar_valor_coluna_c()

            nome_arquivo_saida = (
                f"NIOMANCAD{valor_coluna_c}_{self.data.strftime('%d%m%Y')}.txt"
            )
            caminho_arquivo_saida = os.path.join(
                Config.DIRETORIO_SAIDA, nome_arquivo_saida)

            os.makedirs(Config.DIRETORIO_SAIDA, exist_ok=True)

            with open(caminho_arquivo_saida, 'w', encoding='utf-8') as f:
                f.write(self.criar_header(valor_coluna_c) + '\n')

                registros_mancad = []
                for registro in registros:
                    registro_mancad = FormatadorMancad.formatar_registro(
                        registro)
                    if registro_mancad:
                        f.write(registro_mancad + '\n')
                        registros_mancad.append(registro_mancad)

                f.write(self.criar_trailer() + '\n')

            Logger.log(f"Arquivo gerado: {caminho_arquivo_saida}")
            Logger.log(f"Total de registros: {len(registros_mancad)}")
            return caminho_arquivo_saida

        except Exception as e:
            Logger.log(f"Erro ao gerar arquivo Mancad: {e}", 'ERRO')
            return None

def main():
    # Função principal que executa o processamento e geração de arquivos
    try:
        processador = ProcessadorMancad()
        arquivo_gerado = processador.gerar_arquivo_mancad()

        if arquivo_gerado:
            Logger.log("\nArquivos no diretório de saída:")
            for arquivo in os.listdir(Config.DIRETORIO_SAIDA):
                if arquivo.startswith("NIOMANCAD") and arquivo.endswith(".txt"):
                    Logger.log(arquivo)

    except Exception as e:
        Logger.log(f"Erro na execução principal: {e}", 'ERRO')

if __name__ == "__main__":
    main()
