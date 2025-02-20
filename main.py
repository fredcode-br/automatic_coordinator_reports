import win32com.client
import os
import smtplib
import logging
import sys
import time
from dotenv import load_dotenv
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Carregar variáveis do arquivo .env
load_dotenv()

EMAIL = os.getenv("EMAIL")
SENHA = os.getenv("SENHA")


# Configuração de logging
logging.basicConfig(
    filename="relatorios.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%d/%m/%Y %H:%M:%S"
)

# Criar uma classe para duplicar a saída do print
class DualStream:
    def write(self, message):
        logging.info(message)  # Registra no log (mantém linhas em branco)
        sys.__stdout__.write(message)  # Exibe no terminal

    def flush(self):
        sys.__stdout__.flush()

# Redirecionar stdout para DualStream
sys.stdout = DualStream()

def enviar_email(destinatario, assunto, corpo, arquivos_anexos):
    try:
        remetente = EMAIL
        senha = SENHA
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['Subject'] = assunto

        msg.attach(MIMEText(corpo, 'plain'))

        print(f"Enviando e-mail para {destinatario}...")

        # Anexar os arquivos PDF
        for arquivo in arquivos_anexos:
            if os.path.exists(arquivo):
                with open(arquivo, "rb") as anexo:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(anexo.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(arquivo)}")
                    msg.attach(part)
            else:
                print(f"Aviso: Arquivo não encontrado - {arquivo}")


        with smtplib.SMTP_SSL('smtp.mailcorp.com.br', 465) as servidor:
            servidor.login(remetente, senha)
            servidor.sendmail(remetente, destinatario, msg.as_string())
        print(f"E-mail enviado para {destinatario}.")
    except Exception as e:
        print(f"Erro ao enviar o e-mail para {destinatario}: {e}")

def relatorios(workbook, planilha_relatorio, tabela_relatorio, codigo, coluna_filtro, pasta_destino):
    print(f'GERANDO RELATÓRIOS PARA PEDIDOS {planilha_relatorio}...')

    try:
        os.makedirs(pasta_destino, exist_ok=True)  # Garante que a pasta de destino exista
        
        sheet = workbook.Sheets(planilha_relatorio)
        tabela = sheet.ListObjects(tabela_relatorio)

        # --- FILTRAR A TABELA ---
        tabela.Range.AutoFilter(Field=coluna_filtro, Criteria1=f"={codigo}")

        # --- GERAR PDF ---
        caminho_pdf = os.path.join(pasta_destino, f'PEDIDOS {planilha_relatorio}.pdf')
        sheet.ExportAsFixedFormat(0, caminho_pdf)  # 0 = PDF

        return caminho_pdf
    
    except Exception as e:
        print(f"\nErro ao gerar o relatório: {e}")
        return None
    
    finally:
        try:
            # --- REMOVER FILTRO PARA NÃO AFETAR OUTRAS OPERAÇÕES ---
            tabela.AutoFilter.ShowAllData()
        except:
            pass  # Se não houver filtro, ignora o erro

def atualizarDados(caminho_arquivo_xlsm, data_inicial, data_final, pasta_destino):
    try:
        fechar_instancias_excel()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        print("Abrindo o arquivo...")
        workbook = excel.Workbooks.Open(os.path.abspath(caminho_arquivo_xlsm))

        sheet = workbook.Sheets("PARÂMETROS")
        sheet.Range("C5").Value = "'" + data_inicial.strftime("%d/%m/%Y")
        sheet.Range("C6").Value = "'" + data_final.strftime("%d/%m/%Y")
        
        print("Atualizando dados...")
        # workbook.RefreshAll()
        # time.sleep(150) # Tempo para atualizar os dados
        print("Salvando o arquivo...")
        workbook.Save()
        time.sleep(30)  # Garantir que o Excel processe o comando Save

        if not os.path.exists(caminho_arquivo_xlsm):
            print(f"Erro: O arquivo {caminho_arquivo_xlsm} não foi salvo corretamente.")
        else:
            print(f"Arquivo salvo com sucesso em {caminho_arquivo_xlsm}.")

        # --- PEGAR DADOS DA PLANILHA "COORD" ---
        sheet_coord = workbook.Sheets("COORD")
        destinatarios = []

        linha = 2  # (linha 1 é o cabeçalho)
        while True:
            codigo = sheet_coord.Cells(linha, 1).Value  # Coluna A
            nome = sheet_coord.Cells(linha, 2).Value  # Coluna B
            email = sheet_coord.Cells(linha, 3).Value  # Coluna C

            if not codigo:  # Se a coluna "codigo" estiver vazia, encerra o loop
                break

            destinatarios.append({
                "codigo": codigo,
                "nome": nome,
                "email": email
            })
            
            linha += 1

        # --- PERCORRER A LISTA ---
        for dest in destinatarios:
            print(f"\n***** GERANDO E-MAILS PARA O COORDERNADOR {dest['nome']} - {dest['codigo']} *****\n")
            relatorio_em_aberto  = relatorios(workbook, "EM ABERTO", "Tabela_Em_Aberto", dest['codigo'], 6, f'{pasta_destino}\\Coordenador_{dest['codigo']}')
            relatorio_faturado = relatorios(workbook, "FATURADOS", "Tabela_Faturados", dest['codigo'], 10, f'{pasta_destino}\\Coordenador_{dest['codigo']}')
           
            assunto = f"Relatório Diário de Pedidos"
            corpo = (
                f"Seguem os relatórios de pedidos faturados e em berto para acompanhamento referente ao coordenador {dest['nome']}.\n\n"
                "Favor não responder a este e-mail.\n\n"
                "Atenciosamente,\nEquipe TI Bioleve"
            )
           
            enviar_email(dest["email"], assunto, corpo, [relatorio_em_aberto, relatorio_faturado])

    except Exception as e:
        print(f"\nErro ao processar os dados: {e}")
    finally:
        try:
            print("\nFechando o arquivo e o Excel...")
            workbook.Close(SaveChanges=False)
            excel.Quit()
        except Exception as e:
            print(f"\nErro ao fechar o Excel: {e}")

def fechar_instancias_excel():
    os.system("taskkill /f /im excel.exe >nul 2>&1")

def enviar_logs_do_dia(destinatario):
    try:
        hoje = datetime.now().strftime("%d/%m/%Y")
        logs_do_dia = []

        # Ler os logs do arquivo original
        with open("relatorios.log", "r") as log_file:
            for linha in log_file:
                if linha.startswith(hoje):  # Filtrar apenas as linhas do dia atual
                    logs_do_dia.append(linha)

        if not logs_do_dia:
            print("Nenhum log do dia atual encontrado.")
            return

        # Criar um arquivo temporário com os logs do dia
        caminho_temporario = "logs_do_dia.log"
        with open(caminho_temporario, "w") as temp_file:
            temp_file.writelines(logs_do_dia)

        # Enviar o arquivo por e-mail
        assunto = "Logs Pedidos em Aberto e Faturados"
        corpo = "Segue em anexo os logs gerados no dia atual."
        caminho_temporario = [f'C:\\Scripts\\Relatórios_Coordenadores\\{caminho_temporario}']

        enviar_email(destinatario, assunto, corpo, caminho_temporario)

        print("Logs do dia enviados.")
    except Exception as e:
        print(f"Erro ao enviar os logs do dia: {e}")

# Caminhos dos arquivos
caminho_arquivo_xlsm = r"C:\Scripts\Relatórios_Coordenadores\Dados.xlsb"
pasta_destino = r"C:\Scripts\Relatórios_Coordenadores\relatorios"

data_final = datetime.now()
data_inicial = datetime(data_final.year, data_final.month, 1, 1, 1, 1, 279706)

# Execução principal
atualizarDados(caminho_arquivo_xlsm, data_inicial, data_final, pasta_destino)

# Enviar logs do dia
enviar_logs_do_dia("relatorios@bioleve.com.br")