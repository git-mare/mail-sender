import smtplib
import csv
import time
import random
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import List, Dict, Optional


class EmailDestination:
    """Classe para representar um destinatário de e-mail."""
    
    def __init__(self, nome: str, email: str):
        self.nome = nome
        self.email = email.lower().strip()
    
    def __str__(self) -> str:
        return f"{self.nome} <{self.email}>"


class CSVReader:
    """Classe para leitura e processamento do arquivo CSV."""
    
    def __init__(self, filepath: str):
        self.filepath = filepath
    
    def read_valid_destinations(self) -> List[EmailDestination]:
        """Lê o arquivo CSV e retorna uma lista de destinatários válidos."""
        valid_dests = []
        
        try:
            with open(self.filepath, 'r', encoding='utf-8', errors='ignore') as arquivo:
                leitor = csv.DictReader(arquivo, skipinitialspace=True)
                for linha in leitor:
                    try:
                        if (linha['nome_fantasia'].strip() and 
                            linha['email'].strip() and 
                            '@' in linha['email']):
                            
                            valid_dests.append(EmailDestination(
                                nome=linha['nome_fantasia'],
                                email=linha['email']
                            ))
                    except KeyError:
                        continue
                    except Exception as e:
                        print(f"Erro na linha: {linha} | {str(e)}")
                        continue
        except Exception as e:
            print(f"Erro ao ler o arquivo CSV: {str(e)}")
        
        return valid_dests


class EmailSender:
    """Classe principal para envio de e-mails."""
    
    def __init__(self, email: str, senha: str, 
                 smtp_server: str = "smtp.gmail.com", 
                 porta: int = 587):
        self.email = email
        self.senha = senha
        self.smtp_server = smtp_server
        self.porta = porta
        self.servidor = None
    
    def conectar(self) -> None:
        """Estabelece conexão com o servidor SMTP."""
        try:
            self.servidor = smtplib.SMTP(self.smtp_server, self.porta)
            self.servidor.starttls()
            self.servidor.ehlo()
            self.servidor.login(self.email, self.senha)
            print("Conectado ao servidor SMTP com sucesso")
        except Exception as e:
            print(f"Erro ao conectar ao servidor SMTP: {str(e)}")
            raise
    
    def desconectar(self) -> None:
        """Encerra conexão com o servidor SMTP."""
        if self.servidor:
            self.servidor.quit()
            print("Desconectado do servidor SMTP")
    
    def criar_mensagem(self, destinatario: EmailDestination, 
                    assunto: str, corpo: str, 
                    anexo_path: Optional[str] = None) -> MIMEMultipart:
        """Cria a mensagem de e-mail personalizada."""
        msg = MIMEMultipart()
        msg['From'] = self.email
        msg['To'] = destinatario.email
        
        # Personalizar assunto e corpo com nome da empresa
        assunto_personalizado = assunto.format(nome=destinatario.nome)
        corpo_personalizado = corpo.format(nome=destinatario.nome)
        
        msg['Subject'] = assunto_personalizado
        msg.attach(MIMEText(corpo_personalizado, 'plain'))
        
        if anexo_path:
            try:
                with open(anexo_path, 'rb') as anexo:
                    part = MIMEApplication(anexo.read(), Name=anexo_path.split('/')[-1])
                    part['Content-Disposition'] = f'attachment; filename="{anexo_path.split("/")[-1]}"'
                    msg.attach(part)
            except Exception as e:
                print(f"Erro ao anexar arquivo {anexo_path}: {str(e)}")
        
        return msg

    def enviar_email(self, destinatario: EmailDestination, 
                    assunto: str, corpo: str, 
                    anexo_path: Optional[str] = None) -> bool:
        """Envia um e-mail para um destinatário específico."""
        try:
            if not self.servidor:
                self.conectar()
            
            msg = self.criar_mensagem(destinatario, assunto, corpo, anexo_path)
            self.servidor.sendmail(self.email, destinatario.email, msg.as_string())
            return True
        except smtplib.SMTPServerDisconnected:
            print("Conexão perdida. Tentando reconectar...")
            try:
                self.conectar()
                msg = self.criar_mensagem(destinatario, assunto, corpo, anexo_path)
                self.servidor.sendmail(self.email, destinatario.email, msg.as_string())
                return True
            except Exception as e:
                print(f"Falha ao reenviar para {destinatario.email}: {str(e)}")
                return False
        except Exception as e:
            print(f"Erro ao enviar para {destinatario.email}: {str(e)}")
            return False
    
    def enviar_em_massa(self, destinatarios: List[EmailDestination], 
                    assunto: str, corpo: str, 
                    anexo_path: Optional[str] = None,
                    intervalo_min: int = 30, 
                    intervalo_max: int = 60) -> None:
        """Envia e-mails em massa para uma lista de destinatários."""
        total = len(destinatarios)
        enviados = 0
        
        for idx, dest in enumerate(destinatarios, 1):
            if self.enviar_email(dest, assunto, corpo, anexo_path):
                enviados += 1
                print(f"[{idx}/{total}] Enviado para {dest.email}")
            
            # Reseta a conexão a cada 50 emails enviado
            if idx % 50 == 0:
                print("Renovando conexão SMTP...")
                self.desconectar()
                time.sleep(5)
                self.conectar()
            
            # Pausa entre envios (exceto para o último)
            if idx < total:
                espera = random.randint(intervalo_min, intervalo_max)
                print(f"Aguardando {espera} segundos...")
                time.sleep(espera)
        
        print(f"Processo concluído! {enviados}/{total} e-mails enviados com sucesso!")


class EmailCampaign:
    """Classe para gerenciar uma campanha de e-mails."""
    
    def __init__(self, email: str, senha: str, csv_path: str):
        self.csv_reader = CSVReader(csv_path)
        self.email_sender = EmailSender(email, senha)
    
    def executar(self, assunto: str, corpo: str, 
                anexo_path: Optional[str] = None,
                intervalo_min: int = 30, 
                intervalo_max: int = 60) -> None:
        """Executa a campanha de e-mail completa."""
        try:
            # Carregar destinatários válidos
            destinatarios = self.csv_reader.read_valid_destinations()
            print(f"Total de destinatários válidos: {len(destinatarios)}")
            
            if not destinatarios:
                print("Nenhum destinatário válido encontrado. Abortando.")
                return
            
            # Conectar ao servidor SMTP
            self.email_sender.conectar()
            
            # Enviar e-mails
            self.email_sender.enviar_em_massa(
                destinatarios=destinatarios,
                assunto=assunto,
                corpo=corpo,
                anexo_path=anexo_path,
                intervalo_min=intervalo_min,
                intervalo_max=intervalo_max
            )
        
        except Exception as e:
            print(f"Erro durante a campanha de e-mail: {str(e)}")
        
        finally:
            # Garantir desconexão do servidor SMTP
            self.email_sender.desconectar()


# Exemplo de uso do código:
if __name__ == "__main__":
    # Configurações
    seu_email = ""  # Substitua pelo seu Gmail
    senha_app = ""  # Use senha de app  (necessário 2FA habilitado)
    csv_path = "emails.csv"
    anexo_path = "" # Anexo
    
    # Assunto e corpo do e-mail
    assunto = "" # use {nome} para pegar o nome do destinatário
    corpo = """

    """
    
    # Criar e executar campanha
    campanha = EmailCampaign(seu_email, senha_app, csv_path)
    campanha.executar(assunto, corpo, anexo_path)