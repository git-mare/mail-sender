Este projeto é um script em Python para **envio de e-mails em massa** de forma personalizada, utilizando arquivos CSV como base de destinatários.

---

## Funcionalidades

- Leitura de destinatários a partir de um arquivo CSV.
- Validação automática de e-mails válidos.
- Envio de mensagens personalizadas com **placeholders** (ex: `{nome}`).
- Suporte a anexos (PDF, imagens, documentos).
- Intervalo aleatório entre envios para evitar bloqueios.
- Reconexão automática em caso de perda de conexão SMTP.
- Relatórios de progresso durante o envio.

---

## Estrutura do Projeto
```
├── emails.csv # Lista de destinatários
├── main.py # Script principal
└── README.md # Documentação
```

## Formato do CSV

O script espera um arquivo CSV no seguinte formato:

```csv
nome_fantasia,email
Empresa Exemplo,contato@empresa.com
Outra Empresa,contato@outra.com
```

## Configuração
1. Instale as dependências (Python 3.8+):
```
pip install --upgrade pip
```

2. Configuração do Gmail:
- Ative a autenticação em duas etapas no Gmail.
- Gere uma senha de app para usar no lugar da senha normal.
- Substitua no código:
```
seu_email = "" # Seu Gmail
senha_app = "" # Sua senha de app
```

3. Configuração da campanha:
```
csv_path = "emails.csv"     # Caminho para o arquivo CSV
anexo_path = "arquivo.pdf"  # Caminho do anexo (opcional)

assunto = "Olá, {nome}! Temos novidades para você"
corpo = """
Olá {nome},

Gostaríamos de apresentar nossas novas soluções para sua empresa.

Atenciosamente,
Equipe Exemplo
"""
```

## Uso
```
python main.py
```

## Informativo
- Na classe `EmailCampaign.executar`, você pode ajustar:

- Intervalo entre envios (em segundos, padrão: 30 a 60s entre e-mails):
  - `intervalo_min=30, intervalo_max=60`
- A cada 50 e-mails enviados, a conexão SMTP é renovada.