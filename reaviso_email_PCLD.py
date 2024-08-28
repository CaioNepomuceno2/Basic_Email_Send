import os
import pandas as pd
import win32com.client
from datetime import datetime
from unidecode import unidecode
import re
import time

# Configuração inicial do diretório
try:
    local_folder = os.path.dirname(os.path.abspath(__file__))
except:
    local_folder = os.getcwd()

up_local_folder = os.path.abspath(os.path.join(local_folder, '..'))

# Carregar a lista de e-mails
lista_emails = pd.read_excel('lista_emails.xlsx')
lista_emails.columns = [unidecode(col) for col in lista_emails.columns]

# Data de hoje para registrar no processo
data_hoje = datetime.now().strftime('%d/%m/%Y')

# Instância do Outlook
outlook = win32com.client.Dispatch("Outlook.Application")

# Definir a lista de e-mails válidos e a lista de falhas
valid_email_list = []
failed_emails = []  # Inicialização da lista de e-mails falhos

def is_valid_email(email):
    email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
    return bool(re.match(email_regex, email))

def chunk_email_list(email_list, chunk_size=1):
    for i in range(0, len(email_list), chunk_size):
        yield email_list[i:i + chunk_size]

for _, row in lista_emails.iterrows():
    if not is_valid_email(row.EMAIL):
        print(f"Formato de e-mail inválido: {row.EMAIL}")
        failed_emails.append([row.EMAIL, 'Formato inválido', data_hoje])
        continue
    valid_email_list.append(row.EMAIL)

    whatsapp_link = row["WHATSAPP_LINK"]
    CDC = row["CDC"]
    VALOR_BRUTO = row["VALOR BRUTO"]
    VALOR_DESCONTO = row["VALOR C/ DESCONTO"]
    DESCONTO = row["% DESCONTO"]
    NOMECSD = row["NOMECSD"]

    with open(FR'{up_local_folder}\assets\reaviso_PCLD.html', 'r', encoding='utf8') as f:
        time.sleep(1)
        msg = outlook.CreateItem(0)
        
        msg.Subject = 'Acordo de Pagamento Personalizado para Você - ENERGISA MS'
        msg.Bcc = row.EMAIL
        msg.Cc = ''
        layout_email = f.read()

        layout_email = layout_email.replace('<var>CDC</var>', str(CDC))
        # layout_email = layout_email.replace('<var>VALOR_BRUTO</var>', 'R${:,.2f}'.format(VALOR_BRUTO).replace(',', 'X').replace('.', ',').replace('X', '.'))
        # layout_email = layout_email.replace('<var>VALOR_C_DESCONTO</var>', 'R${:,.2f}'.format(VALOR_DESCONTO).replace(',', 'X').replace('.', ',').replace('X', '.'))
        # layout_email = layout_email.replace('<var>DESCONTO_PERCENTUAL</var>', '{:.2f}%'.format(DESCONTO * 100))
        layout_email = layout_email.replace('<var>NOMECSD</var>', str(NOMECSD))

        layout_email = layout_email.replace('<var>WHATSAPP_LINK</var>', whatsapp_link)
        layout_email = layout_email.replace('_CAMINHO_IMAGENS_', up_local_folder)
        msg.HTMLBody = layout_email
        #msg.Display()
        msg.Send()
        lista_emails.loc[lista_emails['EMAIL'] == row.EMAIL, 'status'] = 'Enviado'

lista_emails.to_excel('lista_emails_validacao.xlsx', index=False)
df_failed_emails = pd.DataFrame(failed_emails, columns=['EMAIL', 'Erro', 'Data'])
df_failed_emails.to_excel('emails_invalidos.xlsx', index=False)

print("Processo concluído. E-mails inválidos registrados em 'emails_invalidos.xlsx'.")
