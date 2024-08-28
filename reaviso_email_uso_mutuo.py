# import os
# import pandas as pd
# import win32com.client
# from datetime import datetime
# from unidecode import unidecode
# import re
# import time

# # Configuração inicial do diretório
# try:
#     local_folder = os.path.dirname(os.path.abspath(__file__))
# except:
#     local_folder = os.getcwd()

# up_local_folder = os.path.abspath(os.path.join(local_folder, '..'))

# # Carregar a planilha
# base_cobranca = pd.read_excel('BASE_COBRANCA_USO_MUTUO_2205.xlsx', sheet_name='Planilha2')
# base_cobranca.columns = [unidecode(col) for col in base_cobranca.columns]

# # Data de hoje para registrar no processo
# data_hoje = datetime.now().strftime('%d/%m/%Y')

# # Instância do Outlook
# outlook = win32com.client.Dispatch("Outlook.Application")

# # Definir a lista de e-mails válidos e a lista de falhas
# valid_email_list = []
# failed_emails = []  # Inicialização da lista de e-mails falhos

# def is_valid_email(email):
#     email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
#     return bool(re.match(email_regex, email))

# def format_cnpj(cnpj):
#     cnpj = str(cnpj).zfill(14)
#     return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/0001-{cnpj[12:14]}"

# for _, row in base_cobranca.iterrows():
#     if not is_valid_email('caiohenriquenepomuceno@gmail.com'):  # Usar o seu e-mail para testes
#         print(f"Formato de e-mail inválido: caiohenriquenepomuceno@gmail.com")
#         failed_emails.append(['caiohenriquenepomuceno@gmail.com', 'Formato inválido', data_hoje])
#         continue
#     valid_email_list.append('caiohenriquenepomuceno@gmail.com')

#     nome = row["NOME"]
#     documento = format_cnpj(row["DOCUMENTO"])
#     contrato = row["Numero do Contrato"]
#     valor = row["Total Geral"]

#     with open(FR'{up_local_folder}\assets\reaviso_Uso_Mutuo.html', 'r', encoding='utf8') as f:
#         time.sleep(1)
#         msg = outlook.CreateItem(0)
        
#         msg.Subject = 'Notificação Extrajudicial para Constituição em Mora - ENERGISA MS'
#         msg.To = 'caiohenriquenepomuceno@gmail.com'
#         msg.Cc = ''
#         layout_email = f.read()

#         layout_email = layout_email.replace('<var>NOMECSD</var>', str(nome))
#         layout_email = layout_email.replace('<var>DOCUMENTO</var>', str(documento))
#         layout_email = layout_email.replace('<var>CONTRATO</var>', str(contrato))
#         layout_email = layout_email.replace('<var>VALOR</var>', f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))

#         layout_email = layout_email.replace('_CAMINHO_IMAGENS_', up_local_folder)
#         msg.HTMLBody = layout_email
#         msg.Display()  # Descomente para visualizar o e-mail antes de enviar
#         msg.Send()
#         base_cobranca.loc[base_cobranca['DOCUMENTO'] == row['DOCUMENTO'], 'status'] = 'Enviado'

# base_cobranca.to_excel('lista_emails_validacao.xlsx', index=False)
# df_failed_emails = pd.DataFrame(failed_emails, columns=['EMAIL', 'Erro', 'Data'])
# df_failed_emails.to_excel('emails_invalidos.xlsx', index=False)

# print("Processo concluído. E-mails inválidos registrados em 'emails_invalidos.xlsx'.")


# import os
# import pandas as pd
# import win32com.client
# from datetime import datetime
# from unidecode import unidecode
# import re
# import time

# # Configuração inicial do diretório
# try:
#     local_folder = os.path.dirname(os.path.abspath(__file__))
# except:
#     local_folder = os.getcwd()

# up_local_folder = os.path.abspath(os.path.join(local_folder, '..'))

# # Carregar a planilha
# base_cobranca = pd.read_excel('BASE_COBRANCA_USO_MUTUO_2205.xlsx', sheet_name='Planilha2')
# base_cobranca.columns = [unidecode(col) for col in base_cobranca.columns]

# # Data de hoje para registrar no processo
# data_hoje = datetime.now().strftime('%d/%m/%Y')

# # Instância do Outlook
# outlook = win32com.client.Dispatch("Outlook.Application")

# # Definir a lista de e-mails válidos e a lista de falhas
# valid_email_list = []
# failed_emails = []  # Inicialização da lista de e-mails falhos

# def is_valid_email(email):
#     email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
#     return bool(re.match(email_regex, email))

# def format_cnpj(cnpj):
#     cnpj = str(cnpj).zfill(14)
#     return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/0001-{cnpj[12:14]}"

# for _, row in base_cobranca.iterrows():
#     email_destino = 'caiohenriquenepomuceno@gmail.com'  # Usar o seu e-mail para testes
#     if not is_valid_email(email_destino):
#         print(f"Formato de e-mail inválido: {email_destino}")
#         failed_emails.append([email_destino, 'Formato inválido', data_hoje])
#         continue
#     valid_email_list.append(email_destino)

#     nome = row["NOME"]
#     documento = format_cnpj(row["DOCUMENTO"])
#     contrato = row["Numero do Contrato"]
#     valor = row["Total Geral"]
    
#     # Caminho do PDF correspondente ao nome
#     pdf_path = os.path.join(local_folder, 'Uso_Mutuo', f'{nome}.pdf')
    
#     if not os.path.isfile(pdf_path):
#         print(f"PDF não encontrado para: {nome}")
#         failed_emails.append([email_destino, 'PDF não encontrado', data_hoje])
#         continue

#     with open(FR'{up_local_folder}\assets\reaviso_Uso_Mutuo.html', 'r', encoding='utf8') as f:
#         time.sleep(1)
#         msg = outlook.CreateItem(0)
        
#         msg.Subject = 'Notificação Extrajudicial para Constituição em Mora - ENERGISA MS'
#         msg.To = email_destino
#         msg.Cc = ''
#         layout_email = f.read()

#         layout_email = layout_email.replace('<var>NOMECSD</var>', str(nome))
#         layout_email = layout_email.replace('<var>DOCUMENTO</var>', str(documento))
#         layout_email = layout_email.replace('<var>CONTRATO</var>', str(contrato))
#         layout_email = layout_email.replace('<var>VALOR</var>', f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.'))

#         layout_email = layout_email.replace('_CAMINHO_IMAGENS_', up_local_folder)
#         msg.HTMLBody = layout_email
#         msg.Attachments.Add(pdf_path)  # Anexar o PDF correspondente
#         msg.Display()  # Descomente para visualizar o e-mail antes de enviar

#         #msg.Send()
#         base_cobranca.loc[base_cobranca['DOCUMENTO'] == row['DOCUMENTO'], 'status'] = 'Enviado'

# base_cobranca.to_excel('lista_emails_validacao.xlsx', index=False)
# df_failed_emails = pd.DataFrame(failed_emails, columns=['EMAIL', 'Erro', 'Data'])
# df_failed_emails.to_excel('emails_invalidos.xlsx', index=False)

# print("Processo concluído. E-mails inválidos registrados em 'emails_invalidos.xlsx'.")





import os
import pandas as pd
import win32com.client
from datetime import datetime
from unidecode import unidecode
import re
import time

def is_valid_email(email):
    email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
    return bool(re.match(email_regex, email))

def format_cnpj(cnpj):
    cnpj = str(cnpj).zfill(14)
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/0001-{cnpj[12:14]}"

def format_currency(value):
    return f'R$ {value:,.2f}'.replace(',', 'v').replace('.', ',').replace('v', '.')

# Configuração inicial do diretório
try:
    local_folder = os.path.dirname(os.path.abspath(__file__))
except:
    local_folder = os.getcwd()

up_local_folder = os.path.abspath(os.path.join(local_folder, '..'))

# Carregar a planilha
base_cobranca = pd.read_excel('BASE_COBRANCA_USO_MUTUO_2205.xlsx', sheet_name='Planilha2')
base_cobranca.columns = [unidecode(col) for col in base_cobranca.columns]

# Data de hoje para registrar no processo
data_hoje = datetime.now().strftime('%d/%m/%Y')

# Instância do Outlook
outlook = win32com.client.Dispatch("Outlook.Application")

# Definir a lista de e-mails válidos e a lista de falhas
valid_email_list = []
failed_emails = []  # Inicialização da lista de e-mails falhos


pdf_folder = os.path.join(local_folder, 'pdfs')

for _, row in base_cobranca.iterrows():
    email_destino = row['E-mail']  # Pegando o e-mail da coluna 'E-mail'
    if not is_valid_email(email_destino):
        print(f"Formato de e-mail inválido: {email_destino}")
        failed_emails.append([email_destino, 'Formato inválido', data_hoje])
        continue
    valid_email_list.append(email_destino)

    nome = row["NOME"]
    documento = format_cnpj(row["DOCUMENTO"])
    contrato = row["Numero do Contrato"]
    valor = format_currency(row["Total Geral"])

    # Caminho do PDF correspondente ao nome
    pdf_path = os.path.join(local_folder, 'Uso_Mutuo', f'{nome}.pdf')

    if not os.path.isfile(pdf_path):
        print(f"PDF não encontrado para: {nome}")
        failed_emails.append([email_destino, 'PDF não encontrado', data_hoje])
        continue

    with open(FR'{up_local_folder}\assets\reaviso_Uso_Mutuo.html', 'r', encoding='utf8') as f:
        msg = outlook.CreateItem(0)
        
        msg.Subject = 'Notificação Extrajudicial para Constituição em Mora - ENERGISA MS'
        # msg.To = email_destino janaina
        
        msg.To = 'janaina.amorim@energisa.com.br'
        msg.Cc = ''
        layout_email = f.read()

        layout_email = layout_email.replace('<var>NOMECSD</var>', str(nome))
        layout_email = layout_email.replace('<var>DOCUMENTO</var>', str(documento))
        layout_email = layout_email.replace('<var>CONTRATO</var>', str(contrato))
        layout_email = layout_email.replace('<var>VALOR</var>', valor)

        layout_email = layout_email.replace('_CAMINHO_IMAGENS_', up_local_folder)
        msg.HTMLBody = layout_email
        msg.Attachments.Add(pdf_path)  # Anexar o PDF correspondente
        msg.Display()  # Descomente para visualizar o e-mail antes de enviar
        
        #msg.Send()
        base_cobranca.loc[base_cobranca['DOCUMENTO'] == row['DOCUMENTO'], 'status'] = 'Enviado'

# Salvar planilha com status de envio
base_cobranca.to_excel('lista_emails_validacao.xlsx', index=False)

# Registrar e-mails inválidos
df_failed_emails = pd.DataFrame(failed_emails, columns=['EMAIL', 'Erro', 'Data'])
df_failed_emails.to_excel('emails_invalidos.xlsx', index=False)

print("Processo concluído. E-mails inválidos registrados em 'emails_invalidos.xlsx'.")
