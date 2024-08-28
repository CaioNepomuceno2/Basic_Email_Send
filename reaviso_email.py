
import os
import pandas as pd
import win32com.client
from datetime import datetime
from unidecode import unidecode
import time
import re


# In[31]:


try:
    local_folder = os.path.dirname(os.path.abspath(__file__))
except:
    local_folder = os.getcwd()


# In[32]:


lista_emails = pd.read_excel('lista_emails.xlsx')


# In[33]:


lista_emails.columns = [unidecode(col) for col in lista_emails.columns]


data_hoje = datetime.now().strftime('%d/%m/%Y')


# In[36]:


outlook = win32com.client.Dispatch("Outlook.Application")


# In[37]:


up_local_folder = os.path.abspath(FR'{local_folder}\..')


# In[38]:


import time
import re

tempo = time.sleep(5)
def is_valid_email(email):
    email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
    return bool(re.match(email_regex, email))

def chunk_email_list(email_list, chunk_size=1):

    for i in range(0, len(email_list), chunk_size):
        yield email_list[i:i + chunk_size]

# grouped_emails = lista_emails.groupby("GESTORA")

# failed_emails = []
# for gestora, group_data in grouped_emails:
#     valid_email_list = [] 
    
#     for _, row in group_data.iterrows():
#         if not is_valid_email(row.EMAIL):
#             print(f"Invalid email format: {row.EMAIL}")
#             failed_emails.append(row.EMAIL)
#             continue
#         valid_email_list.append(row.EMAIL)
    


#     for email_chunk in chunk_email_list(valid_email_list):

#         whatsapp_link = row["WHATSAPP_LINK"] 
#         CDC = row["CDC"]
#         VALOR_BRUTO = row["VALOR BRUTO"]
#         VALOR_DESCONTO = row["VALOR C/ DESCONTO"]
#         DESCONTO = row["% DESCONTO"]
#         NOMECSD = row["NOMECSD"]

#         with open(FR'{up_local_folder}\assets\reaviso_desenrola.html', 'r', encoding='utf8') as f:
#             time.sleep(1)
#             msg = outlook.CreateItem(0)
            
#             msg.Subject = 'DESENROLA BRASIL - ENERGISA MS'
#             msg.Bcc = ';'.join(email_chunk)
#             msg.Cc = ''            
#             layout_email = f.read()

#             layout_email = layout_email.replace('<var>CDC</var>', str(CDC))
#             layout_email = layout_email.replace('<var>VALOR_BRUTO</var>', 'R${:,.2f}'.format(VALOR_BRUTO).replace(',', 'X').replace('.', ',').replace('X', '.'))
#             layout_email = layout_email.replace('<var>VALOR_C_DESCONTO</var>', 'R${:,.2f}'.format(VALOR_DESCONTO).replace(',', 'X').replace('.', ',').replace('X', '.'))
#             layout_email = layout_email.replace('<var>DESCONTO_PERCENTUAL</var>', '{:.2f}%'.format(DESCONTO * 100))
#             layout_email = layout_email.replace('<var>NOMECSD</var>', str(NOMECSD))


#             layout_email = layout_email.replace('<var>WHATSAPP_LINK</var>', whatsapp_link)
#             layout_email = layout_email.replace('_CAMINHO_IMAGENS_', up_local_folder)            
#             msg.HTMLBody = layout_email
#             msg.Display()
#             msg.Send()
#             for email in email_chunk:
#                 lista_emails.loc[lista_emails['EMAIL'] == email, 'status'] = 'Enviado'

#     lista_emails.to_excel('lista_emails_validacao.xlsx', index=False)

valid_email_list = []

for _, row in lista_emails.iterrows():
    if not is_valid_email(row.EMAIL):
        print(f"Invalid email format: {row.EMAIL}")
        failed_emails.append(row.EMAIL)
        continue
    valid_email_list.append(row.EMAIL)

    whatsapp_link = row["WHATSAPP_LINK"]
    CDC = row["CDC"]
    VALOR_BRUTO = row["VALOR BRUTO"]
    VALOR_DESCONTO = row["VALOR C/ DESCONTO"]
    DESCONTO = row["% DESCONTO"]
    NOMECSD = row["NOMECSD"]

    with open(FR'{up_local_folder}\assets\reaviso_desenrola.html', 'r', encoding='utf8') as f:
        time.sleep(1)
        msg = outlook.CreateItem(0)
        
        msg.Subject = 'DESENROLA BRASIL - ENERGISA MS'
        msg.Bcc = row.EMAIL
        msg.Cc = ''
        layout_email = f.read()

        layout_email = layout_email.replace('<var>CDC</var>', str(CDC))
        layout_email = layout_email.replace('<var>VALOR_BRUTO</var>', 'R${:,.2f}'.format(VALOR_BRUTO).replace(',', 'X').replace('.', ',').replace('X', '.'))
        layout_email = layout_email.replace('<var>VALOR_C_DESCONTO</var>', 'R${:,.2f}'.format(VALOR_DESCONTO).replace(',', 'X').replace('.', ',').replace('X', '.'))
        layout_email = layout_email.replace('<var>DESCONTO_PERCENTUAL</var>', '{:.2f}%'.format(DESCONTO * 100))
        layout_email = layout_email.replace('<var>NOMECSD</var>', str(NOMECSD))

        layout_email = layout_email.replace('<var>WHATSAPP_LINK</var>', whatsapp_link)
        layout_email = layout_email.replace('_CAMINHO_IMAGENS_', up_local_folder)
        msg.HTMLBody = layout_email
        msg.Display()
        msg.Send()
        lista_emails.loc[lista_emails['EMAIL'] == row.EMAIL, 'status'] = 'Enviado'

lista_emails.to_excel('lista_emails_validacao.xlsx', index=False)

            