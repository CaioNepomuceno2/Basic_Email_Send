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



def is_valid_email(email):
    """Check if the given email has a valid format using a regex."""
    email_regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
    return bool(re.match(email_regex, email))

def chunk_email_list(email_list, chunk_size=20):
    """Yield successive chunk_size-sized chunks from email_list."""
    for i in range(0, len(email_list), chunk_size):
        yield email_list[i:i + chunk_size]

grouped_emails = lista_emails.groupby("GESTORA")

failed_emails = []  # To store emails that failed due to invalid format

# For each gestora, send emails in chunks of up to 20
for gestora, group_data in grouped_emails:
    valid_email_list = []  # List to store valid emails for the current gestora
    
    for _, row in group_data.iterrows():
        if not is_valid_email(row.EMAIL):
            print(f"Invalid email format: {row.EMAIL}")
            failed_emails.append(row.EMAIL)
            continue
        valid_email_list.append(row.EMAIL)
    
    whatsapp_link = row["WHATSAPP_LINK"]  # Assuming the same WhatsApp link for all emails under a gestora
    
    # Splitting the valid email list into chunks of up to 20
    for email_chunk in chunk_email_list(valid_email_list):
        with open(FR'{up_local_folder}\assets\reaviso_desenrola copy 4.html', 'r', encoding='utf8') as f:
            time.sleep(3)
            msg = outlook.CreateItem(0)
            msg.Subject = 'ENERGISA: CONDIÇÕES ESPECIAIS - Propostas Flexiveis e Personalizadas'
            msg.Bcc = ';'.join(email_chunk)  # Setting emails in Bcc field
            msg.Cc = ''            
            layout_email = f.read()
            layout_email = layout_email.replace('<var>WHATSAPP_LINK</var>', whatsapp_link)
            layout_email = layout_email.replace('_CAMINHO_IMAGENS_', up_local_folder)            
            msg.HTMLBody = layout_email
            msg.Display()
            msg.Send()
            lista_emails.loc[lista_emails['EMAIL'] == row.EMAIL, 'status'] = 'Enviado'

    lista_emails.to_excel(f'lista_emails_validacao_{gestora}.xlsx', index=False)
