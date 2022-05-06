import pandas as pd
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')

gerentes_df = pd.read_excel('Enviar E-mails.xlsx')
#gerentes_df.info()

for i, email in enumerate(gerentes_df['E-mail']):
    gerente = gerentes_df.loc[i, 'Gerente']
    area = gerentes_df.loc[i, 'Relatório']
    
    mail = outlook.CreateItem(0)
    mail.To = email
    # mail.CC = 'contato@iluminarfotografia.com.br'  # copia
    # mail.BCC = 'email@gmail.com'                   # copia oculta

    # assunto
    mail.Subject = 'Relatório de {}'.format(area)
    
    # corpo do email
    mail.Body = '''
    Prezado {}, 
    Segue em anexo o Relatório de {}, conforme solicitado.
    Qualquer dúvida estou à disposição.
    Att.
    '''.format(gerente, area)

    # Anexo
    attachment  = r'CAMINHO PARA ARQUIVO DO ANEXO, EX:(C:\Users\Usuário\Desktop\{}.xlsx)'.format(area)
    mail.Attachments.Add(attachment)

    mail.Send()