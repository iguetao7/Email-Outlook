import win32com.client as win32
import os



#Conectando ao outlook
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)

#Enviando com conta secundária
# conta = outlook.Session.Accounts['OUTRA CONTA']
# email._oleobj_Invoke(*(64209, 0, 8, 0, conta))

#Informações do email
email.To = 'EMAIL DE DESTINO'
email.Cc = 'EMAIL DE DESTINO DA CÓPIA'
email.Bcc = 'EMAIL DE DESTINO DA CÓPIA OCULTA'
email.Subject = 'E-mail enviado pelo outlook'
#email.Body = 'Texto do e-mail'
email.HTMLBody = """<p>Primeiro parágrafo</p>
<p>Segundo parágrafo</p>
<img src='LINK DA IMAGEM'
width=200>"""

#Anexo que estão no computador
caminho_codigo = os.getcwd()
arquivo_anexar = 'anexos/ANEXO.png'
lista_arquivos = os.listdir('anexos')

#Vários anexos
for nome_arquivo in lista_arquivos:
    caminho_anexo = os.path.join(caminho_codigo, 'anexos', nome_arquivo)
    email.Attachments.Add(caminho_anexo)

#enviando o email
email.Send()
