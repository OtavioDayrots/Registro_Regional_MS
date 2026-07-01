import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import time 
from datetime import datetime, timedelta
import os
email_list = ['Elthon.santos@sanesul.ms.gov.br','gerentes-regionais@sanesul.ms.gov.br','equipe.tecnica-ctco@sanesul.ms.gov.br']
cc_list = ['vitor.rodrigues@sanesul.ms.gov.br','fernando.jorge@sanesul.ms.gov.br']


regional = ['Aquidauana', 'Dourados', 'Paranaiba', 'Tres Lagoas', 'Nova Andradina', 'Pontã Pora', 'Coxim', 'Corumbá', 'Naviraí', 'Jardim']
mes_num = (datetime.now() - timedelta(days=30)).strftime('%m')
mes_ext_map = {'01':'Janeiro','02':'Fevereiro','03':'Março','04':'Abril','05':'Maio','06':'Junho','07':'Julho','08':'Agosto','09':'Setembro','10':'Outubro','11':'Novembro','12':'Dezembro'}
mes_extenso = mes_ext_map[mes_num]
ano_ref = (datetime.now() - timedelta(days=30)).strftime('%Y')

# Obter o diretório onde o script está localizado
script_dir = os.path.dirname(os.path.abspath(__file__))
estado_dir = os.path.join(script_dir, 'estado')
os.makedirs(estado_dir, exist_ok=True)
assinatura_img_path = os.path.join(script_dir, 'assinatura.png')

def send_email(subject, body, to_email, attachment_path=None, assinatura_img_path=None):
    # Se attachment_path for obrigatório e não existir, não envia
    if attachment_path and not os.path.exists(attachment_path):
        print(f"[ERRO] Arquivo de anexo NÃO encontrado ({attachment_path}). E-mail para {to_email} NÃO será enviado.")
        return False
    
    sender_email = 'ctco@sanesul.ms.gov.br'
    password = 'INTEGRAsan!1'

    # Container principal para anexos
    message = MIMEMultipart('mixed')
    message["From"] = sender_email
    message["To"] = ", ".join(to_email)
    message["Subject"] = subject
    message["Cc"] = ", ".join(cc_list)

    # Container para corpo do e-mail com imagem embutida
    msg_related = MIMEMultipart('related')
    

    #Correção do código dentro do body html
    body_replace = body.replace('\n', '<br>')
    # Corpo do e-mail em HTML
    html_body = f'''
    <html>
      <body>
        <p>{body_replace}</p>
        <br>
        <img src="cid:assinatura_img">
      </body>
    </html>
    '''
    msg_alternative = MIMEMultipart('alternative')
    msg_alternative.attach(MIMEText(html_body, 'html'))
    msg_related.attach(msg_alternative)

    # Embutir a imagem da assinatura no corpo
    if assinatura_img_path and os.path.exists(assinatura_img_path):
        print(f"[DEBUG] Encontrado arquivo de assinatura: {assinatura_img_path}")
        with open(assinatura_img_path, 'rb') as img:
            mime_img = MIMEImage(img.read())
            mime_img.add_header('Content-ID', '<assinatura_img>')
            mime_img.add_header('Content-Disposition', 'inline', filename=os.path.basename(assinatura_img_path))
            msg_related.attach(mime_img)
    else:
        print(f"[DEBUG] Arquivo de assinatura NÃO encontrado: {assinatura_img_path}")

    # Adicionar o corpo completo ao container principal
    message.attach(msg_related)

    # Anexar arquivo como anexo separado
    if attachment_path:
        if os.path.exists(attachment_path):
            print(f"[DEBUG] Encontrado arquivo de anexo: {attachment_path}")
            with open(attachment_path, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(attachment_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_path)}"'
            message.attach(part)
            print(f"[DEBUG] Anexo adicionado ao e-mail: {os.path.basename(attachment_path)}")
        else:
            print(f"[ERRO] Arquivo de anexo NÃO encontrado: {attachment_path}")
            print(f"[DEBUG] Diretório atual: {os.getcwd()}")
            print(f"[DEBUG] Caminho procurado: {os.path.abspath(attachment_path)}")

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.send_message(message) 
        server.quit()
        print(f"[DEBUG] E-mail enviado para {to_email}")
        return True
    except Exception as e:
        print(f"Erro ao enviar email para {to_email}: {e}")
        return False

def main():
    subject = f'Relatório Mensal do Cadastro Técnico - Regional - {mes_extenso}/{ano_ref}'
    body = f'''Bom dia a todos!!<br>Segue anexo a este e-mail a planilha do Relatório Geral Mensal, referente ao mês de {mes_extenso.lower()}, das regionais operadas pela Sanesul.<br> Obs: Foi adicionado um comentário em cada título de cada coluna, para ver o comentário basta aproximar o mouse no título da respectiva coluna.<br><br>Atenciosamente,'''
    
    # Usar caminho absoluto baseado na localização do script
    attachment_path = os.path.join(estado_dir, 'registroRegional.xlsx')
    
    print(f"[DEBUG] Diretório do script: {script_dir}")
    print(f"[DEBUG] Caminho do anexo: {attachment_path}")
    print(f"[DEBUG] Arquivo existe? {os.path.exists(attachment_path)}")
    
    # Verificar se arquivo existe antes de iniciar envio
    if not os.path.exists(attachment_path):
        print(f"[ERRO] Arquivo de anexo NÃO encontrado: {attachment_path}. Processo abortado.")
        return

    send_email(subject, body, email_list, attachment_path=attachment_path, assinatura_img_path=assinatura_img_path)
    time.sleep(1)

if __name__ == '__main__':
    main()