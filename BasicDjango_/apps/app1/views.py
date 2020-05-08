from django.shortcuts import render
from django.http import HttpResponse
import pandas


import smtplib

from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from .functions import changeName
from .functions import CreateData
from .forms import SenderForm
import xlrd
import time
import logging

from docxtpl import DocxTemplate

def toUTF8(value):
	if type(value) == unicode:
		return value.encode('utf-8')
	if type(value) == str:
		return value
	return str(value)

logger=logging.getLogger('django.sender')

def sender(request):
    
    submitbutton= request.POST.get("submit")
    if request.method == 'POST':
        form = SenderForm(request.POST, request.FILES)
        if form.is_valid():
            LOGIN = form.cleaned_data.get('LOGIN')
            PWD = form.cleaned_data.get('PWD')
            MESSAGE = form.cleaned_data.get('LETTER')
            THEME = form.cleaned_data.get('THEME')
            MAIL_SERVER = 'smtp.gmail.com:587'
            worksheet = pandas.read_excel(request.FILES['WHOM'], sheet_name=0)
            MailData = CreateData(worksheet)
            for mail in MailData:
                msg = MIMEMultipart()
                msg['From'] = LOGIN
                msg['To'] = mail
                msg['Subject'] = THEME
                Text = MESSAGE
                for param in MailData[mail]:
                    Text=Text.replace('{{ ' + param + ' }}', MailData[mail][param])
                msg.attach(MIMEText(Text))
                if 'ATTACH_TPL' in request.FILES: 
                    tpl = DocxTemplate(request.FILES['ATTACH_TPL'])
                    context2 = {param : MailData[mail][param] for param in MailData[mail]}
                    tpl.render(context2)
                    tpl.save('Letter.docx')
                    attachment = 'Letter.docx'
                    if attachment.find('doc') > 0:
                        attachFile = MIMEBase('application', 'msword')
                    elif attachment.find('pdf') > 0:
                        attachFile = MIMEBase('application', 'pdf')
                    else:
                        attachFile = MIMEBase('application', 'octet-stream')
                    
                    fo = open('Letter.docx', 'rb')
                    attachFile.set_payload(fo.read())
                    fo.close()
                    encoders.encode_base64(attachFile)
                    FileName = request.FILES['ATTACH_TPL'].name.split('.')[0]  + '_To_' + mail.split('@')[0] + '.docx'
                    attachFile.add_header('Content-Disposition', 'attachment', filename=FileName)
                    msg.attach(attachFile)

                server = smtplib.SMTP(MAIL_SERVER)  
                server.starttls()  
                server.login(LOGIN, PWD)
                logger.info('Успешно введено')
                server.sendmail(LOGIN, mail , msg.as_string())
                server.quit()
                time.sleep(3)
    else:
        form = SenderForm()
    context= {'form': form, 'submitbutton': submitbutton}

  
    return render(request, 'app1/list.html', context)