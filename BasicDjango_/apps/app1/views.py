# -*- coding: utf-8 -*-
from django.shortcuts import render
from django.http import HttpResponse
import pandas


import smtplib

from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
from .functions import changeName
from .functions import CreateKeys
from .forms import SenderForm
import xlrd
import time

from docxtpl import DocxTemplate

def toUTF8(value):
	if type(value) == unicode:
		return value.encode('utf-8')
	if type(value) == str:
		return value
	return str(value)

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
            Keys= CreateKeys(worksheet)
            if 'ATTACH_TPL' in request.FILES:
                attachment = changeName(request.FILES['ATTACH_TPL'].name)
                if attachment.find('doc') > 0:
                    attachFile = MIMEBase('application', 'msword')
                elif attachment.find('pdf') > 0:
                    attachFile = MIMEBase('application', 'pdf')
                else:
                    attachFile = MIMEBase('application', 'octet-stream')
                    
                attachFile.set_payload(request.FILES['ATTACH_TPL'].read())
                encoders.encode_base64(attachFile)
                attachFile.add_header('Content-Disposition', 'attachment', filename=attachment)
            if 'ATTACH_TPL' in request.FILES:
                tpl = DocxTemplate(request.FILES['ATTACH_TPL'])
            for i in worksheet.index:
                msg = MIMEMultipart()
                msg['From'] = LOGIN
                msg['To'] = worksheet['Email'][i]
                msg['Subject'] = THEME
                Text = MESSAGE
                for param in Keys:
                    Text=Text.replace('__' + param + '__',worksheet[param][i])
                msg.attach(MIMEText(Text))

                if 'ATTACH_TPL' in request.FILES:
                    msg.attach(attachFile)

                server = smtplib.SMTP(MAIL_SERVER)  
                server.starttls()  
                server.login(LOGIN, PWD)
                server.sendmail(LOGIN, worksheet['Email'][i] , msg.as_string())
                server.quit()
                time.sleep(3)
    else:
        form = SenderForm()
    context= {'form': form, 'submitbutton': submitbutton}

  
    return render(request, 'app1/list.html', context)