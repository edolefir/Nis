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
    resultstr=''
    logger.info("Запуск страницы.")
    submitbutton= request.POST.get("submit")
    if request.method == 'POST':
        form = SenderForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                logger.info("Начало работы программы. Обработка данных из формы...")
                LOGIN = form.cleaned_data.get('LOGIN')
                PWD = form.cleaned_data.get('PWD')
                MESSAGE = form.cleaned_data.get('LETTER')
                THEME = form.cleaned_data.get('THEME')
                MAIL_SERVER = 'smtp.gmail.com:587'
                try:
                    logger.info('Cчитывание таблицы...')
                    worksheet = pandas.read_excel(request.FILES['WHOM'], sheet_name=0)
                    if 'ATTACH_TPL' in request.FILES: 
                        tpl = DocxTemplate(request.FILES['ATTACH_TPL'])
                        logger.info('Файл шаблона приклеплён к форме.')
                    else:
                        logger.info('Файл шаблона не приклеплён к форме.')
                    logger.info('Проверка соединения...')
                    server = smtplib.SMTP(MAIL_SERVER)
                    server.starttls()  
                    server.login(LOGIN, PWD)
                    server.quit()
                except xlrd.biffh.XLRDError:
                    log.warning('Неверный формат файла таблицы - работа программы невозможна.')
                except ValueError or docx.opc.exceptions.PackageNotFoundError:
                    log.warning('Неверный формат файла шаблона - работа программы невозможна.')
                except smtplib.socket.gaierror:
                    logger.warning("Не удалось соединиться с сервером.")
                except smtplib.SMTPServerDisconnected:
                    logger.warning("Разрыв соединения.")
                except smtplib.SMTPAuthenticationError:
                    logger.warning("Неверный логин и пароль.")
                else:
                    logger.info('Все данные введены правильно. Вход прошел успешно. Таблица считана.')
                    MailData = CreateData(worksheet)
                    for mail in MailData:
                        msg = MIMEMultipart()
                        msg['From'] = LOGIN
                        msg['To'] = mail
                        msg['Subject'] = THEME
                        Text = MESSAGE
                        for param in MailData[mail]:
                            Text=Text.replace('{{ ' + param + ' }}', MailData[mail][param])
                        msg.attach(MIMEText(Text, "html"))
                        if 'ATTACH_TPL' in request.FILES: 
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
                            logger.info("Файл для " + mail + " сформирован и приклеплен")
                        try:
                            server = smtplib.SMTP(MAIL_SERVER)
                        except smtplib.socket.gaierror:
                            logger.warning("Не удалось соединиться с сервером. Письмо для " + mail + " не удалось отправить.")
                        else:
                            try:
                                server.starttls()  
                                server.login(LOGIN, PWD)
                                server.sendmail(LOGIN, mail , msg.as_string())
                            except smtplib.SMTPServerDisconnected:
                                logger.warning("Разрыв соединения - письмо для " + mail + " не удалось отправить.")
                            except smtplib.SMTPRecipientsRefused:
                                logger.warning("Неверный адресат - письмо для " + mail + " не удалось отправить.")
                            else:
                                logger.info("Письмо для " + mail + " отправлено.")
                                try:
                                    server.quit()
                                except smtplib.SMTPServerDisconnected:
                                    logger.warning("Разрыв соединения.")
                        time.sleep(3)
                    logger.info("Программа закончила свою работу.")
            except exception as e: 
                resultstr='Ошибка: ' + str(e)
            else: 
                resultstr='Сообщения отправлены'
    else:
        form = SenderForm()
    context= {'form': form, 'submitbutton': submitbutton, 'resultstr':resultstr}
    logger.info("Рендеринг страницы.")
    return render(request, 'app1/list.html', context)