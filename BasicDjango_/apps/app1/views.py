from django.shortcuts import render
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

logger=logging.getLogger('django.sender')
MAIL_SERVERS = {'gmail.com': 'smtp.gmail.com:465',
                'miem.hse.ru': 'smtp.gmail.com:465',
                'mail.ru': 'smtp.mail.ru:465',
                'hse.ru': 'smtpnnov.hse.ru:587',
                'yandex.ru': 'smtp.yandex.ru:465',
                'rambler.ru': 'smtp.rambler.ru:465'}

def sender(request):
    resultstr=''
    isError = False
    failedEmails = []
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
                MAIL_SERVER = MAIL_SERVERS[LOGIN.split('@')[1]]
                try:
                    logger.info('Cчитывание таблицы...')
                    worksheet = pandas.read_excel(request.FILES['WHOM'], sheet_name=0)
                    if 'ATTACH_TPL' in request.FILES: 
                        tpl = DocxTemplate(request.FILES['ATTACH_TPL'])
                        logger.info('Файл шаблона приклеплён к форме.')
                    else:
                        logger.info('Файл шаблона не приклеплён к форме.')
                    logger.info('Проверка соединения...')
                    server = smtplib.SMTP_SSL(MAIL_SERVER)
                    server.login(LOGIN, PWD)
                    server.quit()
                except xlrd.biffh.XLRDError:
                    isError = True
                    resultstr = 'Неверный формат файла таблицы - работа программы невозможна.'
                    logger.warning(resultstr)
                except ValueError or docx.opc.exceptions.PackageNotFoundError:
                    isError = True
                    resultstr = 'Неверный формат файла шаблона - работа программы невозможна.'
                    logger.warning(resultstr)
                except smtplib.socket.gaierror:
                    isError = True
                    resultstr = "Не удалось соединиться с сервером."
                    logger.warning(resultstr)
                except smtplib.SMTPServerDisconnected:
                    isError = True
                    resultstr = "Разрыв соединения."
                    logger.warning(resultstr)
                except smtplib.SMTPAuthenticationError:
                    isError = True
                    resultstr = "Неверный логин и пароль."
                    logger.warning(resultstr)
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
                            Text=Text.replace('{{ ' + param + ' }}', str(MailData[mail][param]))
                        msg.attach(MIMEText(Text, "html"))
                        if 'ATTACH_TPL' in request.FILES: 
                            tpl = DocxTemplate(request.FILES['ATTACH_TPL'])
                            context2 = {param : str(MailData[mail][param]) for param in MailData[mail]}
                            tpl.render(context2)
                            tpl.save('Letter.docx')
                            attachment = 'Letter.docx'
                            attachFile = MIMEBase('application', 'msword')
                            fo = open('Letter.docx', 'rb')
                            attachFile.set_payload(fo.read())
                            fo.close()
                            encoders.encode_base64(attachFile)
                            FileName = request.FILES['ATTACH_TPL'].name.split('.')[0]  + '_To_' + mail.split('@')[0] + '.docx'
                            attachFile.add_header('Content-Disposition', 'attachment', filename=FileName)
                            msg.attach(attachFile)
                            logger.info("Файл для " + mail + " сформирован и приклеплен")
                        try:
                            server = smtplib.SMTP_SSL(MAIL_SERVER)
                        except smtplib.socket.gaierror:
                            logger.warning("Не удалось соединиться с сервером. Письмо для " + mail + " не удалось отправить.")
                            failedEmails.append(mail + '-' + 'не удалось соединиться с сервером.')
                        else:
                            try:
                                server.login(LOGIN, PWD)
                                server.sendmail(LOGIN, mail , msg.as_string())
                            except smtplib.SMTPServerDisconnected:
                                logger.warning("Разрыв соединения - письмо для " + mail + " не удалось отправить.")
                                failedEmails.append(mail + '-' + 'разрыв соединения.')
                            except smtplib.SMTPRecipientsRefused:
                                logger.warning("Неверный адресат - письмо для " + mail + " не удалось отправить.")
                                failedEmails.append(mail + '-' + 'неверный адресат.')
                            else:
                                logger.info("Письмо для " + mail + " отправлено.")
                                try:
                                    server.quit()
                                except smtplib.SMTPServerDisconnected:
                                    logger.warning("Разрыв соединения.")
                        time.sleep(3)
                    logger.info("Программа закончила свою работу.")
            except Exception as e: 
                resultstr='Ошибка: ' + str(e)
            else:
                if not(isError):
                    resultstr = 'Сообщения отправлены'
    else:
        form = SenderForm()
    context= {'form': form, 'submitbutton': submitbutton, 'resultstr':resultstr, 'failedEmails': failedEmails}
    logger.info("Рендеринг страницы.")
    return render(request, 'app1/list.html', context)