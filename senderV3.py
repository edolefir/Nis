 # -*- coding: utf-8 -*-

def toUTF8(value):
	"""Convert value to UTF-8 string"""
	if type(value) == unicode:
		return value.encode('utf-8')
	if type(value) == str:
		return value
	return str(value)

THEME = "Invitation to attend ICPNS'2019" #input
#ATTACHES = ["CFP.pdf"]
ATTACH_TPL = "Invitation.docx"  
ATTACH_MASK = ("Invitation_", ".docx")
MAIL_SERVER = 'smtp.gmail.com:587'
LOGIN = '' #input from FROM_EMAIL
PWD = '' #input
FROM_EMAIL = '@miem.hse.ru' #input
	
import smtplib
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email import Encoders

from docxtpl import DocxTemplate

def changeName(name):
	return name.replace('.', '_').replace(' ', '').replace('-', '_')
	
def docx2pdf(input):
	import win32com.client as client
	from os import path
	try:
		word = client.DispatchEx("Word.Application")
		target_path = path.abspath(input.replace(".docx", r".pdf"))
		word_doc = word.Documents.Open(path.abspath(input))
		word_doc.SaveAs(target_path, FileFormat=17)
		word_doc.Close()
	except Exception as e:
		print("Error: " + input + " - " + str(e))
	finally:
		word.Quit()
	

def generateInvitation(data):
	# generate doc file
	doc = DocxTemplate(ATTACH_TPL)
	doc.render(data)
	resFileName = ATTACH_MASK[0] + changeName(data['Name']) + '_' + ''.join([i[0] for i in data['Title'].split(' ')]) + ATTACH_MASK[1]
	doc.save(resFileName)
	
	# save doc as pdf
	docx2pdf(resFileName)
	
	#return resFileName
	
def generateInvitations(allData):
	for key in allData:
		d = {'Name': allData[key][0], 'Title': allData[key][1]}
		try:
			generateInvitation(d)
		except Exception as e:
			print("Error: " + str(d) + " - " + str(e))


def emailHTMLAttachments(toEmailList, textMessage = '', attachments = []):
	COMMASPACE = ', '
	
	# Create the container (outer) email message.
	msg = MIMEMultipart()
	msg['Subject'] = THEME
	msg['From'] = FROM_EMAIL
	msg['To'] = COMMASPACE.join(toEmailList)
	text = textMessage
	
	body = MIMEMultipart('alternative')
	# part1 = MIMEText(text, 'plain')
	part1 = MIMEText(text, 'html')
	body.attach(part1)
	msg.attach(body)
	for attachment in attachments:
		if attachment.find('doc') > 0:
			attachFile = MIMEBase('application', 'msword')
		elif attachment.find('pdf') > 0:
			attachFile = MIMEBase('application', 'pdf')
		else:
			attachFile = MIMEBase('application', 'octet-stream')
		
		fo = open(attachment,'rb')
		attachFile.set_payload(fo.read())
		fo.close()
		
		Encoders.encode_base64(attachFile)
		attachFile.add_header('Content-Disposition', 'attachment', filename=attachment)
		
		msg.attach(attachFile)
"""
	server = smtplib.SMTP(MAIL_SERVER)  
	server.starttls()  
	server.login(LOGIN, PWD)  
	server.sendmail(FROM_EMAIL, toEmailList, msg.as_string())
	server.close()
"""


CONTACT_FILE = 'maillistRedo.xlsx'
LETTER_FILE = 'letter4.txt'
	
import os.path
import datetime
import pandas
import time

mailDict = {}

worksheet = pandas.read_excel(CONTACT_FILE, sheet_name=0)
for i in worksheet.index:
	if pandas.isnull(worksheet[u'Name'][i]):
		mailDict[toUTF8(worksheet[u'Email'][i])] = ("colleague", "tittle")
	else:
		mailDict[toUTF8(worksheet[u'Email'][i])] = (toUTF8(worksheet[u'Name'][i]), toUTF8(worksheet[u'Title'][i]))


# generateInvitations(mailDict)

letterText = ''
with open(LETTER_FILE, 'r') as content_file:
	letterText = content_file.read()
	
dictSize = len(mailDict)
while ( dictSize > 0):
	successEmailList = []
	failedEmailList = []
	for mail in mailDict:
		try:
			# file name form rule
			attach = ATTACH_MASK[0] + changeName(mailDict[mail][0]) + '_' + ''.join([i[0] for i in mailDict[mail][1].split(' ')]) + '.pdf'
			emailHTMLAttachments([mail], letterText.replace("__Name__", mailDict[mail][0]).replace("__Title__", mailDict[mail][1]), [attach])
			successEmailList.append(mail)
		except Exception as e:
			failedEmailList.append(mail + ' - ' + str(e))
		time.sleep(3)
	map(mailDict.pop, successEmailList)
	
	with open(os.path.join("success", datetime.datetime.now().strftime("%m_%d_%H_%M_%S") + ".log"), "w") as f: 
		f.write("\n".join(successEmailList))
		
	with open(os.path.join("fail", datetime.datetime.now().strftime("%m_%d_%H_%M_%S") + ".log"), "w") as f: 
		f.write("\n".join(failedEmailList))
		
	dictSize -= len(successEmailList)
