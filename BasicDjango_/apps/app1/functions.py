def changeName(name):
	Mask = name.split('.')
	return name.replace('.', '_').replace(' ', '').replace('-', '_')

import pandas

def CreateData(data):
	Keys = data.columns
	mailsDict = {}
	for i in data.index:
		if pandas.isnull(data['Email'][i]) == False:
			keyEmail = data['Email'][i]
			mailDict = {}
			for key in Keys:
				if pandas.isnull(data[key][i]):
					mailDict.update({key: key + 'default'})
				else:
					mailDict.update({key: data[key][i]})
			mailsDict.update({keyEmail: mailDict})
	return mailsDict

		
