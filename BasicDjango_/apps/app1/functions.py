def changeName(name):
	Mask = name.split('.')
	return name.replace('.', '_').replace(' ', '').replace('-', '_')

import pandas

def CreateDict(data):
	Keys = data.columns
	for i in data.index:
		if pandas.isnull(data['Email'][i]) == False:

			for key in Keys:
				if pandas.isnull(data[key][i]):

		
