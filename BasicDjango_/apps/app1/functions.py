def changeName(name):
	Mask = name.split('.')
	return name.replace('.', '_').replace(' ', '').replace('-', '_')

import pandas

def CreateKeys(data):
	 return data.columns
