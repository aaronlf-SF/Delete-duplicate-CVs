import win32com.client
import os
from os.path import isfile, join

def convert_backslash(string):
	return string.replace('/','\\')
	
	
def doc_to_text(fullPath):
	try:
		fullPath = convert_backslash(fullPath)
		app = win32com.client.DispatchEx('Word.Application')
		doc = app.Documents.Open(fullPath)
		output_text = doc.Content.Text
		app.Quit()
	except:
		output_text = 'error'
	return output_text
	
