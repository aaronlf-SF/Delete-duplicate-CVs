import win32com.client
import os
from os.path import isfile, join


def convert_backslash(string):
	return string.replace('/','\\')
	
	
def doc_to_text(fullPath):
	fullPath = convert_backslash(fullPath)
	app = win32com.client.Dispatch('Word.Application')
	doc = app.Documents.Open(fullPath)
	output_text = doc.Content.Text
	app.Quit()
	print(output_text)
	return output_text
	
