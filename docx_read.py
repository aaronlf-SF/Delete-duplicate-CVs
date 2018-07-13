import docx2txt

def docx_to_text(path):
	try:
		text = docx2txt.process(path)
	except:
		text = 'error'
	return text