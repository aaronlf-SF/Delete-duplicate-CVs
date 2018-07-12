import docx2txt

def docx_to_text(path):
	text = docx2txt.process(path)
	return text