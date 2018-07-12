import os
from os.path import isfile, join

import pdf_read as pr
import docx_read as dr


#======================================================================


PATH = 'C:/Users/AFitzpatrick/Desktop/Programming/testing/'
filesDict = {} #dictionary of the form {filename : file extension}


#======================================================================


def main():

	organise_files()
	delete_duplicates()
	
	
#======================================================================


def organise_files():
	global PATH
	onlyFiles = sorted([f for f in os.listdir(PATH) if isfile(join(PATH, f))])
	for filename in onlyFiles:
		if '.docx' in filename:
			filesDict[filename] = '.docx'
		elif '.pdf' in filename:
			filesDict[filename] = '.pdf'
		
		
def find_email_address(filename):
	'''
	FIND EMAIL ADDRESSES CONTAINED IN FILES	
	'''
	global PATH
	if filesDict[filename] == '.docx':
		text = dr.docx_to_text(PATH+filename)
	elif filesDict[filename] == '.pdf':
		text = pr.pdf_to_text(PATH+filename)
		
	emailAddress = ''
	for word in text.split():
		if '@' in word:
			emailAddress = word
	return emailAddress
				

def delete_duplicates():
	'''
	ITERATE THROUGH AND DELETE CVS
	'''
	emailAddressFirst = 'genesis'
	for file in sorted(filesDict):
		emailAddressSecond = find_email_address(file)
		if emailAddressFirst == emailAddressSecond or emailAddressSecond == '':
			os.remove(PATH+file)
			print(file + ' deleted.')
		emailAddressFirst = emailAddressSecond
	

#======================================================================


if __name__ == '__main__':
	main()