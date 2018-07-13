import os
from os.path import isfile, join

import pdf_read as pr
import docx_read as dxr
import doc_read as dr


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
		if '.docx' in filename[-5:]:
			filesDict[filename] = '.docx'
		if '.doc' in filename[-4:]:
			filesDict[filename] = '.doc'
		elif '.pdf' in filename[-4:]:
			filesDict[filename] = '.pdf'
		print(info,'\n')
		
		
def find_email_address(filename):
	'''
	FIND EMAIL ADDRESSES CONTAINED IN FILES	
	'''
	global PATH
	emailAddress = ''
	if filesDict[filename] == '.docx':
		text = dxr.docx_to_text(PATH+filename)
	if filesDict[filename] == '.doc':
		text = dr.doc_to_text(PATH+filename)
	elif filesDict[filename] == '.pdf':
		text = pr.pdf_to_text(PATH+filename)
	if text == 'error':
		print('error -',filename)
		emailAddress = 'error'
		return emailAddress
		
	for word in text.split():
		if '@' in word:
			emailAddress = word
	return emailAddress
				

def delete_duplicates():
	'''
	ITERATE THROUGH AND DELETE CVS
	'''
	global PATH
	emailAddressFirst = 'genesis'
	firstOccurence = 'first'
	uniqueCount = 70
	
	for file in sorted(filesDict):
		try:
			emailAddressSecond = find_email_address(file)
		except:
			print('Error encountered when processing ' + file + '. Moved to exceptions folder.')
			emailAddressSecond = 'error'
			os.rename(PATH+file, PATH+'caught exceptions/'+file)
			
		if emailAddressFirst == emailAddressSecond or emailAddressSecond == '':
			if emailAddressSecond != 'error':
				os.remove(PATH+file)
			print(uniqueCount,file + ' deleted.')
			
			if emailAddressSecond != firstOccurence:
				uniqueCount += 1
			firstOccurence = emailAddressSecond
			
		emailAddressFirst = emailAddressSecond
	print('Number of unique deletions:',uniqueCount)
	

#======================================================================


if __name__ == '__main__':
	main()
	
