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
		elif '.doc' in filename[-4:]:
			filesDict[filename] = '.doc'
		elif '.pdf' in filename[-4:]:
			filesDict[filename] = '.pdf'
		
		

def extract_text(filename):
	global PATH
	if filesDict[filename] == '.docx':
		text = dxr.docx_to_text(PATH+filename)
	if filesDict[filename] == '.doc':
		text = dr.doc_to_text(PATH+filename)
	elif filesDict[filename] == '.pdf':
		text = pr.pdf_to_text(PATH+filename)
	return text	
		
		
def find_email_address(filename,text):
	'''
	FIND EMAIL ADDRESSES CONTAINED IN FILES	
	'''
	emailAddress = ''
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
	uniqueCount = 0
	filenames_and_emails = {}
	
	for file in sorted(filesDict):
		filesDeleted = False
		try:
			text = extract_text(file)
			emailAddress = find_email_address(file,text)
			if emailAddress == '':
				#remove files with no email addresses
				os.remove(PATH+file)
				print(file + ' deleted.')
				uniqueCount += 1
			else:
				if len(list(filenames_and_emails.keys())) > 0: #prevents call to an empty dictionary
					if emailAddress not in list(filenames_and_emails.values())[0]['email']:
						print(emailAddress)
						#find latest year
						years = []
						for i in filenames_and_emails:
							years.append(filenames_and_emails[i]['latestYear'])
						maxYear = max(years)
						
						#delete all but that one in that year
						safeFile = [file for file in filenames_and_emails if filenames_and_emails[file]['latestYear'] == maxYear][0]
						print('###',safeFile)
						for file in filenames_and_emails:
							if file != safeFile:
								os.remove(PATH+file)
								print(file + ' deleted.')
								filesDeleted = True
			
						filenames_and_emails = {}
						print(filenames_and_emails)
						if filesDeleted == True:
							uniqueCount += 1
						
				latestYear = find_latest_year(text)
				filenames_and_emails[file] = {'email':emailAddress, 'latestYear':latestYear}

				
		except:
			print('Error encountered when processing ' + file + '. Moved to exceptions folder.')
			emailAddress = 'error'
			os.rename(PATH+file, PATH+'caught exceptions/'+file)
			
	print('Number of unique deletions:',uniqueCount)
	

#======================================================================


def find_latest_year(text):
	years = {
			'1990':1990,
			'1991':1991,
			'1992':1992,
			'1993':1993,
			'1994':1994,
			'1995':1995,
			'1996':1996,
			'1997':1997,
			'1998':1998,
			'1999':1999,
			'2000':2000,
			'2001':2001,
			'2002':2002,
			'2003':2003,
			'2004':2004,
			'2005':2005,
			'2006':2006,
			'2007':2007,
			'2008':2008,
			'2009':2009,
			'2010':2010,
			'2011':2011,
			'2012':2012,
			'2013':2013,
			'2014':2014,
			'2015':2015,
			'2016':2016,
			'2017':2017,
			'2018':2018
			}

	years_found = []
	for word in text.split():
		for year in list(years.keys()):
			if word == year:
				years_found.append(years[year])
	if len(years_found) > 0:			
		latestYear = max(years_found)
	else:
		latestYear = 0
	return latestYear

	
#======================================================================	
	
	
if __name__ == '__main__':
	main()
	
