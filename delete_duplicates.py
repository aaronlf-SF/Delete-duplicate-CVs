import os
from os.path import isfile, join

import threading
import time
import pandas as pd

import pythoncom

import pdf_read as pr
import docx_read as dxr
import doc_read as dr



#======================================================================



PATH = "C:/Users/AFitzpatrick/Desktop/Programming/testing/[CONSULTANT NAME]/"
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
		if '.docx' in filename[-5:] or '.DOCX' in filename[-5:]:
			filesDict[filename] = '.docx' 
		elif '.doc' in filename[-4:] or '.DOC' in filename[-4:]:
			filesDict[filename] = '.doc'
		elif '.pdf' in filename[-4:] or '.PDF' in filename[-4:]:
			filesDict[filename] = '.pdf'
		else:
			print(filename,'not organised')
	
	
		
#======================================================================



def find_email_address(filename,text):
	'''
	FIND EMAIL ADDRESSES CONTAINED IN DOCUMENT BY
	PARSING FOR '@' SYMBOL. 
	'''
	emailAddress = ''
	if text == 'error':
		print('error -',filename)
		emailAddress = 'error'
		return emailAddress
		
	for word in text.split():
		if '@' in word:
			emailAddress = word.lower()
			break
	return emailAddress


def extract_text(filename):
	'''
	RETURNS THE RAW TEXT FROM THE DOCUMENT
	'''
	global PATH
	if filesDict[filename] == '.docx':
		text = dxr.docx_to_text(PATH+filename)
	if filesDict[filename] == '.doc':
		text = dr.doc_to_text(PATH+filename)
	elif filesDict[filename] == '.pdf':
		text = pr.pdf_to_text(PATH+filename)
	return text	
	
	
def extract_text_on_time(file):
	'''
	NECESSARY DEPENDENT FUNCTION FOR EXTRACTING TEXT ON A SEPARATE THREAD
	'''
	pythoncom.CoInitialize()
	global textFound,output_text
	output_text = extract_text(file)
	textFound = True

		
def getText(file):
	'''
	RETRIEVES THE TEXT ON A SEPARATE THREAD WITH A TIMEOUT OF 45 SECONDS
	TO ACCOUNT FOR EXCEPTIONS THE PROGRAM CAN'T CATCH
	'''
	global textFound,output_text
	textFound = False
		
	thread = threading.Thread(target=extract_text_on_time,args=(file,),daemon=True)
	thread.start()
	thread.join(timeout=45)
	
	if textFound == False:
		output_text = 'error'
	return output_text

	
	
#======================================================================



def delete_duplicates():
	'''
	ITERATE THROUGH AND DELETE CVS
	'''
	global PATH,totalCount,uniqueCount
	uniqueCount = 0
	totalCount = 0
	filenames_and_emails = {}
	
	for file in sorted(filesDict):
		print(uniqueCount,file)
		filesDeleted = False
		try:
			text = getText(file)
			emailAddress = find_email_address(file,text)
			if emailAddress == '':
				#remove files with no email addresses
				os.remove(PATH+file)
				print(uniqueCount,file + ' deleted.')
				uniqueCount += 1
				totalCount += 1
				
			elif emailAddress == 'error':
				#remove email addresses with errors
				print('Error encountered when processing ' + file + '. Removing.')
				os.remove(PATH+file)
				uniqueCount += 1
				totalCount += 1
			else:
				compare_and_remove_files(file,emailAddress,filenames_and_emails,text)
				latestYear = find_latest_year(text)
				filenames_and_emails[file] = {'email':emailAddress, 'latestYear':latestYear}

		except Exception as e:
			print(e)
			print('Error encountered when processing ' + file + '. Note made in bad eggs.txt.')
			with open(PATH + 'bad eggs.txt','a') as txt:
				txt.write(file + '\n')
			
	print('Number of unique deletions:',uniqueCount)
	print('Total number of deletions:',totalCount)

	
	
#======================================================================



def compare_and_remove_files(file,emailAddress,filenames_and_emails,text):
	'''
	CHECKS TO SEE IF A DOCUMENT SHOULD BE DELETED BASED ON WHETHER OR NOT 
	ANOTHER FILE SHARES THE SAME EMAIL ADDRESS STRING, OR CONTAINS NO
	EMAIL ADDRESS AT ALL
	'''
	latestYear = find_latest_year(text)
	if file == sorted(filesDict)[-1]: # LAST FILE IN THE DIRECTORY
		filenames_and_emails[file] = {'email':emailAddress, 'latestYear':latestYear}
						
	if len(list(filenames_and_emails.keys())) > 0: #prevents call to an empty dictionary
		if emailAddress not in list(filenames_and_emails.values())[0]['email'] or file == sorted(filesDict)[-1]:
			maxYear = get_max_year(filenames_and_emails)
			safeFile = keep_safe_file(filenames_and_emails,maxYear)
			email = list(filenames_and_emails.values())[0]['email']
			filenames_and_emails.clear()
				
			safeDict = {'filename':[safeFile],'email':[email],'latestYear':[maxYear]}
			compare_against_dataset_file(safeDict)
	
	else:
		fileDict = {'filename':[file],'email':[emailAddress],'latestYear':[latestYear]}
		compare_against_dataset_file(fileDict)			
					
					
def get_max_year(filenames_and_emails):
	years = []
	for i in filenames_and_emails:
		years.append(filenames_and_emails[i]['latestYear'])
	return max(years)
													
							
def keep_safe_file(filenames_and_emails,maxYear):
	'''
	DELETE ALL FILES BUT ONE 'SAFE FILE' FROM THAT YEAR
	'''
	global totalCount, uniqueCount
	duplicate_files = [f for f in filenames_and_emails if filenames_and_emails[f]['latestYear'] == maxYear]
	safeFile = duplicate_files[0]
	initial_totalCount = totalCount
	
	for filename in filenames_and_emails:
		if filename != safeFile:
			os.remove(PATH+filename)
			print(uniqueCount,filename + ' deleted.')
			totalCount += 1
	if totalCount != initial_totalCount:
		uniqueCount += 1
		
	return safeFile
	
	

#======================================================================
	
	
	
def compare_against_dataset_file(safeDict):
	'''
	CHECK THE EXISTING DATAFRAME OF FILES AND EMAIL ADDRESSES
	TO SEE IF IT HAS BEEN FOUND ALREADY, THEN DEAL WITH THE RESULT
	ACCORDINGLY
	'''
	global PATH, uniqueCount,totalCount
	try:
		df = csv_to_df()
	except FileNotFoundError:
		df = pd.DataFrame.from_dict(safeDict)
		df = df.drop(df.index[0])
	else:
		pandasExists = False
		for i in range(df.count(0)['filename']):
			if df.iloc[i]['email'] == safeDict['email'][0]:
				pdIndex = i
				pandasExists = True
				
		if pandasExists == True:
			pdLatestYear = df.iloc[pdIndex]['latestYear']
			maxYear = max([pdLatestYear,safeDict['latestYear'][0]])
			if maxYear ==  safeDict['latestYear'][0]:
				print(df.iloc[pdIndex]['filename'],'deleted and dropped from dataframe.')
				os.remove(PATH+df.iloc[pdIndex]['filename'])
				df.drop(df.index[pdIndex])
				new_df = pd.DataFrame.from_dict(safeDict)
				df = df.append(new_df,ignore_index=True)
			else:
				os.remove(PATH + safeDict['filename'][0])
				print(safeDict['filename'][0],'deleted. (Newer version in dataframe)')
			uniqueCount += 1
			totalCount += 1
		else:
			new_df = pd.DataFrame.from_dict(safeDict)
			df = df.append(new_df,ignore_index=True)
	df_to_csv(df)
	
	
		
def df_to_csv(df):
	global PATH
	df.to_csv(PATH+'data.csv',index=False)

	
def csv_to_df():
	global PATH
	df = pd.read_csv(PATH+'data.csv')
	return df

	
	
#======================================================================



def find_latest_year(text):
	'''
	LOOKS THROUGH THE TEXT TO FIND YEAR KEYWORDS TO DETERMINE
	HOW RECENT THE FILE IS. LOOKS FOR YEARS FROM 1990 - 2018
	'''
	years = {}
	
	for year in range(1990,2019):
		years[year] = [
						str(year),  #e.g. 2012
						"'" + str(year)[2:], # e.g. '12
						str(year)[2:]  # e.g. 12
						]

	years_found = []
	for word in text.split():
		for year in years:
			for year_format in years[year][:2]:
				if word == year_format:
					years_found.append(year)
	
	for word in ( text.split('/') + text.split('-') + text.split('.') ):
		for year_format in years[year]:
				if word == year_format:
					years_found.append(year)

	if len(years_found) > 0:			
		latestYear = max(years_found)
	else:
		latestYear = 0
	return latestYear

	
	
#======================================================================	
	
	
	
if __name__ == '__main__':
	main()
	

	