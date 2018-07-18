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


PATH = 'C:/Users/AFitzpatrick/Desktop/Programming/testing/[CONSULTANT NAME]/'
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
	FIND EMAIL ADDRESSES CONTAINED IN FILES	
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
	global PATH
	if filesDict[filename] == '.docx':
		text = dxr.docx_to_text(PATH+filename)
	if filesDict[filename] == '.doc':
		text = dr.doc_to_text(PATH+filename)
	elif filesDict[filename] == '.pdf':
		text = pr.pdf_to_text(PATH+filename)
	#print(text,'\n',filename)
	return text	
	
	
def extract_text_on_time(file):
	pythoncom.CoInitialize()
	global textFound,output_text
	output_text = extract_text(file)
	textFound = True

		
def getText(file):
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
				print('Error encountered when processing ' + file + 'Removing.')
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
		'''if file == sorted(filesDict)[-1]:
			fileDict = {'filename':[file],'email':[emailAddress],'latestYear':[latestYear]}
			compare_against_dataset_file(fileDict)'''	
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
	delete all but that one in that year
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
	years = {
			1990:['1990',"'90",'90'],
			1991:['1991',"'91",'91'],
			1992:['1992',"'92",'92'],
			1993:['1993',"'93",'93'],
			1994:['1994',"'94",'94'],
			1995:['1995',"'95",'95'],
			1996:['1996',"'96",'96'],
			1997:['1997',"'97",'97'],
			1998:['1998',"'98",'98'],
			1999:['1999',"'99",'99'],
			2000:['2000',"'00",'00'],
			2001:['2001',"'01",'01'],
			2002:['2002',"'02",'02'],
			2003:['2003',"'03",'03'],
			2004:['2004',"'04",'04'],
			2005:['2005',"'05",'05'],
			2006:['2006',"'06",'06'],
			2007:['2007',"'07",'07'],
			2008:['2008',"'08",'08'],
			2009:['2009',"'09",'09'],
			2010:['2010',"'10",'10'],
			2011:['2011',"'11",'11'],
			2012:['2012',"'12",'12'],
			2013:['2013',"'13",'13'],
			2014:['2014',"'14",'14'],
			2015:['2015',"'15",'15'],
			2016:['2016',"'16",'16'],
			2017:['2017',"'17",'17'],
			2018:['2018',"'18",'18'],
			}

	years_found = []
	for word in text.split():
		for year in list(years.keys()):
			if word in years[year][:2]:
				years_found.append(year)
	for word in text.split('/'):
		for year in list(years.keys()):
			if word in years[year]:
				years_found.append(year)
	for word in text.split('-'):
		for year in list(years.keys()):
			if word in years[year]:
				years_found.append(year)
	for word in text.split('.'):
		for year in list(years.keys()):
			if word in years[year]:
				years_found.append(year)
	if len(years_found) > 0:			
		latestYear = max(years_found)
	else:
		latestYear = 0
	return latestYear

	
#======================================================================	
	
	
if __name__ == '__main__':
	main()
	
