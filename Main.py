#!/usr/bin/python
#=============================================
# MIP Report Utility for Release note 
#
# Author: Amogh Jagadale
#
# Date: 2017-11-15
#
# Version: 1.0
#
#=============================================

#***************
# Imported modules 
#***************
import sys
import os
import re
from collections import OrderedDict
import xlsxwriter
import csv
import shutil
import logging
import datetime
#import SummaryPage

#***************
# Loging Properties
#***************
#Create logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Create console handler and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# Create formatter
formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s',datefmt='%m/%d/%Y %I:%M:%S %p')

# Add formatter to ch
ch.setFormatter(formatter)

# Add ch to logger
logger.addHandler(ch)

#***************
# Functions
#***************
def MIPSizeFinder(value):
	table = (csv.reader(value.splitlines(),delimiter=';'))
	for i in table: 

		MIPSizeUnit = (re.findall("(\D+)", i[1].strip())[0]).strip()
		MIPSize = int(re.findall("(\d+)", i[1].strip())[0])

	return (MIPSize,MIPSizeUnit)

def AddWorkSheet(Dict,FileName,CSVFilePath,workbook):
	#print SummaryDict
	try:

		#print FileName	
		#Intializing row and col value
		row = 0
		col = 1

		#Default Stylying for Formatting
		BorederStyle = 1
		MergeCellFormat = {'align': 'center','valign': 'vcenter','border': BorederStyle}
		DefaultCellFormat = {'bg_color': '#FFFFFF','border':BorederStyle}
		DefaultHeaderCellFormat = {'bg_color': '#EEECE1','border':BorederStyle}

		#Column widths
		FirstColWidth = 52.56
		SecondToFourthColWidth = 32.25
		LastColWidth = 40.22
		
		WorkSheetNum = re.findall("_UR(\d+)", CSVFilePath)[0]

		worksheet = workbook.add_worksheet(WorkSheetNum)

		SummaryValDict = {}
		SummaryValDict = OrderedDict()
		for index,v in enumerate(Dict):
			
			value = Dict[v]

			if index == 1:
				(MIPSize,MIPSizeUnit) = MIPSizeFinder(value)
				
			#Converting data to list and spliting on the basis of semicolon
			table = (csv.reader(value.splitlines(),delimiter=';'))

			row +=1
			#print table
			for i,v in enumerate(table):
				#Setting worksheet width
				worksheet.set_column('C:E', SecondToFourthColWidth)
				worksheet.set_column('B:B', FirstColWidth)
				worksheet.set_column('F:F', LastColWidth)

				#Condition for selecting the first row
				if i == 0:
					SummaryValDict['Comments'] = ''
					#print v

					#Making the text of the header in center for each table
					format = workbook.add_format({'bg_color': '#EEECE1','border':1,'align': 'center','valign': 'vcenter'})
					
					if (v[0].lower()).strip() == 'number of files in old-new-mip same?':
						#print v
						
						CellFormat = workbook.add_format(MergeCellFormat)

						CellsToMerge = 4
						MergeStart = chr(col+97).upper() + str(row+1)
						MergeEnd = chr(col+97+CellsToMerge).upper() + str(row+1)
						MergeCell = MergeStart+':'+MergeEnd
						worksheet.merge_range(MergeCell,"", CellFormat)
					
					#MIP Size table Formatting
					elif (v[0].lower()).strip() == 'mip size':
						format = workbook.add_format({'bg_color': '#FFFFFF','border':1})
					

					worksheet.write_row(row,col,v,format)

				else:
					
					#Setting Default format for all rows except the first row
					format = workbook.add_format(DefaultCellFormat)
					if (v[0].lower()).strip() == 'update region':
						UpdateRegion = v[2].strip()
						SummaryValDict['UpdateRegion'] = UpdateRegion
						worksheet.write_row(row,col,v,format)
						#v = v
						#print UpdateRegion

					elif (v[0].lower()).strip() == 'versionid':
						OldVersionID = v[2].strip()
						NewVersionID = v[1].strip()
						SummaryValDict['OldVersionID'] = OldVersionID
						SummaryValDict['NewVersionID'] = NewVersionID
						worksheet.write_row(row,col,v,format)
						#print v	
					#MIP Report Size Validation
					elif (v[0].lower()).strip() == 'mip less than 500 mb':
						SummaryValDict['MIPSize'] = str(MIPSize) + ' ' + MIPSizeUnit
						if MIPSize > 500 and MIPSizeUnit.lower() != 'kb':
							v[1] = ' FAILED'
							MIPFileSizeFailure = True
							SummaryValDict['MIPVerfication'] = 'Failed'
							SummaryValDict['Recomended'] = 'NO'

						elif MIPSize >= 1024 and MIPSizeUnit.lower() == 'kb':
							v[1] = ' FAILED'
							MIPFileSizeFailure = True
							SummaryValDict['MIPVerfication'] = 'Failed'
							SummaryValDict['Recomended'] = 'NO'

						else:
							MIPFileSizeFailure = False
							SummaryValDict['MIPVerfication'] = 'Successful'
							SummaryValDict['Recomended'] = 'YES'
						
						worksheet.write_row(row,col,v,format)

					#Good to go validation of the basis of MIP Report size validation
					elif (v[0].lower()).strip() == 'good to go':

						if MIPFileSizeFailure == True:
							format = workbook.add_format({'bg_color': '#FF0000','border':1,'align': 'center','valign': 'vcenter'})
							v[1] = ' NO'

						elif (v[1].lower()).strip() == 'yes' and MIPFileSizeFailure == False:
							format = workbook.add_format({'bg_color': '#92D050','border':1,'align': 'center','valign': 'vcenter'})
							
						else:
							format = workbook.add_format({'bg_color': '#FF0000','border':1,'align': 'center','valign': 'vcenter'})
							v[1] = ' NO'
						
						worksheet.write_row(row,col,v,format)

					# Making neccesary formatting changes in next row of 'Number of files in Old-new-MIP same?'
					elif (v[0].lower()).strip() == 'folder' and (v[1].lower()).strip() == 'old count':
						format = workbook.add_format({'bg_color': '#EEECE1','border':1,'align': 'center','valign': 'vcenter'})

						worksheet.write_row(row,col,v,format)

					#print MIPSize
					else:

						format = workbook.add_format(DefaultCellFormat)
						worksheet.write_row(row,col,v,format)
				#print MIPSize
				#SummaryDict.append(SummaryList)
				row +=1
				#print MIPSize
		
		#print SummaryList
		return (True,WorkSheetNum,SummaryValDict)
	except:
		return (False,WorkSheetNum,SummaryValDict)

def SummaryPageCreator(SummaryDict,SummaryWorkSheet,workbook):
	FirstColWidth = 40.78
	SecondColWidth = 29.67
	ThirdColWidth = 38.11
	FourthColWidth = 29.89
	FifthColWidth = 19.78
	SixthColWidth = 17.78
	SeventhColWidth = 28.78

	SummaryWorkSheet.set_column('B:B', FirstColWidth)
	SummaryWorkSheet.set_column('C:C', SecondColWidth)
	SummaryWorkSheet.set_column('D:D', ThirdColWidth)
	SummaryWorkSheet.set_column('E:E', FourthColWidth)
	SummaryWorkSheet.set_column('F:F', FifthColWidth)
	SummaryWorkSheet.set_column('G:G', SixthColWidth)
	SummaryWorkSheet.set_column('H:H', SeventhColWidth)
	SummaryWorkSheet.set_row(9, 28.80)

	#PaleBrown = {'bg_color': '#EEECE1','border':1}
	

	def WriteTable1(SummaryDict,workbook):
		row = 1
		col = 3

		for key,val in SummaryDict.items():
			if key != 'Table2' and key != 'Table3':
				#Column 1 in table 1
				format = workbook.add_format({'bg_color': '#EEECE1','border':1})
				SummaryWorkSheet.write(row,col,key,format)
				#Column 2 in table 1
				format = workbook.add_format({'border':1,'align': 'center','valign': 'vcenter'})
				SummaryWorkSheet.write(row,col+1,val,format)
				row += 1

	def WriteTable2(SummaryDict,workbook):
		row = 9
		col = 1

		Header = ['Update Region','Old Version','New Version','MIP Size','Verification Except MIP size','Recommended  to Deliver','Comments']
		format = workbook.add_format({'bg_color': '#EEECE1','border':1,'align': 'center','valign': 'vcenter'})
		format.set_text_wrap()
		SummaryWorkSheet.write_row(row,col,Header,format)

		row += 1
		for key,val in SummaryDict.items():
			if key == 'Table2':
				format = workbook.add_format({'border':1})
				Greenformat = workbook.add_format({'border':1,'bg_color': '#92D050'})
				Redformat = workbook.add_format({'border':1,'bg_color': '#FF0000'})
				for i in val:
					SummaryWorkSheet.write(row,col,i['UpdateRegion'],format)
					SummaryWorkSheet.write(row,col+1,i['OldVersionID'],format)
					SummaryWorkSheet.write(row,col+2,i['NewVersionID'],format)
					if i['Recomended'].lower() == 'no':
						SummaryWorkSheet.write(row,col+3,i['MIPSize'],Redformat)
						SummaryWorkSheet.write(row,col+4,i['MIPVerfication'],Redformat)
						SummaryWorkSheet.write(row,col+5,i['Recomended'],Redformat)
						SummaryWorkSheet.write(row,col+6,'MIP Size > 500 MB',format)
					else:
						SummaryWorkSheet.write(row,col+3,i['MIPSize'],Greenformat)
						SummaryWorkSheet.write(row,col+4,i['MIPVerfication'],Greenformat)
						SummaryWorkSheet.write(row,col+5,i['Recomended'],Greenformat)
						SummaryWorkSheet.write(row,col+6,i['Comments'],format)
					row += 1
		return row
				#print val['UpdateRegion']

	def WriteTable3(SummaryDict,workbook,Row):
		row = Row + 2
		col = 1

		Header = ['MIP File name','Over all file size in bytes','Sha256 sum Hex value']
		format = workbook.add_format({'bg_color': '#EEECE1','border':1,'align': 'center','valign': 'vcenter'})
		
		def FindMergeCells():
			CellsToMerge = 2
			MergeStart = chr(col+2+97).upper() + str(row+1)
			MergeEnd = chr(col+2+97+CellsToMerge).upper() + str(row+1)
			MergeCell = MergeStart+':'+MergeEnd
			return MergeCell

		#x = 2
		MergeCell = FindMergeCells()
		SummaryWorkSheet.merge_range(MergeCell,"",format)

		SummaryWorkSheet.write_row(row,col,Header,format)
		row += 1
		format = workbook.add_format({'border':1})
		for key,val in SummaryDict.items():
			if key == 'Table3':
				for i in val:
					MergeCell = FindMergeCells()
					SummaryWorkSheet.merge_range(MergeCell,"",format)
					SummaryWorkSheet.write_row(row,col,i,format)
					row += 1		
	try:
		WriteTable1(SummaryDict,workbook)
		Row = WriteTable2(SummaryDict,workbook)
		WriteTable3(SummaryDict,workbook,Row)
		return True
	except:
		return False

def DictMaker(CSVFile):
	# Each CSV file is converted into dict in the form
	'''
	{'data0':"Value from paragraph1 of CSV",'data1':"Value from paragraph2 of CSV",.......,'dataN':"Value from paragraphN of CSV"}
	'''
	Dict = {}
	Dict = OrderedDict()

	with open(CSVFile) as Data:
		ReqData = ""
		ReqDataNum = 0
		for Line in Data:
			if Line.strip() != "; ; ;":
				ReqData = ReqData + Line
			else:
				Dict["Data" + str(ReqDataNum)] = ReqData
				ReqDataNum += 1
				ReqData = ""
		Data.close()

	return Dict

def Recomendation(Summary):
	RecomendationCount = 0
	TotalCount = 0
	for x in Summary:
		TotalCount += 1
		if x['MIPVerfication'] == 'Successful':
			RecomendationCount += 1
		
	return (RecomendationCount,TotalCount)


def GetDataFromShaFile(SHAFilePath,SummaryDict):
	VersionName = [i['NewVersionID'] for i in SummaryDict['Table2']]
	with open(SHAFilePath) as Data:
		#print filter(lambda x: x.lower() not in 'done',Data)
		SHAFileData = [i for i in Data if 'done' not in i.lower()] 
		#print [i if "done" not in i.lower() else "" for i in Data]
		# for i in Data:
		# 	if "done" not in i.lower():
		# 		print i
	Data.close()
	Table3Data = []
	#print VersionName.lower()
	for i in VersionName:
		for j in SHAFileData:
			if i.lower() in j.lower():
				table = (csv.reader(j.splitlines(),delimiter=';'))
				#del table[1-2]
				for x in table:
					#Alternative for currently used logic
					#del x[1]
					#del x[1]
					#Table3Data.append(x)
					Table3Data.append([y for y in [a for a in x if '/' not in a]  if ':' not in y])
	return Table3Data

def CreateWorkBook(FileNameWithPath,ListOfCSVFiles,FileName,RegionVar,FromVer,ToVer):
	SummaryDict = {}
	SummaryDict = OrderedDict()
	SummaryDict['Market'] = RegionVar
	SummaryDict['From Version'] = FromVer
	SummaryDict['To Version'] = ToVer
	SummaryDict['Report Created By'] = ReportCreator
	SummaryDict['Date of Report'] = str(datetime.date.today())
	

	#Creating xlsx workbook
	try:
		#print FromVer,ToVer
		workbook = xlsxwriter.Workbook(FileNameWithPath)
		SummaryWorkSheet = workbook.add_worksheet('Summary')
		SummaryList = []
		for CSVFile in ListOfCSVFiles:
			Dict = DictMaker(CSVFile)
			(WorkSheetCreated,WorkSheetNum,SummaryValDict) = AddWorkSheet(Dict,FileName,CSVFile,workbook)
			SummaryList.append(SummaryValDict)
			if WorkSheetCreated == True:
				logger.info("Worksheet %s created in %s Workbook."%(WorkSheetNum,FileName))
			else:
				logger.error("Error encountered while creating Worksheet %s in %s"%(WorkSheetNum,FileName),exc_info=True)
			#Only one worksheet
			#break
		SummaryDict['Table2'] = SummaryList
		(RecomendationCount,TotalCount) = Recomendation(SummaryDict['Table2'])
		SummaryDict['Recommened out of ' + str(TotalCount)] = RecomendationCount
		
		#Summary Page table 3 computation
		PathWithRegions = str(InputPath + '/' + SummaryDict['Market'])
		PathWithVBFDir = PathWithRegions + '/' + filter(lambda z: (z.lower()).strip() == 'vbf', os.listdir(PathWithRegions))[0]
		AbsolutePathOfSHAFile = PathWithVBFDir + '/' + filter(lambda x: (x[-3:].lower()).strip() == 'txt', os.listdir(PathWithVBFDir))[0]

		Table3Data = GetDataFromShaFile(AbsolutePathOfSHAFile,SummaryDict)
		SummaryDict['Table3'] = Table3Data

		SummarypageStatus = SummaryPageCreator(SummaryDict,SummaryWorkSheet,workbook)
		if SummarypageStatus == True:
			logger.info("Summary Page is created for %s."%FileName)
		else:
			logger.error("Error encountered while creating Summary Page in %s"%FileName,exc_info=True)

		workbook.close()
		#print AbsolutePathOfSHAFile
		#SummaryDict['Table3'] = {'CheckSumFile':AbsolutePathOfSHAFile}
		#print SummaryDict	
		
		return True
	except:
		return False	


def ReportsDirChecker(FileName,Path):
	#Check for Report directory present under Input Path/ Returning FileNameWithPath
	if 'MIP-Reports' not in os.listdir(Path):
		CWD = os.getcwd()
		os.chdir(Path)
		os.mkdir('MIP-Reports')
		logger.info("MIP-Reports Directory created under Path: %s",Path)
		os.chdir(CWD)
		FileNameWithPath = Path + '/MIP-Reports/' + FileName
	else:
		#logger.info("MIP-Reports Directory already present under Path: %s",Path)
		FileNameWithPath = Path + '/MIP-Reports/' + FileName
	return FileNameWithPath

def AddSemiColonInEnd(ListOfCSVFiles):
	#Adding semicolon to the end of the file
	for CSVFile in ListOfCSVFiles:
		try:
			with open(CSVFile) as data:
				if list(data)[-1:][0] != '; ; ;':
					f=open(CSVFile, "a+")
					f.write("; ; ;")
					f.close()
		except:
			logger.info("Error encountered while adding the semicolon at the end of %s file",CSVFile)
			logger.error("CSV File not found",exc_info=True)

def ReportNameGenerator(PathWithRegion):
	#global Region,FromVer,ToVer
	Region = re.findall("\w+$",PathWithRegion)[0]
	RegionCodeVal = RegionCode[Region]
	Version = re.findall("\d+-\d+", InputPath)[0]

	FileName = "Test_Report_%s_%s_MIP_SPA_%s"%(Region,RegionCodeVal,Version) + '.xlsx'	
	return FileName

def ZipReportsDir(Path):
	try:
		DirName = Path + '/MIP-Reports'
		#print DirName
		shutil.make_archive(DirName, 'zip', DirName)
		return True
	except:
		return False 
#***************
# Main Program
#***************
def Main(Path):

	global RegionCode
	RegionCode = {'AFME':'901','EU':'904','IND':'905','ISR':'906','NA':'907','PA':'908','SA':'910','SEA':'911','TK':'914'}
	
	global InputPath
	InputPath = Path

	# Removing the backslash '/' from end if present
	Path = Path[:-1] if Path[-1:] == '/' else Path

	Regions = os.listdir(Path)

	#Eliminating the directories which names are not in RegionCode
	Regions = [code for code in Regions if code in RegionCode] 
	
	global PathWithRegions
	# Making list of Paths appending them the region code
	PathWithRegions = [Path + '/' + i for i in Regions]

	#print PathWithRegions

	MainDict = {}
	for i,v in enumerate(PathWithRegions):	
		try:
			#Fetchiing only the directories ending with 'report' under PathWithRegions list.
			try:
				Reports = list(filter(lambda z: (z[-6:].lower()).strip() == 'report', os.listdir(v)))
				PathWithReportDir = [v + '/' + x for x in Reports]
			except:
				logger.error("Directory ending with *Report* string not found under path: %s" %v,exc_info=True)
				sys.exit(1)

			#Fetching only csv files under the 'report' ending directories.
			for index,value in enumerate(PathWithReportDir):
				try:
					PathWithReportDir[index] = value + '/' + filter(lambda a: (a[-3:].lower()).strip() == 'csv' ,os.listdir(value))[0]
				except:
					logger.error("CSV files not found under path: %s" %value,exc_info=True)
			

		except:
			logger.error("Error encountered in working with the data from path: %s" %v,exc_info=True)
			sys.exit(1)
		
		# Create dictionary in the format
		'''
		{filenameRegion1:[file1,file2,...,fileN],
		 filenameRegion2:[file1,file2,...,fileN],
		 filenameRegionN:[file1,file2,...,fileN]}
		'''
		if Reports != []:
			FileName = ReportNameGenerator(PathWithRegions[i])
			MainDict[FileName] = PathWithReportDir

		#For Summary Page
		#RegionVar = re.findall("\w+$",v)[0]
		FromVer = re.findall("\d+",re.findall("\d+-\d+", InputPath)[0])[0]
		ToVer = re.findall("\d+",re.findall("\d+-\d+", InputPath)[0])[1]
		#print RegionVar,FromVer,ToVer
	#print MainDict
	for key,value in MainDict.items():
		AddSemiColonInEnd(value)
		FileNameWithPath = ReportsDirChecker(key,Path)
		logger.info("Creating workbook: %s",key)
		RegionVar = re.findall("Report_(\D+)_",key)[0]
		#RegionVar = re.findall("(\w+)",re.findall("Report_(\w+)",key)[0].replace('_','-'))[0]
		FileCreated = CreateWorkBook(FileNameWithPath,value,key,RegionVar,FromVer,ToVer)

		if FileCreated == True:
			logger.info("File %s creation is completed. Path: %s"%(key,FileNameWithPath))
			#print "File %s creation is completed. Path: %s"%(key,FileNameWithPath)
		else:
			logger.error("Error encountered while creating workbook %s in %s"%(key,FileNameWithPath),exc_info=True)
			#print "Error encountered"
		#only for one workbook
		#break

	FileZipped = ZipReportsDir(Path)
	if FileZipped == True:
		logger.info("MIP-Reports.zip is created under %s"%Path)
	else:
		logger.error("Error encountered while zipping the Report directory",exc_info=True)
		

if __name__ == '__main__':
	try:
		global ReportCreator
		ReportCreator = sys.argv[2]
	except:
		logger.error("Error encountered in second input",exc_info=True)
	try:
		Main(sys.argv[1].strip())
	except IndexError:
		logger.error("No Argumetns passed to script",exc_info=True)
	except OSError:
		logger.error("Invalid Path: %s"%sys.argv[1],exc_info=True)
