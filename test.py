# importing packages
import pandas as pd 
import numpy as np
import xlsxwriter

#Reading data file

data_file = pd.read_csv('201907MarchS6GradesAllColleges.csv')


#Creating list of Different colleges under KTU
unique_college = []
for x in data_file['College']:
	if x not in unique_college:
		unique_college.append(x)

listcoll = len(unique_college)-1

#Creating list of Different Branches offered by colleges
unique_branch = []
for x in data_file['Branch']:
	if x not in unique_branch:
		unique_branch.append(x)
listbran = len(unique_branch)-1

#initialising variables
StudCount = 0
SemPass = 0
a=[0,0,0,0,0,0,0,0,0]		#Used to get count of how many students passed for each subject
row = 0
column = 0


#Opening Excel Worksheet
wb = xlsxwriter.Workbook('abc.xlsx')
worksheet = wb.add_worksheet('second')
work = wb.add_worksheet('one')
bold = wb.add_format({'bold': True})


"""WORKSHEET FOR BRANCHWISE"""

for r in unique_branch:
	print('***************\t'+r+'\t********************')
	column = 0
	row+=2
	worksheet.write(row, column, r , bold)
	row +=1
	line1 = ['','SubAStat ','SubBStat ','SubCStat ','SubDStat ','SubEStat ','SubFStat ','SubGStat ','SubHStat ','SubIStat ','SemPass','StudCount: ']
	for items in line1:
		worksheet.write(row, column , items , bold)
		column+=1
	row+=1
	column = 0
	for p in unique_college:
		a=[0,0,0,0,0,0,0,0,0]
		StudCount = 0
		SemPass = 0
		for i, j in data_file.iterrows():
			if j['College'] == p and j['Branch'] == r:
				StudCount+=1
				if j['SemStat'] == 'P':
					SemPass+=1
				if j['SubAStat'] == 'P':
					a[0]+=1
				if j['SubBStat'] == 'P':
					a[1]+=1
				if j['SubCStat'] == 'P':
					a[2]+=1
				if j['SubDStat'] == 'P':
					a[3]+=1
				if j['SubEStat'] == 'P':
					a[4]+=1
				if j['SubFStat'] == 'P':
					a[5]+=1
				if j['SubGStat'] == 'P':
					a[6]+=1
				if j['SubHStat'] == 'P':
					a[7]+=1
				if j['SubIStat'] == 'P':
					a[8]+=1

		if StudCount != 0:
			print('**\t'+p+'\t**')
			line0 = [r]
			line2 = [p,round((a[0]/StudCount)*100, 2),round((a[1]/StudCount)*100, 2),round((a[2]/StudCount)*100, 2),round((a[3]/StudCount)*100, 2),round((a[4]/StudCount)*100, 2),round((a[5]/StudCount)*100, 2),round((a[6]/StudCount)*100, 2),round((a[7]/StudCount)*100, 2),round((a[8]/StudCount)*100, 2),round((SemPass/StudCount)*100, 2)]
			line3 = ['',a[0],a[1],a[2],a[3],a[4],a[5],a[6],a[7],a[8],SemPass,StudCount]
			for items in line2:
				worksheet.write(row, column , items)
				column+=1
			row +=1
			column = 0
			for items in line3:
				worksheet.write(row, column , items)
				column+=1
			row +=1
			column = 0
		

	print("Branchs left: "+str(listbran))      
	listbran-=1
	print()


"""WORKSHEET FOR COLLEGEWISE"""

for p in unique_college:
	print(p)
	column = 0
	tot_SemPass = 0
	tot_StudCount = 0
	row+=2
	work.write(row, column, p , bold)
	row +=1
	line1 = ['','SubAStat ','SubBStat ','SubCStat ','SubDStat ','SubEStat ','SubFStat ','SubGStat ','SubHStat ','SubIStat ','SemPass','StudCount: ']
	for items in line1:
		work.write(row, column , items , bold)
		column+=1
	row+=1
	column = 0
	for r in unique_branch:
		a=[0,0,0,0,0,0,0,0,0]
		StudCount = 0
		SemPass = 0
		for i, j in data_file.iterrows():
			if j['College'] == p and j['Branch'] == r:
				StudCount+=1
				if j['SemStat'] == 'P':
					SemPass+=1
				if j['SubAStat'] == 'P':
					a[0]+=1
				if j['SubBStat'] == 'P':
					a[1]+=1
				if j['SubCStat'] == 'P':
					a[2]+=1
				if j['SubDStat'] == 'P':
					a[3]+=1
				if j['SubEStat'] == 'P':
					a[4]+=1
				if j['SubFStat'] == 'P':
					a[5]+=1
				if j['SubGStat'] == 'P':
					a[6]+=1
				if j['SubHStat'] == 'P':
					a[7]+=1
				if j['SubIStat'] == 'P':
					a[8]+=1

		if StudCount != 0:
			row+=1
			column = 0
			print('**\t'+r+'\t**')
			tot_SemPass+= SemPass
			tot_StudCount+= StudCount
			line2 = [r,round((a[0]/StudCount)*100, 2),round((a[1]/StudCount)*100, 2),round((a[2]/StudCount)*100, 2),round((a[3]/StudCount)*100, 2),round((a[4]/StudCount)*100, 2),round((a[5]/StudCount)*100, 2),round((a[6]/StudCount)*100, 2),round((a[7]/StudCount)*100, 2),round((a[8]/StudCount)*100, 2),round((SemPass/StudCount)*100, 2)]
			line3 = ['',a[0],a[1],a[2],a[3],a[4],a[5],a[6],a[7],a[8],SemPass,StudCount]
			for items in line2:
				work.write(row, column , items)
				column+=1
			row +=1
			column = 0
			for items in line3:
				work.write(row, column , items)
				column+=1
	row +=1
	work.write(row, column , 'Pass%: ', bold)
	column+=1
	work.write(row, column , (tot_SemPass/tot_StudCount)*100 )
	column = 0
		

	print("Colleges left: "+str(listcoll))
	listcoll-=1
	print()

wb.close()

#The End