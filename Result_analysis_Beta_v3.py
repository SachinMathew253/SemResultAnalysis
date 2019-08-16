# importing packages
import pandas as pd 
import numpy as np
import xlsxwriter
import time
from collections import defaultdict

start_time = time.time()
#Reading data file
data_xls = pd.read_excel('201907MarchS6GradesAllColleges.xlsx', 'gradelist', index_col=None)
data_xls.to_csv('beta.csv', encoding='utf-8', index=False)
data_file = pd.read_csv('beta.csv')


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

#Used to get count of how many students passed for each subject
row = 0
column = 0
j = 0
abc = [0,0,0,0,0,0,0,0,0,0]		
tab = ['A' , 'B' , 'C' , 'D' , 'E' , 'F' , 'G' , 'H' , 'I', 'SemPass']
test = {}
branch = defaultdict(dict)
sub = defaultdict(dict)



#Opening Excel Worksheet
wb = xlsxwriter.Workbook('Result_Analysis_Beta1.xlsx')
worksheet = wb.add_worksheet('BRANCHWISE')
work = wb.add_worksheet('COLLEGEWISE')
work_1 = wb.add_worksheet('SUBJECTWISE')
bold = wb.add_format({'bold': True})



"""WORKSHEET FOR BRANCHWISE"""

for i, j in data_file.iterrows():
		test[j['College'] , j['Branch']]=[0,0,0,0,0,0,0,0,0,0,0]

for b in unique_branch:
	for a in unique_college:
		if test.get((a,b)):
			branch[b][a] = [[0,0,0,0,0,0,0,0,0,0,0] , 0]
	sub[b] = {key: [] for key in tab}
	

for i, j in data_file.iterrows():
				test[j['College'] , j['Branch']][10]+=1
				if j['SemStat'] == 'P':
					test[j['College'] , j['Branch']][9]+=1
				if j['SubAStat'] == 'P':
					test[j['College'] , j['Branch']][0]+=1
				if j['SubBStat'] == 'P':
					test[j['College'] , j['Branch']][1]+=1
				if j['SubCStat'] == 'P':
					test[j['College'] , j['Branch']][2]+=1
				if j['SubDStat'] == 'P':
					test[j['College'] , j['Branch']][3]+=1
				if j['SubEStat'] == 'P':
					test[j['College'] , j['Branch']][4]+=1
				if j['SubFStat'] == 'P':
					test[j['College'] , j['Branch']][5]+=1
				if j['SubGStat'] == 'P':
					test[j['College'] , j['Branch']][6]+=1
				if j['SubHStat'] == 'P':
					test[j['College'] , j['Branch']][7]+=1
				if j['SubIStat'] == 'P':
					test[j['College'] , j['Branch']][8]+=1

for b in unique_branch:
	for a in unique_college:
		if test.get((a,b)):
			branch[b][a][1] = round((test[a,b][9]/test[a,b][10])*100 , 2)
			branch[b][a][0] = test[a,b][:]

for b in unique_branch:
	branch[b] = sorted(branch[b].items(), key = lambda x:x[1][1] , reverse = True)
	branch[b] = dict(branch[b])
			
tot_SemPass =0
tot_StudCount = 0


for b in unique_branch:
	tot_SemPass =0
	tot_StudCount = 0
	row+=2
	worksheet.write(row, column, b , bold)
	row +=1
	line1 = ['','SubAStat ','SubBStat ','SubCStat ','SubDStat ','SubEStat ','SubFStat ','SubGStat ','SubHStat ','SubIStat ','SemPass','StudCount: ']
	for items in line1:
		worksheet.write(row, column , items , bold)
		column+=1
	row+=1
	column = 0
	for key in branch[b]:
		tot_SemPass += branch[b][key][0][9]
		tot_StudCount += branch[b][key][0][10]
		column = 0
		line2 = [key,round((branch[b][key][0][0]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][1]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][2]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][3]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][4]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][5]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][6]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][7]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][8]/branch[b][key][0][10])*100, 2),round((branch[b][key][0][9]/branch[b][key][0][10])*100, 2)]
		line3 = ['',branch[b][key][0][0],branch[b][key][0][1],branch[b][key][0][2],branch[b][key][0][3],branch[b][key][0][4],branch[b][key][0][5],branch[b][key][0][6],branch[b][key][0][7],branch[b][key][0][8],branch[b][key][0][9],branch[b][key][0][10]]
		for items in line2:
			worksheet.write(row, column , items)
			column+=1
		row +=1
		column = 0
		for items in line3:
			worksheet.write(row, column , items)
			column+=1
		row +=1
	column-=2
	worksheet.write(row, column ,tot_SemPass)
	worksheet.write(row, column+1 ,tot_StudCount)
	worksheet.write(row+1, column , 'Pass%: ', bold)
	column+=1
	worksheet.write(row+1, column, round((tot_SemPass/tot_StudCount)*100, 2))
	row+=2
	column = 0			
		

tot_SemPass =0
tot_StudCount = 0

"""WORKSHEET FOR COLLEGEWISE"""
branch.clear()
for a in unique_college:
	for b in unique_branch:
		if test.get((a,b)):
			branch[a]['Pass'] = 0
			branch[a][b] = [[0,0,0,0,0,0,0,0,0,0,0] , 0]
			
row = 0


for a in unique_college:
	tot_StudCount = 0
	for b in unique_branch:
		if test.get((a,b)):
			branch[a][b][1] = round((test[a,b][9]/test[a,b][10])*100 , 2)
			branch[a][b][0] = test[a,b][:]
			tot_StudCount += test[a,b][10] 
			branch[a]['Pass'] += test[a,b][9]
	branch[a]['Pass'] = round((branch[a]['Pass']/tot_StudCount)*100 , 2)
branch = sorted(branch.items(), key = lambda x:x[1]['Pass'] , reverse = True)
branch = dict(branch)

for a in branch:
	tot_StudCount = 0
	tot_SemPass = 0
	row+=2
	work.write(row, column, a , bold)
	row +=1
	line1 = ['','SubAStat ','SubBStat ','SubCStat ','SubDStat ','SubEStat ','SubFStat ','SubGStat ','SubHStat ','SubIStat ','SemPass','StudCount: ']
	for items in line1:
		work.write(row, column , items , bold)
		column+=1
	row+=1
	for key in branch[a]:
			if key == 'Pass':
				continue
			tot_StudCount += branch[a][key][0][10] 
			tot_SemPass += branch[a][key][0][9]
			line2 = [key,round((branch[a][key][0][0]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][1]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][2]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][3]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][4]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][5]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][6]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][7]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][8]/branch[a][key][0][10])*100, 2),round((branch[a][key][0][9]/branch[a][key][0][10])*100, 2)]
			line3 = ['',branch[a][key][0][0],branch[a][key][0][1],branch[a][key][0][2],branch[a][key][0][3],branch[a][key][0][4],branch[a][key][0][5],branch[a][key][0][6],branch[a][key][0][7],branch[a][key][0][8],branch[a][key][0][9],branch[a][key][0][10]]
			column = 0
			row+=1
			for items in line2:
				work.write(row, column , items)
				column+=1
			row +=1
			column = 0
			for items in line3:
				work.write(row, column , items)
				column+=1
	row +=1
	column-=2
	work.write(row, column ,tot_SemPass)
	work.write(row, column+1 ,tot_StudCount)
	work.write(row+1, column , 'Pass%: ', bold)
	work.write(row+1, column+1 ,branch[a]['Pass'] )
	column = 0
		

""" Worksheet SUBJECTWISE """
branch.clear()
branch = defaultdict(dict)
for b in unique_branch:
	for a in unique_college:
		if test.get((a,b)):
			branch[b][a] = [[0,0,0,0,0,0,0,0,0,0,0] , 0]
	

for b in unique_branch:
	for a in unique_college:
		if test.get((a,b)):
			for i in range(10):
				abc[i] = (round((test[a,b][i]/test[a,b][10])*100,2))
			branch[b][a][1] = round((test[a,b][9]/test[a,b][10])*100 , 2)
			branch[b][a][0] = abc[:]


			
row = 0
for b in unique_branch:
	for i in range(10):
		#for a in unique_college:
			branch[b] = sorted(branch[b].items(), key = lambda x:x[1][0][i] , reverse = True)
			branch[b] = dict(branch[b])
			for key in branch[b]:
				sub[b][tab[i]].append(key)	


row = 0
column = 0
j = 0
for b in unique_branch:
	j = 0
	row+=2
	work_1.write(row, column, b , bold)
	row +=1
	line1 = ['SubA ','SubB ','SubC ','SubD ','SubE ','SubF ','SubG ','SubH ','SubI ','','','Rank','SemPass']
	for items in line1:
		work_1.write(row, column , items , bold)
		column+=1
	row+=1
	column = 0
	for i in range(len(sub[b]['A'])):
		j += 1
		line2 = [sub[b]['A'][i],sub[b]['B'][i],sub[b]['C'][i],sub[b]['D'][i],sub[b]['E'][i],sub[b]['F'][i],sub[b]['G'][i],sub[b]['H'][i],sub[b]['I'][i],'','',j,sub[b]['SemPass'][i][:9],branch[b][sub[b]['SemPass'][i]][0][9]]
		line3 = [branch[b][sub[b]['A'][i]][0][0],branch[b][sub[b]['B'][i]][0][1],branch[b][sub[b]['C'][i]][0][2],branch[b][sub[b]['D'][i]][0][3],branch[b][sub[b]['E'][i]][0][4],branch[b][sub[b]['F'][i]][0][5],branch[b][sub[b]['G'][i]][0][6],branch[b][sub[b]['H'][i]][0][7],branch[b][sub[b]['I'][i]][0][8]]
		for items in line2:
			work_1.write(row, column , items)
			column+=1
		row +=1
		column = 0
		for items in line3:
			work_1.write(row, column , items)
			column+=1
		row +=1
		column = 0
		

		

wb.close()
print("---%r Sec---" %(round((time.time() - start_time),2)))


#The End
