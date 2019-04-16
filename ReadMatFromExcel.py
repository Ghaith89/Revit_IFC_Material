# Add reference to .Net assembly named Foo.dll
import clr
from System.IO import Directory, Path
import math
#Sorting and managing command
from collections import OrderedDict 
import random
#CommandToUse: 
#OrderedDict.fromkeys(['a', 'b', 'c', 'c', 'a', 'd', 'p', 'p']).keys()
#from RevitServices.Persistence import DocumentManager


#Add Excel Referances
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel
import Microsoft.Office.Interop.Excel as Excel

from System.Runtime.InteropServices import Marshal


#Add Revit Referances
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import*

#set the active Revit application and document
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

#define a transaction variable and describe the transaction
t = Transaction(doc, 'This is my new transaction')

#start a transaction in the Revit database
t.Start()

#********************** Code Start ******************************

#readFromExcel
	#File Directory:  C:\GhaithT\IFC\MaterialScheduelRead.xlsx
#Write to Excel
	# opening a workbook
excel = Excel.ApplicationClass()
workbookToRead = excel.Workbooks.Open("C:\Users\TishGhaith\Desktop\MatToRead.xlsx")
#selection = excel.Selection
excel.Visible = True
	# adding a worksheet
#worksheetinRead = workbookToRead.Worksheets.Add()
worksheetinRead =  workbookToRead.Sheets[1]

#print(worksheetinRead.Cells[2,1].Value2)

MatNames = []
for i in range(20):
	i = i+1
	#print(worksheetinRead.Cells[i,1].Value2)
	MatNames.append(worksheetinRead.Cells[i,1].Value2)
#print(MatNames)

#GetMaterials From Revit
collector = FilteredElementCollector( doc ).OfClass(  Material  )
         
materialsEnum = collector.ToElements() 
#print(materialsEnum[0].Name)
ifcMatsCol = []
ifcMatsColKey = []

co = -1
for i in materialsEnum:
	co+=1
	Name = i.Name
	#print(i.Name)
	if Name[0]+Name[1]+Name[2] == "Ifc":
		ColorValues = i.Color.Red*1000000+((i.Color.Green*1000)+i.Color.Blue)
		#ColorValues = 5
		ifcMatsCol.append(ColorValues)
		#print(i.Color)
		

#Get Color key value
ifcMatsColKey = OrderedDict.fromkeys(ifcMatsCol).keys()

#Categorise the sellected ifc materials into key groups
ListListsMat = []
ListListsCols = []
ListListsKeyCols = []

for k in ifcMatsColKey:
	SubMat = []
	SubCol = []
	for i in materialsEnum:
		Color = []
		MatColVal = i.Color.Red*1000000+((i.Color.Green*1000)+i.Color.Blue)
		if MatColVal == k :
			SubMat.append(i)
			Color.append(i.Color.Red)
			Color.append(i.Color.Green)
			Color.append(i.Color.Blue)
			SubCol.append(Color)
	ListListsMat.append(SubMat)
		
	ListListsCols.append(SubCol)
	

for i in ListListsCols:
	ListListsKeyCols.append(i[0])
	
#print(ListListsMat)
#print(ListListsCols)
#print(ListListsCols.Count)
#Changing Material Appearance based on Data From Excel
for i in ListListsMat:
	IN=0
	for n in i:
		IN+=1
		co = 1
		MatKey = n.Color.Red*1000000+((n.Color.Green*1000)+n.Color.Blue)
		#print(MatKey)
	#match the color value with the value from Excel
		for l in range(20):
			co+=1
			
			if worksheetinRead.Cells[co, 1].Value2 != None :
				#Changing appearance according to the material name from excel
				
				keyExcel = (worksheetinRead.Cells[co, 2].Value2*1000000) + ((worksheetinRead.Cells[co, 3].Value2*1000) + worksheetinRead.Cells[co, 4].Value2)
				#print(str(keyExcel))
				
				
				if MatKey == keyExcel:
					print("Ja00000")
					appearanceId = n.AppearanceAssetId
					#Check if the Excel Layer is not empty
					
					IntendedMaApName = worksheetinRead.Cells[co, 1].Value2
					RealMatName = worksheetinRead.Cells[co, 5].Value2+ str(IN)+ str(co)
					#Searching for the relevant Material in Revit to get its appearance
					#Dr=0
					for mat in materialsEnum:
						#Dr+=1
						
						if mat.Name == IntendedMaApName:
						 	#print(IntendedMaApName)
						 	#get the apearance Id of the intended material
						 	MatApAsId = mat.AppearanceAssetId
						 	n.AppearanceAssetId = MatApAsId
						 	#print(RealMatName)
						 	n.Name = RealMatName
						 	print(RealMatName)
						 	
					 				

#********************** Code End  ******************************

#commit the transaction
t.Commit()