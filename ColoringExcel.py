# Add reference to .Net assembly named Foo.dll
import clr
from System.IO import Directory, Path
import math
#Sorting and managing command
from collections import OrderedDict
import System

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
# Sellect All IFC materials to group them
	#get all the materials in the file
collector = FilteredElementCollector( doc ).OfClass(  Material  )
         
materialsEnum = collector.ToElements() 
print(materialsEnum[0].Name)
ifcMatsCol = []
ifcMatsColKey = []

for i in materialsEnum:
	Name = i.Name
	#print(i.Name)
	if Name[0]+Name[1]+Name[2] == "Ifc":
		ColorValues = i.Color.Red*1000000+((i.Color.Green*1000)+i.Color.Blue)
		ifcMatsCol.append(ColorValues)
		print(ColorValues)

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

#ErrorDetection	
print(ifcMatsColKey.Count)
print(ifcMatsCol.Count)
print(ListListsMat[2].Count)
print(ListListsKeyCols)

#Write to Excel
	# opening a workbook
excel = Excel.ApplicationClass()
workbook = excel.Workbooks.Open("C:\Users\TishGhaith\Desktop\ColTimTray.xlsx")
excel.Visible = True
	# adding a worksheet
worksheet = workbook.Worksheets.Add()
worksheet =  workbook.Sheets[1]

#ChangingColor
def rgb_to_hex(rgb):
    strValue = '%02x%02x%02x' % rgb
    iValue = int(strValue, 16)
    return iValue

	#WriteToCells
#worksheet.Cells[1,1] = 5;
co =1
for color in ListListsKeyCols:
	co+=1
	Red = color[2]
	Green = color[1]
	Blue = color[0]
	co1 = 1
	CellColor =rgb_to_hex((Red , Green, Blue)) 
	
	
	#worksheet.Cells[co1,co] = color	
	for val in color:
		co1+=1
		worksheet.Cells[co,co1] = val
		worksheet.Cells[co,co1].Interior.Color = CellColor

print(ListListsKeyCols[0][1])

workbook.Save()
#workbook.close()
#excel.Quite





#********************** Code End  ******************************

#commit the transaction
t.Commit()

#__window__.Close()
