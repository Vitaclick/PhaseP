# -*- coding: utf-8 -*-
import clr

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")
from operator import itemgetter
from itertools import groupby

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

def tolist(obj1):
	if hasattr(obj1,"__iter__"): return obj1
	else: return [obj1]
def groupRooms(spaces, partOfName):
	rooms = filter(lambda x: partOfName in x.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), spaces)
	sortRooms = sorted(rooms, key = lambda e: e.get_Parameter(BuiltInParameter.ROOM_NAME).AsString())
	grRoomsName = [[x for x in g] for k,g in groupby(sortRooms, lambda e: e.get_Parameter(BuiltInParameter.ROOM_NAME).AsString())]
	grRoomsLevel = [sorted(g, key = lambda i: i.Level.Elevation) for g in grRoomsName]			
	return grRoomsLevel 
def createAnnoRoom(startPoint, rooms):
	x = startPoint.X
	#TransactionManager.Instance.EnsureInTransaction(doc)
	for g in rooms:
		y = startPoint.Y
		for e in g:
			spName = e.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
			spNum = e.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString()
			spS = e.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble()
			spVol = e.get_Parameter(BuiltInParameter.ROOM_VOLUME).AsDouble()
			spH = e.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).AsDouble()
			spHeat = e.LookupParameter("ОВ.Отопительный прибор").AsString()
			spCold = e.LookupParameter("ОВ.Фанкойл").AsString()
			#spExhaust = e.LookupParameter("ОВ.Вытяжка").AsString()
			
			if not ftype.IsActive : ftype.Activate()
			nf = doc.Create.NewFamilyInstance(XYZ(x,y,0), ftypeS, view)
			nfName = nf.LookupParameter("Д.Имя")
			nfNum = nf.LookupParameter("Д.Номер")
			nfS = nf.LookupParameter("Д.Площадь")
			nfVol = nf.LookupParameter("Д.Объем")
			nfH = nf.LookupParameter("Д.Высота")
			nfHeat = nf.LookupParameter("Г.Отопительный прибор")
			nfCold = nf.LookupParameter("Г.Фанкойл")
			nfName.Set(spName)
			nfNum.Set(spNum)
			nfS.Set(spS)
			nfVol.Set(spVol)
			nfH.Set(spH)
			if spHeat != None and spHeat !="":
				nfHeat.Set(1)
			if spCold != None and spHeat !="":
				nfCold.Set(1)
				
			y+=20/304.8
		x+=25/304.8
	#TransactionManager.Instance.TransactionTaskDone()
	endP = XYZ(x+30/304.8,0,0)
	return endP
		
famAnno = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsElementType().ToElements()
lol = [i.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() for i in famAnno]
ftypeS = UnwrapElement(IN[1])
#ftype = UnwrapElement(IN[2])

view = doc.ActiveView

if view.ViewType != ViewType.DraftingView:
    MessageBox.Show("Открой чертежный вид!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)


"""
spaces = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()

otherRooms = filter(lambda x: "Лестница" not in x.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
								and "ЛХ" not in x.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
								and "ЛФ" not in x.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), spaces)
sortRoomsByLev = sorted(otherRooms, key = lambda e: e.Level.Elevation)
grRoomsByLev = [[x for x in g] for k,g in groupby(sortRoomsByLev, lambda e: e.Level.Elevation)]

TransactionManager.Instance.EnsureInTransaction(doc)

startPoint = XYZ(0,0,0)
# лестницы
stairs = groupRooms(spaces, "Лестница")
stPointElHall = createAnnoRoom(startPoint, stairs)
# лифтовые холлы
elevHall = groupRooms(spaces, "ЛХ")
stPointElev = createAnnoRoom(stPointElHall, elevHall)
# лифты
elev = groupRooms(spaces, "ЛФ")
stPointRooms = createAnnoRoom(stPointElev, elev)

uniqRoomsBylevel = []
y = 0
for group in grRoomsByLev:
	x = stPointRooms.X
	sortRoomsByName = sorted(group, key = lambda i: i.get_Parameter(BuiltInParameter.ROOM_NAME).AsString())
	groupRoomsByName = [[i for i in g] for k,g in groupby(sortRoomsByName, lambda e: e.get_Parameter(BuiltInParameter.ROOM_NAME).AsString())]
	uniqNames = []
	for gr in groupRoomsByName:
		nums = [r.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString() for r in gr]
		nums.sort()
		uName = gr[0].get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
		
		if len(nums) == 1:
			uNum = nums[0]
		elif len(nums) == 2:
			uNum = nums[0] + ", " + nums[1]
		else:
			uNum = nums[0] + "/" + nums[len(nums)-1]
		uniqNames.append([uName, uNum])
		if not ftype.IsActive : ftype.Activate()
		nf = doc.Create.NewFamilyInstance(XYZ(x,y,0), ftype, view)
		nfName = nf.LookupParameter("Д.Имя")
		nfNum = nf.LookupParameter("Д.Номер")
		nfHeat = nf.LookupParameter("Г.Отопительный прибор")
		nfCold = nf.LookupParameter("Г.Фанкойл")
		nfName.Set(uName)
		nfNum.Set(uNum)
		if any(r.LookupParameter("ОВ.Отопительный прибор").AsString() != None and 
		r.LookupParameter("ОВ.Отопительный прибор").AsString() != "" for r in gr):
			nfHeat.Set(1)
		if any(r.LookupParameter("ОВ.Фанкойл").AsString() != None and 
		r.LookupParameter("ОВ.Фанкойл").AsString() != "" for r in gr):
			nfCold.Set(1)		
		
		x+=25/304.8
	y+=20/304.8
	uniqRoomsBylevel.append(uniqNames)

TransactionManager.Instance.TransactionTaskDone()	
		
"""
#MessageBox.Show(str(dir(sys)), "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
