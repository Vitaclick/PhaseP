# -*- coding: utf-8 -*-
import clr

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

import sys
sys.path.append("C:/Program Files (x86)/IronPython 2.7/Lib")
from operator import itemgetter
from itertools import groupby

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import *
from System.Windows.Forms import *

EXVALUES = (None, "")
PARTNAMES = ('Лестница', 'ЛХ', 'ЛФ')

def tolist(obj1):
	if hasattr(obj1,"__iter__"): return obj1
	else: return [obj1]

def groupRooms(spaces, partOfName):
	rooms = filter(lambda x: partOfName in x.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(), spaces)
	sortRooms = sorted(rooms, key = lambda e: e.get_Parameter(BuiltInParameter.ROOM_NAME).AsString())
	grRoomsName = [[x for x in g] for k,g in groupby(sortRooms, lambda e: e.get_Parameter(BuiltInParameter.ROOM_NAME).AsString())]
	grRoomsLevel = [sorted(g, key = lambda i: i.Level.Elevation) for g in grRoomsName]
	return grRoomsLevel

def setOnAndSystem(spaceParameter, annoParameterOn, annoParameterSys):
	if spaceParameter not in EXVALUES:
		annoParameterOn.Set(1)
		annoParameterSys.Set(spaceParameter)

def setGroupOnAndSystem(gr, paramName, nfParamOn, nfParamSys):
	for r in gr:
		a = r.LookupParameter(paramName).AsString()
		if a not in EXVALUES:
			nfParamOn.Set(1)
			nfParamSys.Set(a)
			break

def setHeatColdOn(spaceParameter, annoParameterOn):
	if spaceParameter not in EXVALUES:
		annoParameterOn.Set(1)

def setGroupHeatCold(gr, paramName, nfParam):
	if any(r.LookupParameter(paramName).AsString() not in EXVALUES for r in gr):
		nfParam.Set(1)

def createAnnoRoom(startPoint, rooms):
	x = startPoint.X
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
			spExhaust1 = e.LookupParameter("ОВ.Вытяжка1").AsString()
			spExhaust2 = e.LookupParameter("ОВ.Вытяжка2").AsString()
			spInflux1 = e.LookupParameter("ОВ.Приток1").AsString()
			spInflux2 = e.LookupParameter("ОВ.Приток2").AsString()

			if not ftypeS.IsActive : ftypeS.Activate()
			nf = doc.Create.NewFamilyInstance(XYZ(x,y,0), ftypeS, view)
			nfName = nf.LookupParameter("Д.Имя")
			nfNum = nf.LookupParameter("Д.Номер")
			nfS = nf.LookupParameter("Д.Площадь")
			nfVol = nf.LookupParameter("Д.Объем")
			nfH = nf.LookupParameter("Д.Высота")
			nfHeat = nf.LookupParameter("Г.Отопительный прибор")
			nfCold = nf.LookupParameter("Г.Фанкойл")
			nfExhaust1On = nf.LookupParameter("Г.Вытяжка1")
			nfExhaust1Sys = nf.LookupParameter("Т.Вытяжка1_Система")
			nfExhaust2On = nf.LookupParameter("Г.Вытяжка2")
			nfExhaust2Sys = nf.LookupParameter("Т.Вытяжка2_Система")
			nfInflux1On = nf.LookupParameter("Г.Приток1")
			nfInflux1Sys = nf.LookupParameter("Т.Приток1_Система")
			nfInflux2On = nf.LookupParameter("Г.Приток2")
			nfInflux2Sys = nf.LookupParameter("Т.Приток2_Система")
			nfName.Set(spName)
			nfNum.Set(spNum)
			nfS.Set(spS)
			nfVol.Set(spVol)
			nfH.Set(spH)
			setHeatColdOn(spHeat, nfHeat)
			setHeatColdOn(spCold, nfCold)
			setOnAndSystem(spExhaust1, nfExhaust1On, nfExhaust1Sys)
			setOnAndSystem(spExhaust2, nfExhaust2On, nfExhaust2Sys)
			setOnAndSystem(spInflux1, nfInflux1On, nfInflux1Sys)
			setOnAndSystem(spInflux2, nfInflux2On, nfInflux2Sys)
				
			y+=20/304.8
		x+=25/304.8
	endP = XYZ(x+30/304.8,0,0)
	return endP
		

view = doc.ActiveView

if view.ViewType != ViewType.DraftingView:
    MessageBox.Show("Открой чертежный вид!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
else:
	fAnno = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).WhereElementIsElementType().ToElements()
	ftypeS = filter(lambda x: x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == "ТА_Помещение_с площадью", fAnno)[0]
	ftype = filter(lambda x: x.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() == "ТА_Помещение_без площади", fAnno)[0]

	spaces = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_MEPSpaces).WhereElementIsNotElementType().ToElements()

	

	otherRooms = filter(lambda x: all(pn not in x.get_Parameter(BuiltInParameter.ROOM_NAME).AsString() for pn in PARTNAMES), spaces)
	sortRoomsByLev = sorted(otherRooms, key = lambda e: e.Level.Elevation)
	grRoomsByLev = [[x for x in g] for k,g in groupby(sortRoomsByLev, lambda e: e.Level.Elevation)]


	t = Transaction(doc, "CreateP")
	t.Start()

	startPoint = XYZ(0,0,0)
	# лестницы, лифты, лифтовые холлы
	for pn in PARTNAMES:
		sp = groupRooms(spaces, pn)
		startPoint = createAnnoRoom(startPoint, sp)

	# остальные помещения
	uniqRoomsBylevel = []
	y = 0
	for group in grRoomsByLev:
		x = startPoint.X
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
			nfExhaust1On = nf.LookupParameter("Г.Вытяжка1")
			nfExhaust1Sys = nf.LookupParameter("Т.Вытяжка1_Система")
			nfExhaust2On = nf.LookupParameter("Г.Вытяжка2")
			nfExhaust2Sys = nf.LookupParameter("Т.Вытяжка2_Система")
			nfInflux1On = nf.LookupParameter("Г.Приток1")
			nfInflux1Sys = nf.LookupParameter("Т.Приток1_Система")
			nfInflux2On = nf.LookupParameter("Г.Приток2")
			nfInflux2Sys = nf.LookupParameter("Т.Приток2_Система")
			nfName.Set(uName)
			nfNum.Set(uNum)
			setGroupHeatCold(gr, "ОВ.Отопительный прибор", nfHeat)
			setGroupHeatCold(gr, "ОВ.Фанкойл", nfCold)
			setGroupOnAndSystem(gr, "ОВ.Вытяжка1", nfExhaust1On, nfExhaust1Sys)
			setGroupOnAndSystem(gr, "ОВ.Вытяжка2", nfExhaust2On, nfExhaust2Sys)
			setGroupOnAndSystem(gr, "ОВ.Приток1", nfInflux1On, nfInflux1Sys)
			setGroupOnAndSystem(gr, "ОВ.Приток2", nfInflux2On, nfInflux2Sys)

			x+=25/304.8
		y+=20/304.8
		uniqRoomsBylevel.append(uniqNames)


	t.Commit()


#MessageBox.Show(str(dir(sys)), "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Information)
