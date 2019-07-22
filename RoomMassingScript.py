#__doc__ = 'Print the full path to the central model (if model is workshared).'

# IMPORTS
import clr
import System


clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("RevitServices")
clr.AddReference("RevitNodes")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB.Architecture import RoomFilter
from RevitServices.Persistence import DocumentManager
from System import Guid

#PREDEFINED VARIABLES
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = __revit__.ActiveUIDocument.Selection


#PARAMS
forFilterFlatType = 'ROM_Type of Layout'		#Для получения нужной строки из комнаты для названия материала
forFilterLevel = "02" 							#Для фильтра по уровню
forFilterTower = "R1"							#Для фильтра по башне

def materialname(materialname,docname=doc): 	#Берет материал из mname. На ввод Имя материала и Документ с которым работаем. Можно обрать от имени материала Id
	collectorMaterials = FilteredElementCollector(docname)
	matitr = collectorMaterials.WherePasses(ElementClassFilter(Material)).ToElements()
	matname = None
	for material in matitr:
		if material.Name == materialname:
			matname = material
			break
	return matname

def materialcreator(newmaterialname, docname=doc):		#Создает материал
	collectorMaterials = FilteredElementCollector(docname)
	matitr = collectorMaterials.WherePasses(ElementClassFilter(Material)).ToElements()
	listofmaterials = []
	for material in matitr:
		listofmaterials.append(material.Name)
	if newmaterialname in listofmaterials:
		return
	else:
		Material.Create(docname, newmaterialname)

#Use a RoomFilter to find all room elements in the document
roomfilter = RoomFilter()
collectorRooms = FilteredElementCollector(doc)		#Apply the filter to the elements in the active document
collectorRooms.WherePasses(roomfilter)
roomiditr = collectorRooms.GetElementIdIterator()	#Get result as ElementId iterator
roomiditr.Reset()
opt = SpatialElementBoundaryOptions()				#Boundary options for rooms

#Пробегаемся по каждой существующей комнате
while (roomiditr.MoveNext()):
	roomid = roomiditr.Current						#Вытащить roomid
	room = doc.GetElement(roomid) 					#Вытащить комнату по roomid
	roomLevel = room.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID).AsValueString()
	roomTower = room.LookupParameter('ROM_Number (Section)').AsString()
	roomFlatType = room.LookupParameter(forFilterFlatType).AsString()		#Получаем название тима квартиры для названия материала
	if roomLevel == forFilterLevel and roomTower == forFilterTower and roomFlatType != None:
		
		#Отбираем высоту помещения
		if type(room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsValueString()) == str:
			
			roomHeightMm = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsValueString()
			roomHeightFt = roomHeightMm.replace(' ','')
			roomHeightFt = roomHeightFt.replace(',','.')
			roomHeightFt = float(roomHeightFt) / 304.8

		else:
			roomHeightFt = int(10)

		#Загружаем шаблон семейства
		docFam = app.NewFamilyDocument("C:\\ProgramData\Autodesk\\RVT 2016\\Family Templates\\English\\Metric Generic Model.rft")
		docFam.SaveAs("D:\\Temp\\" + room.UniqueId.ToString() + ".rfa")

		#Начинаем транзакцию для модели
		m_Trans = Transaction(doc, 'Model transaction to create room boundaries')
		m_Trans.Start()
		#Начинаем транзакцию для семейства
		m_TransFam = Transaction(docFam, 'Family transaction to create each Family')
		m_TransFam.Start()

		#Работаем с контурами помещения
		rvBoundary = room.GetBoundarySegments(opt) #Get room boundary
		#Фильтруем помещение если неудалось выдавить форму

		for rvLoop in rvBoundary: 			#For each loop in the room boundary
			crvarr = CurveArray()			#Здесь будет кривая для каждой комнаты объединенная из отрезков
			for rvPiece in rvLoop:			#Retrieve each segment of the loop
				dsPiece = rvPiece.Curve 	#Трансформируем в отрезки
				crvarr.Append(dsPiece)		#Добавляем отрезки в Кривую crvarr

			#Генерация форм в семействе
			ptOrigin = XYZ(0,0,0)
			ptNormal = XYZ(0,0,1)
			plane = app.Create.NewPlane(ptNormal, ptOrigin)
			sketchPlane = SketchPlane.Create(docFam, plane)
			#Преобразуем контур. Хз чем отличется CurveArrArray от CurveArray
			curveArrArray = CurveArrArray();
			curveArrArray.Append(crvarr);
			#Выдавливаем форму. Фильтруем если не удалось выдавить
			try:
				extrusion = docFam.FamilyCreate.NewExtrusion(True, curveArrArray, sketchPlane, roomHeightFt)
				print('Форма выдавилась')
			except:
				print('Неудалось выдавить форму')
				break

			#Создаем материал в семействе
			matforflattype = "PNT-" + roomFlatType + "-FormAlgorithm"
			materialcreator(matforflattype,docFam)

			#Find the first geometry face of the given extrusion object
			geomElement = extrusion.get_Geometry(Options())
			geoObjectItor = geomElement.GetEnumerator()
			while (geoObjectItor.MoveNext()):
				solid = geoObjectItor.Current		# need to find a solid first
				for face in solid.Faces:			# берем каждую поверхность и красим ее
					docFam.Paint(extrusion.Id, face, materialname(matforflattype,docFam).Id)


		#Завершаем транзакцию для семейства
		m_TransFam.Commit()
		docFam.Save()
		#Завершаем транзакцию для модели
		m_Trans.Commit()

		#Загружаем семейство в модель
		family = docFam.LoadFamily(doc)

		#ElementId iterator for Family
		collectorFamilySymbols = FilteredElementCollector(doc)
		collectorFamilySymbols.OfClass(FamilySymbol)
		famtypeitr = collectorFamilySymbols.GetElementIdIterator()
		famtypeitr.Reset()

		#Пробегаемся по каждому существующему семейству в модели и сравниваем
		familySymbol = None 	#Глобальная переменная в которую пишем из локальной в цикле

		"""
		#Комментарии от Алексея
		cl_sheets = FilteredElementCollector(doc)
		allsheets = cl_sheets.OfCategory(BuiltInCategory.OST_Sheets)
		allviews = allviews.UnionWith(allsheets).WhereElementIsNotElementType().ToElements()
		a = filter(lambda x: x.Name == family.Name.ToString(),famtypeitr) 	#Не позволяет взять Name от х, как и FamilyName
		"""

		#doc.GetElement(famtypeitr.Current).FamilyName

		while (famtypeitr.MoveNext()):
			famid = famtypeitr.Current	#Вытащить familyid
			fam = doc.GetElement(famid)	#Вытащить семейство по famid
			famname = fam.FamilyName
			if family.Name.ToString() == famname: 	#Сравниваем название созданного семейства и название семейства в модели
				#print(family.Name.ToString() + " eaqual " + famname)
				familySymbol = fam
				break
			else:
				#print(family.Name.ToString() + " NOT " + famname)
				continue

		m_Trans = Transaction(doc, 'Model transaction to place the Family into the model')
		m_Trans.Start()

		#Активируем Symbol
		familySymbol.Activate()
		#Размещаем семейство
		familyInstance = doc.Create.NewFamilyInstance(XYZ(0, 0, 0), familySymbol, Structure.StructuralType.NonStructural)
		
		m_Trans.Commit()

TaskDialog.Show("Done", "Did it.")
