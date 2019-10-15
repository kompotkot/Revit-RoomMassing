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

# PREDEFINED VARIABLES
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = __revit__.ActiveUIDocument.Selection


# PARAMS
forFilterFlatType = 'ROM_Type of Layout'				# Getting row from room properties for material name
forFilterLevel = "02" 							# For level filter 
forFilterTower = "R1"							# For tower filter

def materialname(materialname,docname=doc): 	# Takes material from mname. Input: Material name and working Document. Possible to draw on behalf of the material Id
	collectorMaterials = FilteredElementCollector(docname)
	matitr = collectorMaterials.WherePasses(ElementClassFilter(Material)).ToElements()
	matname = None
	for material in matitr:
		if material.Name == materialname:
			matname = material
			break
	return matname

def materialcreator(newmaterialname, docname=doc):			# Create material
	collectorMaterials = FilteredElementCollector(docname)
	matitr = collectorMaterials.WherePasses(ElementClassFilter(Material)).ToElements()
	listofmaterials = []
	for material in matitr:
		listofmaterials.append(material.Name)
	if newmaterialname in listofmaterials:
		return
	else:
		Material.Create(docname, newmaterialname)

# Use a RoomFilter to find all room elements in the document
roomfilter = RoomFilter()
collectorRooms = FilteredElementCollector(doc)		# Apply the filter to the elements in the active document
collectorRooms.WherePasses(roomfilter)
roomiditr = collectorRooms.GetElementIdIterator()	# Get result as ElementId iterator
roomiditr.Reset()
opt = SpatialElementBoundaryOptions()			# Boundary options for rooms

# Cycle through each existing room
while (roomiditr.MoveNext()):
	roomid = roomiditr.Current						# Take roomid
	room = doc.GetElement(roomid) 						# Take room by roomid
	roomLevel = room.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID).AsValueString()
	roomTower = room.LookupParameter('ROM_Number (Section)').AsString()
	roomFlatType = room.LookupParameter(forFilterFlatType).AsString()		# Get name of room type for material name
	if roomLevel == forFilterLevel and roomTower == forFilterTower and roomFlatType != None:
		
		# Get room height
		if type(room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsValueString()) == str:
			
			roomHeightMm = room.get_Parameter(BuiltInParameter.ROOM_HEIGHT).AsValueString()
			roomHeightFt = roomHeightMm.replace(' ','')
			roomHeightFt = roomHeightFt.replace(',','.')
			roomHeightFt = float(roomHeightFt) / 304.8

		else:
			roomHeightFt = int(10)

		# Load family temlate
		docFam = app.NewFamilyDocument("C:\\ProgramData\Autodesk\\RVT 2016\\Family Templates\\English\\Metric Generic Model.rft")
		docFam.SaveAs("D:\\Temp\\" + room.UniqueId.ToString() + ".rfa")

		# Start transaction for model
		m_Trans = Transaction(doc, 'Model transaction to create room boundaries')
		m_Trans.Start()
		# Start transaction for family
		m_TransFam = Transaction(docFam, 'Family transaction to create each Family')
		m_TransFam.Start()

		# Working with room contours
		rvBoundary = room.GetBoundarySegments(opt) 	# Get room boundary
		# Filter room if there are no possibility to generate from

		for rvLoop in rvBoundary: 			# For each loop in the room boundary
			crvarr = CurveArray()			# There will be curve fro each room segments 
			for rvPiece in rvLoop:			# Retrieve each segment of the loop
				dsPiece = rvPiece.Curve 	# Transform to segments
				crvarr.Append(dsPiece)		# Add segments to curve crvarr

			# Form generation in family
			ptOrigin = XYZ(0,0,0)
			ptNormal = XYZ(0,0,1)
			plane = app.Create.NewPlane(ptNormal, ptOrigin)
			sketchPlane = SketchPlane.Create(docFam, plane)
			# Convert the outline. No idea waht difference CurveArrArray of CurveArray 
			curveArrArray = CurveArrArray();
			curveArrArray.Append(crvarr);
			# Generate from. Filter out if no success
			try:
				extrusion = docFam.FamilyCreate.NewExtrusion(True, curveArrArray, sketchPlane, roomHeightFt)
				print('Form generated')
			except:
				print('Form can\'t be generated')
				break

			# Create material in family
			matforflattype = "PNT-" + roomFlatType + "-FormAlgorithm"
			materialcreator(matforflattype,docFam)

			# Find the first geometry face of the given extrusion object
			geomElement = extrusion.get_Geometry(Options())
			geoObjectItor = geomElement.GetEnumerator()
			while (geoObjectItor.MoveNext()):
				solid = geoObjectItor.Current			# Need to find a solid first
				for face in solid.Faces:			# Take each surface and paint it
					docFam.Paint(extrusion.Id, face, materialname(matforflattype,docFam).Id)


		# Finish transaction for family
		m_TransFam.Commit()
		docFam.Save()
		# Finish transaction for model
		m_Trans.Commit()

		# Load family in model
		family = docFam.LoadFamily(doc)

		# ElementId iterator for Family
		collectorFamilySymbols = FilteredElementCollector(doc)
		collectorFamilySymbols.OfClass(FamilySymbol)
		famtypeitr = collectorFamilySymbols.GetElementIdIterator()
		famtypeitr.Reset()

		# Go through each existing family in model and compare
		familySymbol = None 	# Global var, write in from local in cycle

		"""
		# Comments
		cl_sheets = FilteredElementCollector(doc)
		allsheets = cl_sheets.OfCategory(BuiltInCategory.OST_Sheets)
		allviews = allviews.UnionWith(allsheets).WhereElementIsNotElementType().ToElements()
		a = filter(lambda x: x.Name == family.Name.ToString(),famtypeitr) 	# No possibility to take Name from x, like FamilyName
		"""

		#doc.GetElement(famtypeitr.Current).FamilyName

		while (famtypeitr.MoveNext()):
			famid = famtypeitr.Current	# Get familyid
			fam = doc.GetElement(famid)	# Get family by famid
			famname = fam.FamilyName
			if family.Name.ToString() == famname: 	# Compare name of created family and name of family in model
				#print(family.Name.ToString() + " eaqual " + famname)
				familySymbol = fam
				break
			else:
				#print(family.Name.ToString() + " NOT " + famname)
				continue

		m_Trans = Transaction(doc, 'Model transaction to place the Family into the model')
		m_Trans.Start()

		# Activate Symbol
		familySymbol.Activate()
		# Place family
		familyInstance = doc.Create.NewFamilyInstance(XYZ(0, 0, 0), familySymbol, Structure.StructuralType.NonStructural)
		
		m_Trans.Commit()

TaskDialog.Show("Massing Done.", "Did it.")
