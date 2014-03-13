import sys
reload(sys)
sys.setdefaultencoding("latin-1")

import xlrd, arcpy, string, os, zipfile, fileinput, time
from datetime import date
start_time = time.time()

INPUT_PATH = "input"
OUTPUT_PATH = "output"
if arcpy.Exists(OUTPUT_PATH + "\\LakePartner.gdb"):
	os.system("rmdir " + OUTPUT_PATH + "\\LakePartner.gdb /s /q")
os.system("del " + OUTPUT_PATH + "\\*LakePartner*.*")
arcpy.CreateFileGDB_management(OUTPUT_PATH, "LakePartner", "9.3")
arcpy.env.workspace = OUTPUT_PATH + "\\LakePartner.gdb"

def parseLatLng(latlng):
	if len(str(latlng).strip()) == 0:
		return 0
	latlngInt = int(latlng)
	degree = int(latlngInt / 10000)
	temp = latlngInt - degree * 10000
	minute = int(temp / 100)
	second = int(temp - minute * 100)
	return degree + (minute/60.0) + (second/3600.0)
def parseValue(value):
	if len(str(value).strip()) == 0:
		return None
	return value
	
def createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields):
	print "Create " + featureName + " feature class"
	featureNameNAD83 = featureName + "_NAD83"
	featureNameNAD83Path = arcpy.env.workspace + "\\"  + featureNameNAD83
	arcpy.CreateFeatureclass_management(arcpy.env.workspace, featureNameNAD83, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
	# Process: Define Projection
	arcpy.DefineProjection_management(featureNameNAD83Path, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
	# Process: Add Fields	
	for featrueField in featureFieldList:
		arcpy.AddField_management(featureNameNAD83Path, featrueField[0], featrueField[1], featrueField[2], featrueField[3], featrueField[4], featrueField[5], featrueField[6], featrueField[7], featrueField[8])
	# Process: Append the records
	cntr = 1
	try:
		with arcpy.da.InsertCursor(featureNameNAD83, featureInsertCursorFields) as cur:
			for rowValue in featureData:
				cur.insertRow(rowValue)
				cntr = cntr + 1
	except Exception as e:
		print "\tError: " + featureName + ": " + e.message
	# Change the projection to web mercator
	arcpy.Project_management(featureNameNAD83Path, arcpy.env.workspace + "\\" + featureName, "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
	arcpy.Delete_management(featureNameNAD83Path, "FeatureClass")
	print "Finish " + featureName + " feature class."

stationsDict = {}
TPCountDict = {}
featureName = "TotalPhosphorus"
featureData = []
wb = xlrd.open_workbook('input\\TP for annual report for the web 2013.xls')
sh = wb.sheet_by_name(u'LPP TP Data 2013')
startRowNum = 8; # start with 0
for rownum in range(startRowNum, sh.nrows):
	row = sh.row_values(rownum)
	#print row
	#print "start.."
	STN = int(row[2])
	#print type(STN)
	SITEID = int(row[3])
	#print type(SITEID)
	Date_ = row[7]
	#print type(Date_)
	TP1 = parseValue(row[8])
	#print type(TP1)
	TP2 = parseValue(row[9])
	#print type(TP2)
	DataCollector = row[10]  # unicode
	#print type(DataCollector)
	MajorDifference = row[11]  # str
	#print type(MajorDifference)
	ID = STN * 10000 + SITEID
	if ID in TPCountDict:
		TPCountDict[ID] = TPCountDict[ID] + 1
	else:
		TPCountDict[ID] = 1
	#ID = 0
	stationsDict[ID] = row
	Latitude = parseLatLng(row[5])
	#print type(Latitude)
	Longitude = -parseLatLng(row[6])
	#print type(Longitude)
	featureData.append([(Longitude, Latitude), STN, SITEID, Date_, TP1, TP2, DataCollector, MajorDifference, ID, Latitude, Longitude])
#print featureData
featureFieldList = [["STN", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["SITEID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["Date_", "DATE", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["TP1", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["TP2", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DataCollector", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MajorDifference", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["Latitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Longitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "STN", "SITEID", "Date_", "TP1", "TP2", "DataCollector", "MajorDifference", "ID", "Latitude", "Longitude")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)


featureName = "SecchiDepth"
SECountDict = {}
featureData = []
wb = xlrd.open_workbook('input\\Average Secchi data for annual report 2013.xls')
sh = wb.sheet_by_name(u'LPP Secchi Data 2013')
startRowNum = 2; # start with 0
for rownum in range(startRowNum, sh.nrows):
	row = sh.row_values(rownum)
	STN = int(row[2])
	SITEID = int(row[3])
	Year_ = int(row[7])
	SecchiDepth = parseValue(row[8])
	ID = STN * 10000 + SITEID
	if ID in SECountDict:
		SECountDict[ID] = SECountDict[ID] + 1
	else:
		SECountDict[ID] = 1	
	stationsDict[ID] = row
	Latitude = parseLatLng(row[5])
	Longitude = -parseLatLng(row[6])
	featureData.append([(Longitude, Latitude), STN, SITEID, Year_, SecchiDepth, ID, Latitude, Longitude])
featureFieldList = [["STN", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["SITEID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["Year_", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["SecchiDepth", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["Latitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Longitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "STN", "SITEID", "Year_", "SecchiDepth", "ID", "Latitude", "Longitude")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)


featureName = "LAKE_PARTNERS_STATIONS"
featureData = []
for ID in stationsDict.keys():
	row = stationsDict[ID]
	STN = int(row[2])
	SITEID = int(row[3])
	ID = STN * 10000 + SITEID
	LAKENAME = row[0]
	TOWNSHIP = row[1]
	SITEDESC = row[4]
	Latitude = parseLatLng(row[5])
	Longitude = -parseLatLng(row[6])
	SE_COUNT = SECountDict[ID] if ID in SECountDict else 0
	PH_COUNT = TPCountDict[ID] if ID in TPCountDict else 0	
	featureData.append([(Longitude, Latitude), ID, LAKENAME, TOWNSHIP, STN, SITEID, SITEDESC, Latitude, Longitude, SE_COUNT, PH_COUNT])	
featureFieldList = [["ID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["LAKENAME", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["TOWNSHIP", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["STN", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["SITEID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["SITEDESC", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["LATITUDE", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["LONGITUDE", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SE_COUNT", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["PH_COUNT", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "ID", "LAKENAME", "TOWNSHIP", "STN", "SITEID", "SITEDESC", "LATITUDE", "LONGITUDE", "SE_COUNT", "PH_COUNT")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

os.system("copy " + INPUT_PATH + "\\LakePartner.msd " + OUTPUT_PATH)
os.system("copy " + INPUT_PATH + "\\LakePartner.mxd " + OUTPUT_PATH)
f = open (INPUT_PATH + "\\readme_LakePartner.txt","r")
data = f.read()
f.close()
import time
dateString = time.strftime("%Y/%m/%d", time.localtime())
data = data.replace("[DATE]", dateString)
f = open (OUTPUT_PATH + "\\readme_LakePartner.txt","w")
f.write(data)
f.close()

# Compress the msd, mxd, readme.txt and file geodatabase together into a zip file named LakePartner.zip, which will be send to web service publisher. 

target_dir = OUTPUT_PATH + '\\LakePartner.gdb'
zip = zipfile.ZipFile(OUTPUT_PATH + '\\LakePartner.zip', 'w', zipfile.ZIP_DEFLATED)
rootlen = len(target_dir) + 1
for base, dirs, files in os.walk(target_dir):
   for file in files:
      fn = os.path.join(base, file)
      zip.write(fn, "LakePartner.gdb\\" + fn[rootlen:])
zip.write(OUTPUT_PATH + '\\LakePartner.msd', "LakePartner.msd")
zip.write(OUTPUT_PATH + '\\LakePartner.mxd', "LakePartner.mxd")
zip.write(OUTPUT_PATH + '\\readme_LakePartner.txt', "readme_LakePartner.txt")
zip.close()

# Remove the msd, mxd, readme.txt and file geodatabase. 
os.system("del " + OUTPUT_PATH + "\\LakePartner.msd")
os.system("del " + OUTPUT_PATH + "\\LakePartner.mxd")
os.system("del " + OUTPUT_PATH + "\\readme_LakePartner.txt")
os.system("rmdir " + OUTPUT_PATH + "\\LakePartner.gdb /s /q")


elapsed_time = time.time() - start_time
print elapsed_time
