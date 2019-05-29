#-------------------------------------------------------------------------------
# Name:        SAT Query Tool - Roadrunner version
# Purpose:     Calculates SAT efficiency and effectiveness scores for proposed project areas and writes to results spreadsheet
#               (with optional output as pdf report and spatial data)
#
# Author:      shannon.thol
#
# Created:     15/05/2019
# Copyright:   (c) shannon.thol 2019
# Licence:
#-------------------------------------------------------------------------------

#Load packages, etc.
import arcpy
print "Loading packages ..."
arcpy.AddMessage("Loading packages ...")

import shutil, re, sys
sys.path.append(r"D:\gisdata\Resources\conda\arcgispro-py3\Lib\site-packages")
from win32com import client
import xlsxwriter, os, csv, time, datetime, xlrd
import numpy as np
from arcpy import env
arcpy.env.overwriteOutput = True

start = time.time()

#Get user name (stripped of periods) and current date/time; concatenate to create runStamp for use in file names when saving
user = os.getenv('username').replace('.','').replace('-','').replace('_','')
currTime = ['0' + str(t) if t < 10 else str(t) for t in time.localtime()][0:5] #year=0, month=1, day=2, hour=3, min=4
fooTime = time.strftime("%b %d, %Y %I:%M %p", time.localtime()) #Mon Day, Year HH:MM AM/PM
runStamp = user + "_" + str(currTime[0]) + str(currTime[1]) + str(currTime [2]) + "_" + str(currTime[3]) + str(currTime[4]) #name_yyyymmdd_hhmm

#Set cell size environment, snapping environment, coordinate system as consistent with the 2011 NLCD grid
print "Setting environment parameters and locating reference data ..."
arcpy.AddMessage("Setting environment parameters ...")
coorSystem = arcpy.SpatialReference(5070) #5070 = WKID code for NAD_1983_Contiguous_USA_Albers

##############################################################################################################################################################################################################################################################
#Get paths of external data and files needed for analysis

#percentile rank master reference table for calculating efficiency scores
pctTable = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\PercentileTables_MASTER.xls"

#spp presence/absence csv table for use in terrestrial diversity calculations
sppData = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\hyperspp_vals.csv"

#30k polygon grids for identifying point feature class(es) to use in spatial intersection(s)
grid30K = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\SAT_30K_GRID.gdb\SAT_30K_GRID"

#gdb of sensitivity point feature classes for use in calculating intersection(s)
ptsGdb = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\SAT_30K_GRID_PTS.gdb"

#paths to shape png files for writing to results spreadsheet legend
circle = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\shape_circle.png"
diamond = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\shape_diamond.png"
square = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\shape_square.png"
star = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\shape_star.png"
triangle = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_Data4SATquerytool\\shape_triangle.png"

#paths to shape and results archives for saving a copy of the spatial input and results spreadsheet
shapeArchive = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_QueryToolArchive\\ShapeArchive.gdb"
resultsArchive = "D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_QueryToolArchive\\ResultsArchive"

############################################################################################################################################################################################################################################################
#get location of user's default geodatabase and set it as the scratch space
scratch = 'D:\\Users\\'+ os.getenv('username') + '\\Documents\\ArcGIS\\Default.gdb'

#set path to default SAT workspace
tempGdb = 'D:\\Users\\'+ os.getenv('username') + '\\Documents\\ArcGIS\\SATworkspace.gdb'

#if the SATworkspace gdb already exists, set is as the workspace
if arcpy.Exists(tempGdb):
    arcpy.env.workspace = tempGdb

#if the SATworkspace gdb doesn't already exist, create it and set it as the workspace
else:
    arcpy.CreateFileGDB_management('D:\\Users\\'+ os.getenv('username') + '\\Documents\\ArcGIS', 'SATworkspace.gdb')
    arcpy.env.workspace = tempGdb

#create empty list for storing paths of temp data that need to be deleted at end of run
toDelete = list()

##############################################################################################################################################################################################################################################################
#Get path for proposed project shapes and name of unique ID field from user input
projects = arcpy.GetParameterAsText(0)

#Get query name from user input
queEntName = arcpy.GetParameterAsText(1)
queName = re.sub('\W+','',queEntName.lower())

#Get name of unique numeric ID field (MUST BE INTEGER FIELD) from user input
projIDField = arcpy.GetParameterAsText(2)

#Get optional name of project "name" field from user input
projNameField = arcpy.GetParameterAsText(3)

#Get path for directory to write individual reports from user input
outDir = arcpy.GetParameterAsText(4)

#Get user input on whether they want the results written in individual reports for projects
report = arcpy.GetParameterAsText(5)

#Get user input on whether they want maps auto generated **Currently not operational!
maps = arcpy.GetParameterAsText(6)

#Get user input on whether they want the results in spatial format (dissolved feature class with attributes)
spatial = arcpy.GetParameterAsText(7)

############################################################################################################################################################################################################################################################
#Derive path for output excel spreadsheet and pdf report based on user supplied output directory, query name, and name/time run stamp
outPath = outDir + "\\SATResults_" + str(queName) + "_" + str(runStamp) + ".xlsx"

outReport = outDir + "\\SATReport_" + str(queName) + "_" + str(runStamp) + ".xlsx"

outPdf = outDir + "\\SATReport_" + str(queName) + "_" + str(runStamp) + ".pdf"

############################################################################################################################################################################################################################################################
#Attribute and LULC name reference data (UNLESS ATTRIBUTES ARE ADDED OR REMOVED FROM THE TOOL, OR OUTPUTS ARE CHANGED SIGNIFICANTLY, THESE SHOULD NOT CHANGE)

#fieldNames = {field name: attribute}
fieldNames = {'grwasu': 'groundwatersupply', 'aqudiv': 'aquaticdiversity', 'riflmi_bv': 'riverinefpfloodmitigation_bv', 'suwasu_bp': 'surfacewatersupply_bp', 'riflre': 'riverinelandscapefloodreduction', 'grwasu_bp': 'groundwatersupply_bp',
'grwasu_fn': 'groundwatersupply_fn', 'carsto': 'carbonstorage', 'tercliflo': 'terrestrialclimateflow', 'fldapr': 'flooddamageprevention', 'suwasu_fn': 'surfacewatersupply_fn', 'maaqre': 'marineaquaticrecreation', 'terres': 'terrestrialresilience',
'mastmi_fn': 'marinestormsurgemitigation_fn', 'suwasu': 'surfacewatersupply', 'frenpspre': 'freshwaternpsprevention', 'flohab': 'floodplainhabitat', 'mastmi_bp': 'marinestormsurgemitigation_bp', 'mastmi_bv': 'marinestormsurgemitigation_bv',
'riflmi_fn': 'riverinefpfloodmitigation_fn', 'lanuseint': 'landuseintensificationrisk', 'carseq': 'carbonsequestration', 'wethab': 'wetlandhabitat', 'marnpsred': 'marinenpsreduction', 'terrec_bv': 'terrestrialrecreation_bv',
'marnpre': 'marinenprevention', 'mastmi': 'marinestormsurgemitigation', 'riflre_bp': 'riverinelandscapefloodreduction_bp', 'pretidwet': 'presenttidalwetlands', 'riflre_bv': 'riverinelandscapefloodreduction_bv', 'heamit': 'heatmitigation',
'fldapr_fn': 'flooddamageprevention_fn', 'maaqre_fn': 'marineaquaticrecreation_fn', 'riflmi': 'riverinefpfloodmitigation', 'freflo': 'freshwaterprovision', 'ripfun': 'riparianfunction', 'fldapr_bp': 'flooddamageprevention_bp',
'strphyvar': 'streamphysicalvariety', 'fldapr_bv': 'flooddamageprevention_bv', 'frenpsmit': 'freshwaternpsmitigation', 'terhabqua': 'terrestrialhabitatquality', 'riflmi_bp': 'riverinefpfloodmitigation_bp',
'fraqre_bp': 'freshwateraquaticrecreation_bp', 'fraqre_bv': 'freshwateraquaticrecreation_bv', 'terrec_fn': 'terrestrialrecreation_fn', 'terrec': 'terrestrialrecreation', 'futtidwet': 'futuretidalwetlands', 'maaqre_bp': 'marineaquaticrecreation_bp',
'maaqre_bv': 'marineaquaticrecreation_bv', 'heamit_bv': 'heatmitigation_bv', 'shodyn': 'shorelinedynamics', 'heamit_bp': 'heatmitigation_bp', 'heamit_fn': 'heatmitigation_fn', 'fraqre': 'freshwateraquaticrecreation',
'riflre_fn': 'riverinelandscapefloodreduction_fn', 'flofun': 'floodplainfunction', 'strres': 'streamresilience', 'terrec_bp': 'terrestrialrecreation_bp', 'terphyvar': 'terrestrialphysicalvariety', 'fraqre_fn': 'freshwateraquaticrecreation_fn',
'tercon': 'terrestrialconnectivity'}

#fieldReverse = {attribute: field name}
fieldReverse = {'terrestrialrecreation_fn': 'terrec_fn', 'streamresilience': 'strres', 'marinenpsreduction': 'marnpsred', 'flooddamageprevention_bp': 'fldapr_bp', 'riverinelandscapefloodreduction_fn': 'riflre_fn', 'presenttidalwetlands': 'pretidwet',
'flooddamageprevention_fn': 'fldapr_fn', 'surfacewatersupply_fn': 'suwasu_fn', 'marineaquaticrecreation_bp': 'maaqre_bp', 'terrestrialrecreation_bv': 'terrec_bv', 'aquaticdiversity': 'aqudiv', 'terrestrialrecreation_bp': 'terrec_bp',
'terrestrialclimateflow': 'tercliflo', 'marineaquaticrecreation_bv': 'maaqre_bv', 'heatmitigation': 'heamit', 'groundwatersupply_fn': 'grwasu_fn', 'marinenprevention': 'marnpre', 'riverinefpfloodmitigation_bv': 'riflmi_bv', 'riverinefpfloodmitigation_bp': 'riflmi_bp',
'riverinefpfloodmitigation_fn': 'riflmi_fn', 'floodplainhabitat': 'flohab', 'marinestormsurgemitigation_bv': 'mastmi_bv', 'marinestormsurgemitigation_bp': 'mastmi_bp', 'freshwaternpsprevention': 'frenpspre', 'riverinelandscapefloodreduction': 'riflre',
'terrestrialresilience': 'terres', 'terrestrialrecreation': 'terrec', 'groundwatersupply_bp': 'grwasu_bp', 'marinestormsurgemitigation_fn': 'mastmi_fn', 'flooddamageprevention': 'fldapr', 'heatmitigation_fn': 'heamit_fn', 'terrestrialphysicalvariety': 'terphyvar',
'groundwatersupply': 'grwasu', 'riverinelandscapefloodreduction_bp': 'riflre_bp', 'riverinelandscapefloodreduction_bv': 'riflre_bv', 'freshwaterprovision': 'freflo', 'surfacewatersupply_bp': 'suwasu_bp', 'heatmitigation_bp': 'heamit_bp',
'terrestrialhabitatquality': 'terhabqua', 'marineaquaticrecreation': 'maaqre', 'heatmitigation_bv': 'heamit_bv', 'riparianfunction': 'ripfun', 'carbonsequestration': 'carseq', 'streamphysicalvariety': 'strphyvar', 'floodplainfunction': 'flofun',
'freshwateraquaticrecreation': 'fraqre', 'landuseintensificationrisk': 'lanuseint', 'flooddamageprevention_bv': 'fldapr_bv', 'freshwateraquaticrecreation_fn': 'fraqre_fn', 'riverinefpfloodmitigation': 'riflmi', 'marinestormsurgemitigation': 'mastmi',
'marineaquaticrecreation_fn': 'maaqre_fn', 'freshwaternpsmitigation': 'frenpsmit', 'shorelinedynamics': 'shodyn', 'futuretidalwetlands': 'futtidwet', 'carbonstorage': 'carsto', 'freshwateraquaticrecreation_bv': 'fraqre_bv', 'surfacewatersupply': 'suwasu',
'wetlandhabitat': 'wethab', 'freshwateraquaticrecreation_bp': 'fraqre_bp', 'terrestrialconnectivity': 'tercon', 'terrestrialdiversity': 'terdiv', 'terrestrialhabitatvariety':'terhabvar'}

#attDictionary = {attribute: long name}
attDictionary = {"aquaticdiversity":"Aquatic diversity", "carbonstorage":"Carbon storage", "carbonsequestration":"Carbon sequestration", "floodplainfunction":"Floodplain function", "floodplainhabitat":"Floodplain habitat",
"freshwaternpsmitigation":"Freshwater NPS mitigation", "freshwaternpsprevention":"Freshwater NPS prevention", "freshwaterprovision":"Freshwater flows", "futuretidalwetlands":"Future tidal wetlands", "marinenprevention":"Marine N prevention",
"marinenpsreduction":"Marine NPS reduction", "presenttidalwetlands":"Present tidal wetlands", "riparianfunction":"Riparian function", "shorelinedynamics":"Shoreline dynamics", "streamphysicalvariety":"Stream physical variety",
"streamresilience":"Stream resilience", "terrestrialclimateflow":"Terrestrial climate flow", "terrestrialconnectivity":"Terrestrial connectivity", "terrestrialdiversity":"Terrestrial diversity",
"terrestrialhabitatquality":"Terrestrial habitat quality", "terrestrialhabitatvariety":"Terrestrial habitat variety", "terrestrialphysicalvariety":"Terrestrial physical variety", "terrestrialresilience":"Terrestrial resilience",
"wetlandhabitat":"Wetland habitat", "heatmitigation": "Heat mitigation", "groundwatersupply": "Ground water supply", "surfacewatersupply": "Surface water supply", "freshwateraquaticrecreation": "Freshwater aquatic recreation",
"marineaquaticrecreation":"Marine aquatic recreation", "terrestrialrecreation":"Terrestrial recreation", "marinestormsurgemitigation": "Marine storm surge mitigation", "flooddamageprevention":"Flood damage prevention",
"riverinelandscapefloodreduction":"Riverine flood reduction", "riverinefpfloodmitigation": "Riverine flood mitigation", "landuseintensificationrisk": "Land use intensification"}

#classDict = {attribute: [Target system, Attribute group]}
classDict = {'aquaticdiversity': ['Freshwater','Diversity'],'streamphysicalvariety': ['Freshwater','Diversity'],'streamresilience': ['Freshwater','Resilience'],'floodplainfunction': ['Freshwater','Condition'],'riparianfunction': ['Freshwater','Condition'],
'floodplainhabitat': ['Freshwater','Habitat'],'wetlandhabitat': ['Freshwater','Habitat'],'freshwaterprovision': ['Freshwater','Condition'],'freshwaternpsmitigation': ['Freshwater','Condition'],'freshwaternpsprevention': ['Freshwater','Condition'],
'terrestrialdiversity': ['Terrestrial','Diversity'],'terrestrialhabitatvariety': ['Terrestrial','Diversity'],'terrestrialphysicalvariety': ['Terrestrial','Diversity'],'terrestrialhabitatquality': ['Terrestrial','Condition'],
'terrestrialclimateflow': ['Terrestrial','Condition'],'terrestrialconnectivity': ['Terrestrial','Condition'],'terrestrialresilience': ['Terrestrial','Resilience'],'carbonstorage': ['Terrestrial','Climate mitigation'],
'carbonsequestration': ['Terrestrial','Climate mitigation'],'presenttidalwetlands': ['Marine','Habitat'],'futuretidalwetlands': ['Marine','Habitat'],'shorelinedynamics': ['Marine','Condition'],'marinenprevention': ['Marine','Condition'],
'marinenpsreduction': ['Marine','Condition'],'heatmitigation': ['People','Temperature regulation'],'surfacewatersupply': ['People','Water supply'],'groundwatersupply': ['People','Water supply'],'marineaquaticrecreation': ['People','Recreation'],
'freshwateraquaticrecreation': ['People','Recreation'],'terrestrialrecreation': ['People','Recreation'],'marinestormsurgemitigation': ['People','Flooding'],'flooddamageprevention': ['People','Flooding'],'riverinelandscapefloodreduction': ['People','Flooding'],
'riverinefpfloodmitigation': ['People','Flooding'],'landuseintensificationrisk': ['Threat','Threat']}

#list of field names for creating summary method lists of intersection results
atts = [u'flofun', u'freflo', u'frenpsmit', u'frenpspre', u'ripfun', u'aqudiv', u'strphyvar', u'flohab', u'wethab', u'strres', u'carseq', u'carsto', u'tercliflo', u'tercon', u'terhabqua', u'terphyvar', u'terres', u'marnpre', u'marnpsred', u'shodyn',
u'futtidwet', u'pretidwet', u'fldapr', u'fldapr_fn', u'fldapr_bp', u'fldapr_bv', u'mastmi', u'mastmi_fn', u'mastmi_bp', u'mastmi_bv', u'riflmi', u'riflmi_fn', u'riflmi_bp', u'riflmi_bv', u'riflre', u'riflre_fn', u'riflre_bp', u'riflre_bv', u'fraqre',
u'fraqre_fn', u'fraqre_bp', u'fraqre_bv', u'maaqre', u'maaqre_fn', u'maaqre_bp', u'maaqre_bv', u'terrec', u'terrec_fn', u'terrec_bp', u'terrec_bv', u'heamit', u'heamit_fn', u'heamit_bp', u'heamit_bv', u'grwasu', u'grwasu_fn', u'grwasu_bp', u'suwasu',
u'suwasu_fn', u'suwasu_bp', u'lanuseint', u'terdivwgt', u'terdivsos', u'habvarwgt', u'habvarsos']

#tuple of land use/land cover class names
lulcTuple = (u'Water', u'Open Space Developed', u'Low Intensity Developed', u'Medium Intensity Developed', u'High Intensity Developed', u'Undetermined Developed', u'Pasture/Hay',u'Cultivated Crops', u'Undetermined Agriculture', u'Central Oak-Pine',
u'Undetermined Forest', u'Northern Hardwood-Conifer', u'Boreal Upland Forest', u'Ruderal Shrubland/Grassland', u'Undetermined Shrub/Grassland', u'Glade, Barren and Savanna', u'Large River Floodplain', u'Coastal Plain Swamp', u'Northern Swamp',
u'Northern Peatland', u'Wet Meadow/Shrub Marsh', u'Central Hardwood Swamp', u'Emergent Marsh',u'Undetermined Emergent Wetlands', u'Undetermined Woody Wetlands', u'Alpine', u'Cliff/Talus', u'Outcrop/Summit Scrub', u'Undetermined Barren',
u'Coastal Grassland/Shrubland', u'Coastal Plain Peatland', u'Coastal Plain Peat Swamp', u'Rocky Coast', u'Tidal Swamp', u'Tidal Marsh')

#lulcClasses = {class value in grid: lulc class}
lulcClasses = {3200: u'Undetermined Developed', 3800: u'Undetermined Agriculture', 900: u'Boreal Upland Forest', 1800: u'Ruderal Shrubland/Grassland', 11: u'Water', 400: u'Coastal Grassland/Shrubland', 2200: u'Coastal Plain Peatland',
1300: u'Northern Peatland', 21: u'Open Space Developed', 22: u'Low Intensity Developed', 23: u'Medium Intensity Developed', 24: u'High Intensity Developed', 800: u'Tidal Marsh', 1700: u'Central Hardwood Swamp', 1200: u'Rocky Coast',
3400: u'Undetermined Forest', 2100: u'Emergent Marsh', 700: u'Coastal Plain Swamp', 1600: u'Northern Hardwood-Conifer', 200: u'Outcrop/Summit Scrub', 1100: u'Cliff/Talus', 2000: u'Wet Meadow/Shrub Marsh', 81: u'Pasture/Hay', 850: u'Tidal Swamp',
600: u'Central Oak-Pine', 1500: u'Glade, Barren and Savanna', 3600: u'Undetermined Shrub/Grassland', 3300: u'Undetermined Barren', 3950: u'Undetermined Emergent Wetlands', 1000: u'Alpine', 3900: u'Undetermined Woody Wetlands', 1900: u'Northern Swamp',
82: u'Cultivated Crops', 750: u'Coastal Plain Peat Swamp', 1400: u'Large River Floodplain'}

###############################################################################################################################################################################################################################################################
#Reference data for writing results to excel spreadsheet and pdf report (UNLESS ATTRIBUTES ARE ADDED OR REMOVED FROM THE TOOL, OR OUTPUTS ARE CHANGED SIGNIFICANTLY, THESE SHOULD NOT CHANGE)

#meanColDictionary = {attribute: column number for mean results in excel spreadsheet}
meanColDictionary = {'riverinelandscapefloodreduction_bv': 43, 'terrestrialrecreation_fn': 53, 'streamresilience': 13, 'marinenpsreduction': 24, 'flooddamageprevention_bp': 30, 'surfacewatersupply_bv': 67, 'terrestrialphysicalvariety': 21,
'presenttidalwetlands': 27, 'flooddamageprevention_fn': 29, 'surfacewatersupply_fn': 65, 'riparianfunction': 8, 'terrestrialrecreation_bv': 55, 'aquaticdiversity': 9, 'riverinefpfloodmitigation': 36, 'terrestrialclimateflow': 16,
'marineaquaticrecreation_bv': 51, 'terrestrialhabitatvariety': 20, 'heatmitigation': 56, 'terrestrialhabitatquality': 18, 'floodplainhabitat': 11, 'marinestormsurgemitigation_bv': 35, 'marinestormsurgemitigation_bp': 34,
'riverinefpfloodmitigation_fn': 37, 'freshwaternpsprevention': 7, 'terrestrialrecreation_bp': 54, 'riverinelandscapefloodreduction': 40, 'groundwatersupply_fn': 61, 'freshwaternpsmitigation': 6, 'terrestrialrecreation': 52,
'groundwatersupply_bp': 62, 'marinestormsurgemitigation_fn': 33, 'terrestrialresilience': 22, 'flooddamageprevention': 28, 'heatmitigation_fn': 57, 'marinenprevention': 23, 'groundwatersupply': 60, 'riverinelandscapefloodreduction_bp': 42,
'freshwaterprovision': 5, 'surfacewatersupply_bp': 66, 'heatmitigation_bp': 58, 'marineaquaticrecreation': 48, 'heatmitigation_bv': 59, 'terrestrialdiversity': 19, 'carbonsequestration': 14, 'groundwatersupply_bv': 63, 'streamphysicalvariety': 10,
'floodplainfunction': 4, 'marineaquaticrecreation_bp': 50, 'freshwateraquaticrecreation': 44, 'landuseintensificationrisk': 3, 'flooddamageprevention_bv': 31, 'freshwateraquaticrecreation_fn': 45, 'marinestormsurgemitigation': 32,
'riverinefpfloodmitigation_bp': 38, 'marineaquaticrecreation_fn': 49, 'riverinefpfloodmitigation_bv': 39, 'shorelinedynamics': 25, 'futuretidalwetlands': 26, 'riverinelandscapefloodreduction_fn': 41, 'freshwateraquaticrecreation_bv': 47,
'surfacewatersupply': 64, 'wetlandhabitat': 12, 'carbonstorage': 15, 'freshwateraquaticrecreation_bp': 46, 'terrestrialconnectivity': 17}

#colDictionary = {attribute: column number for efficiency and effectiveness results in excel spreadsheet}
colDictionary = {'streamresilience': 13, 'marinenpsreduction': 24, 'presenttidalwetlands': 27, 'terrestrialdiversity': 19, 'riverinefpfloodmitigation': 30, 'terrestrialclimateflow': 16, 'terrestrialhabitatvariety': 20, 'heatmitigation': 35,
'terrestrialhabitatquality': 18, 'floodplainhabitat': 11, 'freshwaternpsprevention': 7, 'terrestrialresilience': 22, 'terrestrialphysicalvariety': 21, 'terrestrialrecreation': 34, 'landuseintensificationrisk': 3, 'flooddamageprevention': 28,
'marinenprevention': 23, 'groundwatersupply': 36, 'freshwaterprovision': 5, 'marineaquaticrecreation': 33, 'aquaticdiversity': 9, 'riparianfunction': 8, 'carbonsequestration': 14, 'streamphysicalvariety': 10, 'floodplainfunction': 4,
'freshwateraquaticrecreation': 32, 'marinestormsurgemitigation': 29, 'riverinelandscapefloodreduction': 31, 'freshwaternpsmitigation': 6, 'futuretidalwetlands': 26, 'carbonstorage': 15, 'surfacewatersupply': 37, 'wetlandhabitat': 12,
'shorelinedynamics': 25, 'terrestrialconnectivity': 17}

#rowDictionary = {attribute: output row in pdf report}
rowDictionary = {"floodplainfunction":9,"freshwaterprovision":10,"freshwaternpsmitigation":11, "freshwaternpsprevention":12,"riparianfunction":13, "aquaticdiversity":14, "streamphysicalvariety":15, "floodplainhabitat":16, "wetlandhabitat":17,
"streamresilience":18,"carbonsequestration":19, "carbonstorage":20,"terrestrialclimateflow":21, "terrestrialconnectivity":22, "terrestrialhabitatquality":23, "terrestrialdiversity":24, "terrestrialhabitatvariety":25, "terrestrialphysicalvariety":26,
"terrestrialresilience":27,"marinenprevention":28, "marinenpsreduction":29,"shorelinedynamics":30,"futuretidalwetlands":31,"presenttidalwetlands":32,"flooddamageprevention":38, "flooddamageprevention_fn":39, "flooddamageprevention_bp":40,
"flooddamageprevention_bv":41, "marinestormsurgemitigation": 42, "marinestormsurgemitigation_fn":43, "marinestormsurgemitigation_bp":44, "marinestormsurgemitigation_bv":45,"riverinefpfloodmitigation": 46, "riverinefpfloodmitigation_fn":47,
"riverinefpfloodmitigation_bp":48, "riverinefpfloodmitigation_bv":49,"riverinelandscapefloodreduction":50, "riverinelandscapefloodreduction_fn":51, "riverinelandscapefloodreduction_bp":52, "riverinelandscapefloodreduction_bv":53, "freshwateraquaticrecreation":54,
"freshwateraquaticrecreation_fn":55, "freshwateraquaticrecreation_bp":56, "freshwateraquaticrecreation_bv":57, "marineaquaticrecreation":58, "marineaquaticrecreation_fn":59, "marineaquaticrecreation_bp":60, "marineaquaticrecreation_bv":61,
"terrestrialrecreation":62, "terrestrialrecreation_fn":63, "terrestrialrecreation_bp":64, "terrestrialrecreation_bv":65, "heatmitigation":66, "heatmitigation_fn":67, "heatmitigation_bp":68, "heatmitigation_bv":69,
"groundwatersupply":70, "groundwatersupply_fn":71, "groundwatersupply_bp":72, "groundwatersupply_bv":73, "surfacewatersupply":74, "surfacewatersupply_fn":75,
"surfacewatersupply_bp":76, "surfacewatersupply_bv":77, "landuseintensificationrisk":81}

#mergeDictionary = {attribute: merge rows for efficiency and effectiveness results in pdf report}
mergeDictionary = {"flooddamageprevention":[38,41],"marinestormsurgemitigation":[42,45],"riverinefpfloodmitigation":[46,49],"riverinelandscapefloodreduction":[50,53],"freshwateraquaticrecreation":[54,57],"marineaquaticrecreation":[58,61],
"terrestrialrecreation":[62,65],"heatmitigation":[66,69], "groundwatersupply":[70,73],"surfacewatersupply":[74,77] }

#scatDict = {attribute: scatter plot table row number}
scatDict = {'floodplainfunction':40,'freshwaterprovision':41,'freshwaternpsmitigation':42,'freshwaternpsprevention':43,'riparianfunction':44, 'aquaticdiversity':45,'streamphysicalvariety':46, 'floodplainhabitat':47, 'wetlandhabitat':48, 'streamresilience':49,
'carbonsequestration':50, 'carbonstorage':51, 'terrestrialclimateflow':52, 'terrestrialconnectivity':53, 'terrestrialhabitatquality':54, 'terrestrialdiversity':55, 'terrestrialhabitatvariety':56, 'terrestrialphysicalvariety':57,
'terrestrialresilience':58,'marinenprevention':59, 'marinenpsreduction':60,'shorelinedynamics':61,'futuretidalwetlands':62,'presenttidalwetlands':63, 'flooddamageprevention':64, 'marinestormsurgemitigation':65, 'riverinefpfloodmitigation':66,
'riverinelandscapefloodreduction':67, 'freshwateraquaticrecreation':68, 'marineaquaticrecreation':69, 'terrestrialrecreation':70, 'heatmitigation':71, 'groundwatersupply':72, 'surfacewatersupply':73, 'landuseintensificationrisk':74}

#attShapes = {attribute: marker shape in scatter plots}
attShapes = {"aquaticdiversity":"circle", "carbonstorage":"square", "carbonsequestration":"square", "floodplainfunction":"circle", "floodplainhabitat":"circle","freshwaternpsmitigation":"circle", "freshwaternpsprevention":"circle",
"freshwaterprovision":"circle", "futuretidalwetlands":"triangle", "marinenprevention":"triangle","marinenpsreduction":"triangle", "presenttidalwetlands":"triangle", "riparianfunction":"circle", "shorelinedynamics":"triangle",
"streamphysicalvariety":"circle","streamresilience":"circle", "terrestrialclimateflow":"square", "terrestrialconnectivity":"square", "terrestrialdiversity":"square","terrestrialhabitatquality":"square", "terrestrialhabitatvariety":"square",
"terrestrialphysicalvariety":"square", "terrestrialresilience":"square","wetlandhabitat":"circle",'heatmitigation':'diamond', 'surfacewatersupply':'diamond','groundwatersupply':'diamond','marineaquaticrecreation':'diamond',
'freshwateraquaticrecreation':'diamond','terrestrialrecreation':'diamond','marinestormsurgemitigation':'diamond','flooddamageprevention':'diamond','riverinelandscapefloodreduction':'diamond','riverinefpfloodmitigation':'diamond', 'landuseintensificationrisk':'x'}

#radDict = {attribute: [efficiency radar plot column number, effectiveness radar plot column number]}
radDict = {'floodplainfunction':[1,36],'freshwaterprovision':[2,37],'freshwaternpsmitigation':[3,38],'freshwaternpsprevention':[4,39],'riparianfunction':[5,40],'aquaticdiversity':[6,41], 'streamphysicalvariety':[7,42], 'floodplainhabitat':[8,43],
'wetlandhabitat':[9,44], 'streamresilience':[10,45],'carbonsequestration':[11,46], 'carbonstorage':[12,47],'terrestrialclimateflow':[13,48],'terrestrialconnectivity':[14,49], 'terrestrialhabitatquality':[15,50], 'terrestrialdiversity':[16,51],
'terrestrialhabitatvariety':[17,52], 'terrestrialphysicalvariety':[18,53], 'terrestrialresilience':[19,54],'marinenprevention':[20,55],'marinenpsreduction':[21,56],'shorelinedynamics':[22,57],'futuretidalwetlands':[23,58],'presenttidalwetlands':[24,59],
'flooddamageprevention':[25,60], 'marinestormsurgemitigation':[26,61], 'riverinefpfloodmitigation':[27,62], 'riverinelandscapefloodreduction':[28,63], 'freshwateraquaticrecreation':[29,64], 'marineaquaticrecreation':[30,65],'terrestrialrecreation':[31,66],
'heatmitigation':[32,67], 'groundwatersupply':[33,68], 'surfacewatersupply':[34,69], 'landuseintensificationrisk':[35,70]}

#lulcDictionary = {lulc class: column number for results in excel spreadshet}
lulcDictionary = {u'Undetermined Forest': 13, u'Coastal Plain Peat Swamp': 34, u'Glade, Barren and Savanna': 18, u'Large River Floodplain': 19, u'Undetermined Woody Wetlands': 27, u'Coastal Plain Swamp': 20, u'Undetermined Emergent Wetlands': 26,
u'Northern Swamp': 21, u'Northern Peatland': 22, u'Cliff/Talus': 29, u'Water': 3, u'Undetermined Barren': 31, u'Open Space Developed': 4, u'Northern Hardwood-Conifer': 14, u'High Intensity Developed': 7, u'Undetermined Shrub/Grassland': 17,
u'Pasture/Hay': 9, u'Alpine': 28, u'Coastal Plain Peatland': 33, u'Rocky Coast': 35, u'Medium Intensity Developed': 6, u'Emergent Marsh': 25, u'Central Oak-Pine': 12, u'Undetermined Developed': 8, u'Wet Meadow/Shrub Marsh': 23,
u'Coastal Grassland/Shrubland': 32, u'Undetermined Agriculture': 11, u'Central Hardwood Swamp': 24, u'Boreal Upland Forest': 15, u'Low Intensity Developed': 5, u'Cultivated Crops': 10, u'Outcrop/Summit Scrub': 30, u'Ruderal Shrubland/Grassland': 16,
u'Tidal Swamp': 36, u'Tidal Marsh': 37}

###############################################################################################################################################################################################################################################################
#Reference information about data status and condition (LIKELY TO CHANGE MORE FREQUETLY THAN REFERENCE DATA IN ABOVE SECTIONS)

#list of limited scope attributes
scopeList = [u'riparianfunction', u'streamresilience', u'streamphysicalvariety', u'floodplainfunction', u'floodplainhabitat', u'wetlandhabitat', u'carbonstorage', u'terrestrialresilience', u'terrestrialclimateflow', u'terrestrialconnectivity',
u'futuretidalwetlands', u'marinenprevention', u'presenttidalwetlands', u'shorelinedynamics', u'groundwatersupply', u'surfacewatersupply', u'freshwateraquaticrecreation', u'marineaquaticrecreation', u'flooddamageprevention', u'flooddamageprevention_fn',
u'marinestormsurgemitigation_fn', u'marinestormsurgemitigation', u'heatmitigation', u'riverinefpfloodmitigation', u'flooddamageprevention_bp', u'flooddamageprevention_bv', u'marineaquaticrecreation_bp', u'marineaquaticrecreation_bv',
u'marineaquaticrecreation_fn', u'marinestormsurgemitigation_bp', u'marinestormsurgemitigation_bv']

#list of data that is under development
devList = ['riverinelandscapefloodreduction','riverinelandscapefloodreduction_fn','riverinelandscapefloodreduction_bp','riverinelandscapefloodreduction_bv','carbonsequestration', 'terrestrialclimateflow']

###############################################################################################################################################################################################################################################################
#Reference information for calculating attribute effectiveness scores (MUST BE UPDATED ANYTIME A SENSITIVITY GRID CHANGES USING THE FOLLOWING SCRIPT:
#D:\gisdata\Projects\Regional\ConservationDimensions\ZonalStatsTool\ZonalStatsTool_workingfiles\CalculatingScalars_Sums_030119.py)

#scalarDict = {attribute: scalar for calculating effectiveness scores (observed maximums for 1,000 acre neighborhood sums)}
scalarDict = {u'streamresilience': 90260.0, u'marinenpsreduction': 73164.58, u'presenttidalwetlands': 76556.25, u'terrestrialdiversity': 154683.95, u'terrestrialclimateflow': 71731.98, u'terrestrialhabitatvariety': 6090.58, u'heatmitigation': 61331.74,
u'terrestrialhabitatquality': 90260.0, u'marinenprevention': 88748.0, u'floodplainhabitat': 88660.0, u'freshwaternpsprevention': 90164.35, u'terrestrialresilience': 82228.46, u'terrestrialrecreation': 68479.18, u'landuseintensificationrisk': 46758.96,
u'flooddamageprevention': 52075.22, u'terrestrialphysicalvariety': 84863.46, u'groundwatersupply': 87906.1, u'freshwaterprovision': 90260.0, u'marineaquaticrecreation': 60027.43, u'aquaticdiversity': 84232.26, u'riparianfunction': 66173.13,
u'streamphysicalvariety': 83387.74, u'floodplainfunction': 88660.0, u'freshwateraquaticrecreation': 55585.03, u'marinestormsurgemitigation': 46368.62, u'riverinelandscapefloodreduction': 55867.92, u'freshwaternpsmitigation': 90260.0,
u'futuretidalwetlands': 71602.0, u'carbonstorage': 64877.66, u'surfacewatersupply': 65891.88, u'wetlandhabitat': 90214.87, u'shorelinedynamics': 41931.13, u'terrestrialconnectivity': 83698.2, u'riverinefpfloodmitigation': 51832.42}

###############################################################################################################################################################################################################################################################
#Prepare proposed project shapes and attributes
print "Preparing vector data ..."
arcpy.AddMessage("Preparing vector data ...")

#Test to see if input project shapes are a polygon feature class, if not, print error message and terminate the script
if arcpy.Describe(projects).shapeType != 'Polygon':
    print "FAILED FEATURE CLASS TYPE CHECK: You submitted a " + arcpy.Describe(projects).shapeType + " feature class, terminating analysis ... Please try again using a polygon feature class."
    arcpy.AddMessage("FAILED FEATURE CLASS TYPE CHECK: You submitted a " + arcpy.Describe(projects).shapeType + " feature class, terminating analysis ... Please try again using a polygon feature class.")
    sys.exit(0)

#Test to see if input project shapes have a spatial reference defined, if not, print error message and terminate the script
if arcpy.Describe(projects).spatialReference.name == 'Unknown':
    print "FAILED SPATIAL REFERENCE CHECK: Project feature class does not have a spatial reference defined, terminating analysis ... Please fix the problem and try again."
    arcpy.AddMessage("FAILED SPATIAL REFERENCE CHECK: Project feature class does not have a spatial reference defined, terminating analysis ... Please fix the problem and try again.")
    sys.exit(0)

##try:
#Dissolve the input shapes on the projIDField field, and reproject the input shapes to use NAD_1983_Contiguous_USA_Albers coordinate system (standard for SAT)
projDiss = arcpy.Dissolve_management(projects, "in_memory\\projectsdiss", projIDField, [[projNameField, 'FIRST']])

projWork = arcpy.Project_management(projDiss, scratch + "\\" + str(queName) + "_" + str(runStamp) + "_allprojects", coorSystem)
arcpy.Delete_management(projDiss)
toDelete.append(projWork)

#If a field named "SAT_Hectares" already exists in the project attribute table, overwrite the values with a new geometry calculation
if 'SAT_Hectares' in [n.name for n in arcpy.ListFields(projWork)]:
    arcpy.CalculateField_management(projWork, "SAT_Hectares", "!SHAPE.AREA@HECTARES!", "PYTHON")
#If a field named "SAT_Hectares" doesn't already exists in the project attribute table, proceed with adding it and calculating area
else:
    arcpy.AddField_management(projWork, "SAT_Hectares", "DOUBLE")
    arcpy.CalculateField_management(projWork, "SAT_Hectares", "!SHAPE.AREA@HECTARES!", "PYTHON")

#Create empty dictionaries for storing results
projDict = dict() #projDict = {projID: [projName, outRow, vector hectares, raster hectares, number cells] ...}
projResults = dict() #projResults = {projID: {attribute: [mean sensitivity, efficiency, effectiveness, efficiency*effectiveness], attribute: [mean sensitivity, efficiency, effectiveness, efficiency*effectiveness] ...} ...}
lulcResults = dict() #lulcResults = {projID: {lulcclass: percent, lulcclass: percent ...} ...}
divwgtResults = dict() #divwgtResults = {projID: [mean terrestrial diversity weight, mean habitat variety weight] ...}

#Create lists for storing size flagged projects (don't meet the minimum or maximum size thresholds)
smallList = list()
bigList = list()

#Set counter for storing project row position (outRow) for writing results to excel spreadsheet
outRow = 3
print "Checking project polygons ..."
arcpy.AddMessage("Checking project polygons ...")

#Iterate through the projects in the dissolved project feature class and retrieve the IDs and vector area in hectares
with arcpy.da.UpdateCursor(projWork,[projIDField,'FIRST_' + str(projNameField),"SAT_Hectares"]) as cursor:
    for row in cursor:
        #check project size and add items to small or big list as appropriate
        if row[2] < 1:
            print "   FAILED CHECKS: Project " + str(row[0]) + " is smaller than the MINIMUM size threshold (1 Ha) and will not be analyzed ..."
            arcpy.AddMessage("   FAILED CHECKS: Project " + str(row[0]) + " is smaller than the MINIMUM size threshold (1 Ha) and will not be analyzed ...")
            cursor.deleteRow()
            smallList.append([row[0], row[1]])
        elif row[2] > 24600:
            print "   FAILED CHECK: Project " + str(row[0]) + " is larger than the MAXIMUM size threshold (24,600 Ha) and will not be analyzed ..."
            arcpy.AddMessage("   FAILED CHECK: Project " + str(row[0]) + " is larger than the MAXIMUM size threshold (24,600 Ha) and will not be analyzed ...")
            cursor.deleteRow()
            bigList.append([row[0], row[1]])
        else:
            #populate projDict information for current project polygon
            projDict[row[0]] = ['']*5 #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells] ...}
            projDict[row[0]][0] = str(row[1]) #postion 0 in projDict = projName
            projDict[row[0]][1] = outRow #position 1 in projDict = outRow
            projDict[row[0]][2] = round(float(row[2]),2) #position 2 in projDict = vectorhectares

            #prepare projResults dictionary for current project polygon (to populate later)
            projResults[row[0]] = dict()
            for rowAtt in rowDictionary:
                if rowAtt[-3:] in ['_fn', '_bp', '_bv']:
                    projResults[row[0]][rowAtt] = ['']*1
                else:
                    projResults[row[0]][rowAtt] = ['']*4

            #prepare lulcResults dictionary for current project polygon (to populate later)
            lulcResults[row[0]] = dict()
            for rowLulc in lulcTuple:
                lulcResults[row[0]][rowLulc] = 0.00

            #prepare diversity weights dictionary for current project polygon (to populate later)
            divwgtResults[row[0]] = ['']*2

            #advance outRow by 1
            outRow += 1

#check number of features remaining in projWork, if no features remain after removing too big and too small projects, terminate analysis
if int(arcpy.GetCount_management(projWork)[0]) == 0:
    print "   FAILED CHECKS: No projects met the size requirements, terminating analysis ..."
    arcpy.AddMessage("   FAILED CHECKS: No projects met the size requirements, terminating analysis ...")
    arcpy.Delete_management(projWork)
    sys.exit(0)

##except Exception:
##    arcpy.Delete_management(projWork)
##    print "There was a problem preparing the vector data, please check the data and try again ..."
##    arcpy.AddMessage("There was a problem preparing the vector data, please check the data and try again ...")
##    e = sys.exc_info()[1]
##    print(e.args[0])
##    arcpy.AddError(e.args[0])
##    sys.exit(0)

#############################################################################################################################################################################################################################################################
#Create dictionary of standardized percentiles for all attributes
print "Importing reference data, please wait ..."
arcpy.AddMessage("Importing reference data, please wait ...")

try:
    #open standardized percentile master workbook for SAT attributes as inBook
    inBook = xlrd.open_workbook(pctTable)
    #Get list of sheet names in inBook
    sheetNames = inBook.sheet_names()

    #Create empty dictionary for storing standardized percentile data
    #bookDict = {attribute: [sens, 0.45pctrnk, 1.17pctrnk, ...], attribute: [sens, 0.45pctrnk, 1.17pctrnk, ...], ...}
    bookDict = dict()
    #iterate through sheets in standardized percentile master workbook and add them to the book dictionary
    for sheet in sheetNames:
        inSheet = inBook.sheet_by_name(sheet)
        sheetName = inSheet.name

        valsList = list()
        #iterate through rows in current sheet and add row values to a new vals list and then to the book dictionary
        for row in range(inSheet.nrows):
            rowVals = inSheet.row_values(row)
            valsList.append(rowVals)

        bookDict[sheet] = valsList

    #retrieve size list and largest and smallest sizes for reference
    sizeList = [x for x in bookDict['aquaticdiversity'] if x[0] == "Sens"][0]
    lgstSize = max(sizeList)
    smstSize = min(sizeList)

    #create new empty list for storing spp presence/absence data read from the csv table
    sppList = list()
    #iterate through the lines from the sppData file and add them to the sppList
    with open(sppData) as sppFile:
        reader = csv.reader(sppFile)
        for row in reader:
            sppList.append(row)

    #convert the sppList data into an array of presence/absence values for spp
    sppArray = np.asarray(sppList, dtype = int)

except Exception:
    arcpy.Delete_management(projWork)
    print "There was a problem importing the reference data, please try again ..."
    arcpy.AddMessage("There was a problem importing the reference data, please try again ...")
    e = sys.exc_info()[1]
    print(e.args[0])
    arcpy.AddError(e.args[0])
    sys.exit(0)

##########################################################################################################################################################################################################################################################
#Begin analysis to summarize SAT data in project extents and calculate efficiency and effectiveness scores
print "Finding project polygon extents ..."
arcpy.AddMessage("Finding project polygon extents ...")
try:
    #Identify which input point datasets need to be used in the intersections by running an intersection with the 30k grid polygon features
    #Specify path for saving the grid intersect results
    gridInt = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_gridintersect"

    #Calculate intersection of points with the 30K grid and get list of all grid cells the full project extent intersects
    arcpy.Intersect_analysis([grid30K, projWork], gridInt)
    intList = list()

    with arcpy.da.SearchCursor(gridInt, "GRID_ID") as cursor:
        for row in cursor:
            intList.append(row[0])

    #Delete the grid intersect results
    arcpy.Delete_management(scratch + "\\" + str(queName) + "_" + str(runStamp) + "_gridintersect")

    #Get list of unique grid cell IDs that the full project extent intersects
    gridList = list(set(intList))

except Exception:
    print "There was a problem getting the vector data extent, please try again ..."
    arcpy.AddMessage("There was a problem getting the vector data extent, please try again ...")
    e = sys.exc_info()[1]
    print(e.args[0])
    arcpy.AddError(e.args[0])
    sys.exit(0)

##############################################################################################################################################################################################################################################################
#create new empty list for storing effectiveness results for all attributes and all projects (for determining the scale of the scatter plot x-axes)
effectScores = list()

##try:
#Calculate intersections between project polygons and appropriate SAT point grid(s) and retrieve project stats
#If gridList has more than one item in it, iterate through grids in gridList calculating intersection with each appropriate point grid
print "Computing spatial intersection ..."
arcpy.AddMessage("Computing spatial intersection ...")
if len(gridList) > 1:
    #Define mergeList for storing paths of individual intersection results that need to get merged
    mergeList = list()

    #Iterate through grids in the gridList
    for grid in gridList:
        #Get path for current point data
        currPts = ptsGdb + "\\" + grid + "_pts"

        #Calculate intersection between projects and current points
        currInt = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersect_" + grid
        toDelete.append(currInt)
        mergeList.append(currInt)
        arcpy.Intersect_analysis([projWork, currPts],currInt)

    #Merge the individual intersection results
    intersect = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersect"
    toDelete.append(intersect)
    arcpy.Merge_management(mergeList, intersect)

    #Clean up individual intersection results
    for intItem in mergeList:
        arcpy.Delete_management(intItem)

#If gridList has only one item, proceed with a single intersect and summary step
else:
    #Get path for current point data
    pts = ptsGdb + "\\" + gridList[0] + "_pts"

    #Calculate intersection between projects and points
    intersect = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersect"
    toDelete.append(intersect)
    arcpy.Intersect_analysis([projWork, pts], intersect)

#Calculate attribute means by supplied projID, excluding the terrestrial diversity and habitat variety effectiveness references
finalMeans = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersectmeans_FINAL"
toDelete.append(finalMeans)
arcpy.Statistics_analysis(intersect, finalMeans, [[a,'MEAN'] for a in atts if a not in [u'terdivsos', u'habvarsos']], projIDField)

#Calculate attribute sums by supplied projID, excluding the terrestrial diversity and habitat variety weights, and all people attribute components
finalSums = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersectsums_FINAL"
toDelete.append(finalSums)
arcpy.Statistics_analysis(intersect, finalSums, [[a,'SUM'] for a in atts if a not in [u'terdivwgt', u'habvarwgt'] and a[-3:] not in ['_fn','_bp','_bv']], projIDField)

#Calculate lulc class counts by supplied projID
finalCounts = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersectcounts_FINAL"
toDelete.append(finalCounts)
arcpy.Statistics_analysis(intersect, finalCounts, [['POINTID','COUNT']], [projIDField,'satlulc'])

#Retrieve stats from the finalMeans table and add estimate of raster hectares to projDict (raster hectares = (number points in intersection*900)/10000)#projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}
#and add project means and efficiency scores to projResults #projResults = {projID: {attname: [meansens, efficiency, effectiveness, efficiency*effectiveness], attname: [meansens, efficiency, effectiveness, efficiency*effectiveness] ...} ...}
print "Calculating mean sensitivities and efficiency scores ..."
arcpy.AddMessage("Calculating mean sensitivities and efficiency scores ...")
#get list of field names in the finalMeans results to use below
meanFields = [f.name for f in arcpy.ListFields(finalMeans) if f.name != 'OBJECTID']
with arcpy.da.SearchCursor(finalMeans, meanFields) as cursor:
    for row in cursor:
        #get current project id and project raster size in hectares as (number of points in intersection * 900)/10000
        projID = row[0] #projID is stored in first field (position 0)
        projSize = round((float(row[1])*900.0)/10000.0,2) #number of points in intersection is stored in second field (position 1)

        #add raster hectares and cell counts to projDict
        projDict[projID][3] = projSize
        projDict[projID][4] = int(row[1]) #number of points in intersection is stored in second field (position 1)

        #add water supply beneficiary vulnerability 'no data' entries to projResults dictionary (because we lack people vulnerability data for these two attributes)
        projResults[projID]['surfacewatersupply_bv'][0] = 'No Data'
        projResults[projID]['groundwatersupply_bv'][0] = 'No Data'

        #iterate through remaining items in row using the meanCounter, starting with index position 2 (positions 0 and 1 are projID and number of point, respectively)
        meanCounter = 2
        while meanCounter < len(meanFields):

            #get name of current field
            currField = meanFields[meanCounter].replace('MEAN_','')

            #if current attribute is a diversity weight, proceed with writing the results to the diversity weight dictionary (have to handle differently because efficiency is not calculated until later)
            if currField == u'terdivwgt':
                projAttMean = float(row[meanCounter])/1000000.0 #scalar because weight data had to be multiplied by 10^6 before summarizing to points to maintain adequate level of precision
                divwgtResults[projID][0] = projAttMean #divwgtResults = {projID: [terrdivwgt, habvarwgt], projID: [terrdivwgt, habvarwgt] ...}

            elif currField == u'habvarwgt':
                projAttMean = float(row[meanCounter])/100.0 #scalar because weight data had to be multiplied by 10^2 before summarizing to points to maintain adequate level of precision
                divwgtResults[projID][1] = projAttMean #divwgtResults = {projID: [terrdivwgt, habvarwgt], projID: [terrdivwgt, habvarwgt] ...}

            #test whether currAtt is under development
            elif fieldNames[currField] in devList:
                #test whether current under development data is a people component, if yes populate results dictionary with single 'Under development' value
                if fieldNames[currField][-3:] in['_fn','_bp','_bv']:
                    projResults[projID][fieldNames[currField]] = ["Under development"]
                #if no, populate results dictionary with list of four 'Under development' values
                else:
                    projResults[projID][fieldNames[currField]] = ['Under development']*4

            #if current value is NOT a diversity weight but IS a nonetype (no value, which indicates out of scope), proceed with writing "Under development and "Null" as appropriate
            elif row[meanCounter] is None:

                #test whether currAtt is under development
                if fieldNames[currField] in devList:
                    #test whether current under development data is a people component, if yes populate results dictionary with single 'Under development' value
                    if fieldNames[currField][-3:] in['_fn','_bp','_bv']:
                        projResults[projID][fieldNames[currField]] = ["Under development"]
                    #if no, populate results dictionary with list of four 'Under development' values
                    else:
                        projResults[projID][fieldNames[currField]] = ['Under development']*4

                #test whether currAtt has limited scope
                elif fieldNames[currField] in scopeList:
                    #test whether current limited scope data is a people component, if yes populate results dictionary with single 'Null' value
                    if fieldNames[currField][-3:] in ['_fn','_bp','_bv']:
                        projResults[projID][fieldNames[currField]] = ["Null"]
                    #if no, populate results dictionary with list of two 'Null' values (only two because out of scope areas get assigned 0 values for effectiveness and effectiveness*efficiency scores)
                    else:
                        projResults[projID][fieldNames[currField]][0] = 'Null'
                        projResults[projID][fieldNames[currField]][1] = 'Null'


            #if current value is NOT a nonetype, but IS a people component proceed with writing just mean result to the results dictionary (efficiency and effectiveness are calculated for sensitivity grids only, not people components)
            elif fieldNames[currField][-3:] in ['_fn', '_bp', '_bv']:
                projAttMean = round(float(row[meanCounter])/100.0,2)
                projResults[projID][fieldNames[currField]] = [projAttMean]

            #if current value is NOT a diversity weight, is NOT under development, is NOT a nonetype, and is NOT a people component, proceed with calculating efficiency and effectiveness scores and writing to results dictionary
            else:
                #retrieve project attribute mean and write it to the results dictionary
                projAttMean = round(float(row[meanCounter])/100.0,2)
                projResults[projID][fieldNames[currField]][0] = projAttMean

                #get sensitivity efficiency reference data for current attribute
                attList = bookDict[fieldNames[currField]]

                #if current project size in sizes list, get corresponding percentile from pct list and assign it as currPct
                if projSize in sizeList:
                    projSizeIn = sizeList.index(projSize)
                    currPct =[x[projSizeIn] for x in attList if x[0] == projAttMean][0]

                #if current project size is not in sizes list, find nearest values and extrapolate between them
                else:
                    #calculate absolute difference between project size and standard sizes and get minimum absolute difference
                    absDiff = [round(abs(projSize - x),2) for x in sizeList[1:]]
                    minAd = min(absDiff)
                    minAdIn = absDiff.index(minAd) + 1  #add 1 because absDiff list excludes attList index 1 "Sens" text value
                    minAdSize = sizeList[minAdIn]
                    minAdPct = [x[minAdIn] for x in attList if x[0] == projAttMean][0]

                    #if min absolute difference size is less than project size, extrapolate based on plus 1 index value
                    if minAdSize < projSize:
                        onePlsSize = sizeList[minAdIn + 1]
                        onePlsPct = [x[minAdIn + 1] for x in attList if x[0] == projAttMean][0]
                        slope = (onePlsPct - minAdPct)/(onePlsSize - minAdSize)
                        currPct = round(slope*(projSize - minAdSize) + minAdPct,2)

                    #else if min absolute difference size is greater than project size, extrapolate based on minus 1 index value
                    elif minAdSize > projSize:
                        oneMinSize = sizeList[minAdIn - 1]
                        oneMinPct = [x[minAdIn - 1] for x in attList if x[0] == projAttMean][0]
                        slope = (minAdPct - oneMinPct)/(minAdSize - oneMinSize)
                        currPct = round(minAdPct - slope*(minAdSize - projSize),2)

                #write efficiency score to the results dictionary
                projResults[projID][fieldNames[currField]][1] = currPct

            #advance meanCounter by 1
            meanCounter += 1

#########################################################################################################################################################################################################################################
#Retrieve stats from the finalSums table and add effectiveness scores to projResults #projResults = {projID: {attname: [meansens, efficiency, effectiveness, efficiency*effectiveness], attname: [meansens, efficiency, effectiveness, efficiency*effectiveness] ...} ...}
print "Calculating effectiveness scores ..."
arcpy.AddMessage("Calculating effectiveness scores ...")
#get list of field names in the finalSums results to use below
sumFields = [f.name for f in arcpy.ListFields(finalSums) if f.name != 'OBJECTID']
with arcpy.da.SearchCursor (finalSums, sumFields) as cursor:
    for row in cursor:
        #get current project id
        projID = row[0]

        #iterate through remaining items in row using the sumCounter, starting with index position 2 (because positions 0 and 1 are projID and intersection count, respectively)
        sumCounter = 2
        while sumCounter < len(sumFields):
            #get name of current field
            currField = sumFields[sumCounter].replace('SUM_','')

            #if current sum is for the habitat variety effectiveness score, retrieve scalar and calculate current effectiveness score (have to handle differently because efficiency is not calculated until later, so can't calculate efficiency*effectiveness yet)
            if currField == 'habvarsos':
                #retrieve scalar for terrestrial habitat variety and calculate effectiveness score
                currScalar = scalarDict['terrestrialhabitatvariety']
                currSum = float(row[sumCounter])/100.0
                currEffect = round((currSum/currScalar),3)

                #add current effectiveness score to effectiveness score list (for determining scale of x axis in graphs)
                effectScores.append(currEffect)

                #write effectiveness score for current project to results dictionary
                projResults[projID]['terrestrialhabitatvariety'][2] = currEffect

            #if current sum is for the terrestrial diversity effectiveness score, retrieve scalar and calculate current effectiveness score (have to handle differently because efficiency is not calculated until later, so can't calculate efficiency*effectiveness yet)
            elif currField == 'terdivsos':
                #retrieve scalar for terrestrial diversity and calculate effectiveness score
                currScalar = scalarDict['terrestrialdiversity']
                currSum = float(row[sumCounter])/100.0
                currEffect = round((currSum/currScalar),3)

                #add current effectiveness score to effectiveness score list (for determining scale of x acis in graphs)
                effectScores.append(currEffect)

                #write effectiveness score for current project to results dictionary
                projResults[projID]['terrestrialdiversity'][2] = currEffect

            #if current sum is for an attribute under development, pass ("Under development" has already been written to the projResults dictionary)
            elif fieldNames[currField] in devList:
                pass

            #if current value is nonetype (no value, which indicates out of scope), write 0s in the effectiveness and efficiency*effectivness scores
            elif row[sumCounter] is None:
                projResults[projID][fieldNames[currField]][2] = 0
                projResults[projID][fieldNames[currField]][3] = 0

            else:
                #retrieve scalar for current attribute and calculate current effectiveness score
                currScalar = scalarDict[fieldNames[currField]]
                currEffect = round((float(row[sumCounter])/100.0)/currScalar,3)

                #add current effectiveness score to effectiveness score list (for determining scale of x axis in graphs)
                effectScores.append(currEffect)

                #retrieve current efficiency score from results dictionary
                currEffic = projResults[projID][fieldNames[currField]][1]

                #write effectiveness and efficiency*effectiveness scores for current project to results dictionary
                projResults[projID][fieldNames[currField]][2] = currEffect
                projResults[projID][fieldNames[currField]][3] = round(currEffect*currEffic,2)

            #advance sumCounter by 1
            sumCounter += 1

##except Exception:
##    print "There was a problem calculating scores, please try again ..."
##    arcpy.AddMessage("There was a problem calculating scores, please try again ...")
##    e = sys.exc_info()[1]
##    print(e.args[0])
##    arcpy.AddError(e.args[0])
##    sys.exit(0)
############################################################################################################################################################################################################################################################
#Calculate scores for terrestrial diversity attribute
print "Calculating scores for terrestrial diversity ..."
arcpy.AddMessage("Calculating scores for terrestrial diversity ...")

#create new empty sppDict dictionary #sppDict = projID: [[sppcode, count], [sppcode, count] ...]], projID: [[sppcode, count], [sppcode, count] ...]] ...}
sppDict = dict()
#Create new empty list for saving unique spp codes for all projects
sppList = list()

#iterate through the species fields and calculate counts by species
for spp in [u'spp1', u'spp2', u'spp3', u'spp4', u'spp5']:
    finalSppCounts = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersectcounts_FINAL_" + spp
    toDelete.append(finalSppCounts)
    arcpy.Statistics_analysis(intersect, finalSppCounts, [['POINTID','COUNT']], [projIDField,spp])

    #define field names for search cursor
    sppFields = [projIDField, spp, 'COUNT_POINTID']

    #iterate through rows in the final spp counts table, retreive the results and write to the sppDict
    with arcpy.da.SearchCursor(finalSppCounts, sppFields) as cursor:
        for row in cursor:
            projID = row[0]
            #if current spp code is already in the species list, pass
            if row[1] in sppList:
                pass
            #else, add current spp code to species list
            else:
                sppList.append(row[1])

            #if current project ID is already in the sppDict, retrieve the list of values, append new data, and write results back to sppDict
            if projID in sppDict:
                currVals = sppDict[projID]
                currVals.append([row[1],row[2]])
                sppDict[projID] = currVals
            #if current project ID is not already in the sppDict, write the new results to sppDict
            else:
                sppDict[projID] = [[row[1],row[2]]]

#get sensitivity efficiency reference data for terrestrial diversity attribute
attList = bookDict['terrestrialdiversity']

#iterate through projects in the sppDict results, calculating mean sensitivity and efficiency scores
for projID in sppDict:

    #retrieve size of current project (raster hectares)
    projSize = projDict[projID][3]

    #retrieve spp values and counts for current projects
    values = [x[0] for x in sppDict[projID]]
    valCounts = [x[1] for x in sppDict[projID]]

    #create empty array of 0s (1 row X 441 columns) for using as base on which to append pixel values
    pixelArray = np.zeros((1,441),dtype=np.int)

    #iterate through values getting current index position, current count, and current spp from array; calculate number of pixels as product of current spp array and current count
    for val in values:
        currPosition = values.index(val)
        ##overallPosition = sppList.index(val)  #NOTE: this line is a legacy from other approach tried for faster import of spp data (only import rows for spp that are represented in the current projects) that I couldn't get to work, but we may want to revisit
        currCount = int(valCounts[currPosition])
        currSpp = sppArray[val]
        numPixels = np.resize(np.array(currCount*currSpp),(1,441))

        #Append the numPixel values by species to the pixelArray
        pixelArray = np.concatenate((pixelArray,numPixels))

    #compute the n term as the sum of columns in the pixelArray, the nxn-1 term, and the summed nxn-1 term
    nterm = pixelArray.sum(axis=0).astype(long)
    nxn_1 = (pixelArray.sum(axis=0))*(pixelArray.sum(axis=0)-1).astype(long)
    summednxn_1 = sum(nxn_1).astype(long)

    #compute N term as the sum of all values in the nterm array, and then calculate the NxN-1 term
    Nterm = sum(nterm).astype(long)
    NxN_1 = Nterm * (Nterm-1)

    #Compute Simpson's Index of Diversity and terrestrial diversity degree
    simpsons = 1.0-(float(summednxn_1)/float(NxN_1))
    terrdivDgr = simpsons*10.0

    #retrieve mean terrestrial diversity weight for current project from diversity weights dictionary
    terrdivWgt = divwgtResults[projID][0]  #divwgtResults = {projID: [terrdivwgt, habvarwgt], projID: [terrdivwgt, habvarwgt] ...}

    #calculate mean sensitivity for current project and write to projResults dictionary
    projAttMean = round(terrdivDgr*terrdivWgt,2)
    projResults[projID]['terrestrialdiversity'][0] = projAttMean

    #if current project size in sizes list, get corresponding percentile from pct list and assign it as currPct
    if projSize in sizeList:
        projSizeIn = sizeList.index(projSize)
        currPct =[x[projSizeIn] for x in attList if x[0] == projAttMean][0]

    #if current project size is not in sizes list, find nearest values and extrapolate between them
    else:
        #calculate absolute difference between project size and standard sizes and get minimum absolute difference
        absDiff = [round(abs(projSize - x),2) for x in sizeList[1:]]
        minAd = min(absDiff)
        minAdIn = absDiff.index(minAd) + 1  #add 1 because absDiff list excludes attList index 1 "Sens" text value
        minAdSize = sizeList[minAdIn]
        minAdPct = [x[minAdIn] for x in attList if x[0] == projAttMean][0]

        #if min absolute difference size is less than project size, extrapolate based on plus 1 index
        if minAdSize < projSize:
            onePlsSize = sizeList[minAdIn + 1]
            onePlsPct = [x[minAdIn + 1] for x in attList if x[0] == projAttMean][0]
            slope = (onePlsPct - minAdPct)/(onePlsSize - minAdSize)
            currPct = round(slope*(projSize - minAdSize) + minAdPct,2)

        #else if min absolute difference size is greater than project size, extrapolate based on minus 1 index
        elif minAdSize > projSize:
            oneMinSize = sizeList[minAdIn - 1]
            oneMinPct = [x[minAdIn - 1] for x in attList if x[0] == projAttMean][0]
            slope = (minAdPct - oneMinPct)/(minAdSize - oneMinSize)
            currPct = round(minAdPct - slope*(minAdSize - projSize),2)

    #write efficiency score to the results dictionary
    projResults[projID]['terrestrialdiversity'][1] = currPct

    #retrieve terrestrial diversity effectiveness score for current project and write efficiency*effectiveness for current project to the results dictionary
    currEffect = projResults[projID]['terrestrialdiversity'][2]
    projResults[projID]['terrestrialdiversity'][3] = round(currPct*currEffect,2)

##############################################################################################################################################################################################################################################
#Calculate scores for terrestrial habitat variety

print "Calculating values for terrestrial habitat variety ..."
arcpy.AddMessage("Calculating values for terrestrial habitat variety ...")

#Get percentile rank reference lists for terrestrial habitat variety attribute from bookDict
attList = bookDict['terrestrialhabitatvariety']

#Calculate habitat counts by project
finalHabCounts = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersectcounts_FINAL_hab"
toDelete.append(finalHabCounts)
arcpy.Statistics_analysis(intersect, finalHabCounts, [['POINTID','COUNT']], [projIDField,'habvarlulc'])

#create empty habDict dictionary #habDict = projID: [[habcode, count], [habcode, count] ...]], projID: [[habcode, count], [habcode, count] ...]] ...}
habDict = dict()

#define field names for search cursor
habFields =[projIDField, 'habvarlulc', 'COUNT_POINTID']

#iterate through rows in the final habitat counts table, retrieve list of values, append new data, and write results back to dictionary
with arcpy.da.SearchCursor(finalHabCounts, habFields) as cursor:
    for row in cursor:
        projID = row[0]

        #if current project ID is already in the habDict, retrieve the list of values, append new data, and write results back to habDict (NOTE: this check is probably not needed - carried over from spp count process, which I used as the basis for this code)
        if projID in habDict:
            currVals = habDict[projID]
            currVals.append([row[1],row[2]])
            habDict[projID] = currVals

        #if current project ID is not already in the sppDict, write the new results to sppDict
        else:
            habDict[projID] = [[row[1],row[2]]]

#iterate through projects in the habDict results, calculating mean sensitivity and efficiency scores
for projID in habDict:

    #retrieve size of current project (raster hectares)
    projSize = projDict[projID][3]

    #retrieve habitat values and counts for current project
    values = [x[0] for x in habDict[projID]]
    nValues = [x[1] for x in habDict[projID]]

    #calculate n*(n-1) values for all habitat types
    nxn_1Values = [n*(n-1) for n in nValues]

    #calculate sum of n*(n-1) term
    nxn_1 = sum(nxn_1Values)

    #calculate N term as sum of n values
    NValue= sum(nValues)
    NxN_1 = NValue*(NValue-1)

    #calculate simpson's index of diversity for habitats
    habitatD = 1.0-(float(nxn_1)/float(NxN_1))
    habvarDgr = habitatD*20.0

    #retrieve mean habitat variety weight for current project from diversity weights dictionary
    habvarWgt= divwgtResults[projID][1]  #divwgtResults = {projID: [terrdivwgt, habvarwgt], projID: [terrdivwgt, habvarwgt] ...}

    #calculate mean sensitivity for current project and write to projResults dictionary
    projAttMean = round(habvarDgr*habvarWgt,2)
    projResults[projID]['terrestrialhabitatvariety'][0] = projAttMean

    #if current project size in sizes list, get corresponding percentile from pct list and assign it as currPct
    if projSize in sizeList:
        projSizeIn = sizeList.index(projSize)
        currPct =[x[projSizeIn] for x in attList if x[0] == projAttMean][0]

    #if current project size is not in sizes list, find nearest values and extrapolate between them
    else:
        #calculate absolute difference between project size and standard sizes and get minimum absolute difference
        absDiff = [round(abs(projSize - x),2) for x in sizeList[1:]]
        minAd = min(absDiff)
        minAdIn = absDiff.index(minAd) + 1  #add 1 because absDiff list excludes attList index 1 "Sens" text value
        minAdSize = sizeList[minAdIn]
        minAdPct = [x[minAdIn] for x in attList if x[0] == projAttMean][0]

        #if min absolute difference size is less than project size, extrapolate based on plus 1 index
        if minAdSize < projSize:
            onePlsSize = sizeList[minAdIn + 1]
            onePlsPct = [x[minAdIn + 1] for x in attList if x[0] == projAttMean][0]
            slope = (onePlsPct - minAdPct)/(onePlsSize - minAdSize)
            currPct = round(slope*(projSize - minAdSize) + minAdPct,2)

        #else if min absolute difference size is greater than project size, extrapolate based on minus 1 index
        elif minAdSize > projSize:
            oneMinSize = sizeList[minAdIn - 1]
            oneMinPct = [x[minAdIn - 1] for x in attList if x[0] == projAttMean][0]
            slope = (minAdPct - oneMinPct)/(minAdSize - oneMinSize)
            currPct = round(minAdPct - slope*(minAdSize - projSize),2)

    #write efficiency score to the results dictionary
    projResults[projID]['terrestrialhabitatvariety'][1] = currPct
    #retrieve terrestrial habitat variety effectiveness score for current project and write efficiency*effectiveness for current project to the results dictionary
    currEfect = projResults[projID]['terrestrialhabitatvariety'][2]
    projResults[projID]['terrestrialhabitatvariety'][3] = round(currPct*currEffect,2)

##################################################################################################################################################################################################################
#Retrieve lulc stats from the finalCounts table and add to the lulcResults dictionary  #lulcResults =  {projID: {lulcclass: percent, lulcclass: percent ...} ...}
print "Calculating land use/land cover percentages in projects extents ..."

#get path of final lulc counts table
finalCounts = scratch + "\\" + str(queName) + "_" + str(runStamp) + "_intersectcounts_FINAL"

#define field names for search cursor
countFields =[projIDField, 'satlulc', 'COUNT_POINTID']

#iterate through rows in finalCounts table retrieving lulc class and number of cells, and writing to lulcResults dictionary
with arcpy.da.SearchCursor(finalCounts, countFields) as cursor:
    for row in cursor:
        projID = row[0]
        currlulc = row[1]
        lulcName = lulcClasses[currlulc]

        #retrieve number of cells in current project extent from projDict, calculate percent and write to lulcResults dictionary
        totCells = float(projDict[row[0]][4]) #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}
        lulcResults[projID][lulcName] = round((float(row[2])/totCells)*100.0,2)

#################################################################################################################################################################################################################
#Check number of projects, and if greater than 23, pull out top scoring "efficiency*effectiveness" projects for inclusion in the scatterplot and radar plots
#retrieve number of projects in projResults
numProj = len(projResults)

#if numProj is less than or equal to 23, pass
if numProj <= 23:
    pass
#if numProj is more than 23, proceed with identifying the top scoring projects to be included on the scatter plot
else:
    #create empty list for storing top project scores
    projTopScores = [] #[max efficiency*effectiveness score across all attributes, projID]

    #iterate through projects getting list of the results
    for proj in projResults:
        results = projResults[proj]

        #create list for storing efficiency*effectiveness scores (product) for each attribute
        prodResults = list()

        #iterate through results in the results list
        for result in results:
            #skip the people component results
            if result[-2:] in ['bp','bv','fn']:
                pass
            #skip the landuseintensificationrisk result
            elif result == 'landuseintensificationrisk':
                pass
            else:
                #else, get list of results for the current attribute
                resultVals = results[result]
                #skip if the product of scores is 'Under development'
                if resultVals[3] == 'Under development':
                    pass
                #else, append the current product of scores to the prodResults list
                else:
                    prodResults.append(resultVals[3])

        #get the max value in prodResults and add that and current projID the projTopScores list
        maxResult = max(prodResults)
        projTopScores.append([maxResult,proj])

    #sort the projTopScores list in descending order by product scores and pull out the top 23 scoring projects
    projTopScores.sort(key=lambda x: x[0],reverse=1) #descending order [product score, projID]
    topProjs = projTopScores[:23]

    #create a list of project ids that have the same product score value as what was originally identified as project 23 (to check for ties in the 23rd position)
    toCheck = [d[1] for d in projTopScores if d[0] == topProjs[22][0]]

    #if there are ties, proceed with getting the sizes (raster area) of all tied projects and adding to a list
    if len(toCheck) >1:
        tieList = list() #[raster area, proj ID]
        for check in toCheck:
            tieList.append([projDict[check][3],check]) #{projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}

        #get the projID for largest project in tieList
        tieList.sort(key=lambda y: y[0],reverse=1) #descending order on project size [project size, projID]
        num23 = tieList[0][1]

        #if the projID for largest project in tieList is the same as what's already in the topProjs list, pass
        if num23 == topProjs[22][1]:
            pass
        #else, remove the original 23rd listing and insert the new 23rd listing (largest of the tied projects)
        else:
            topProjs.remove(topProjs[22])
            topProjs.insert(22,[c for c in projTopScores if c[1]==num23][0])
            topProjs.insert(22,tieList[0])
    else:
        pass


##################################################################################################################################################################################################################
#Set up Excel spreadsheet for writing results
print "Writing results to Excel spreadsheet ..."
arcpy.AddMessage("Writing results to Excel spreadsheet ...")

#create new workbook with xlswriter
workbook = xlsxwriter.Workbook(outPath)

#define workbook colors
dblue = '#004dce'
blue = '#0071c6'
green = '#00b252'
ygreen = 'effb94'
yellow = '#ffff9c'
grey = '#bdbebd'
dgrey = '#808080'

#define workbook style formats
boldForm = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 10})
lgboldForm = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 11})
italForm = workbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 10})
italralForm = workbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 10, 'align': 'right'})
italwrapForm = workbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 10, 'text_wrap':True, 'valign': 'vcenter'})
bordForm = workbook.add_format({'border': True, 'font_name': 'Calibri', 'font_size': 10, 'text_wrap': True, 'align': 'center'})
lalBordForm = workbook.add_format({'border': True, 'font_name': 'Calibri', 'font_size': 10, 'align': 'left'})
regForm = workbook.add_format({'font_name': 'Calibri', 'font_size': 10})
colForm = workbook.add_format({'font_name': 'Calibri', 'font_size': 10, 'bg_color': dgrey, 'font_color': 'white', 'bold': True, 'align': 'center', 'border': True, 'text_wrap': True})
colitalForm = workbook.add_format({'font_name': 'Calibri', 'font_size': 10, 'bg_color': dgrey, 'font_color': 'white', 'bold': True, 'align': 'center', 'border': True, 'text_wrap': True, 'italic': True})
mergeForm = workbook.add_format({'italic':True, 'align': 'right', 'font_name': 'Calibri', 'font_size': 10, 'border': True})
dblueForm = workbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':dblue})
blueForm = workbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':blue})
greenForm = workbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':green})
ygreenForm = workbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':ygreen})
yellForm = workbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':yellow})
greyForm = workbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':grey})
greycentForm = workbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':grey,'align':'center'})
whiteForm = workbook.add_format({'font_name':'Calibri','font_size':10,'align':'center','bg_color':'white'})
redForm = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 10, 'font_color': 'red'})

#add new worksheet for writing information about the query run and write template cells
querysheet = workbook.add_worksheet("QueryInformation")
querysheet.set_column(0,0,14.0)
querysheetList = [[0,0,'SAT Project Query Results',boldForm],[1,0,'NOTE: Results are subject to field verification. Scores are calculated as summaries for whole project areas - see maps to evaluate sub-project variation.',boldForm],[3,0,'Query name:',italForm],[4,0,'Submitted by:',italForm],[5,0,'Generated on:',italForm],[6,0,'Input file:',italForm],[7,0,'Unique ID field:',italForm],[8,0,'Name field:',italForm]]

for queryItem in querysheetList:
    querysheet.write(queryItem[0],queryItem[1],queryItem[2],queryItem[3])

#write user-supplied query information to query sheet
querysheet.write(3,1,queName,regForm)
querysheet.write(4,1,user,regForm)
querysheet.write(5,1,fooTime,regForm)
querysheet.write(6,1,projects,regForm)
querysheet.write(7,1,projIDField,regForm)
querysheet.write(8,1,projNameField,regForm)

b = 11
#if there were any projects that didn't meet the size thresholds, print these in the querysheet under query information
if len(bigList) > 0:
    querysheet.write(b,1,"FAILED CHECKS: The following project(s) are larger than the MAXIMUM size threshold (24,600 Ha) and were not analyzed:",redForm)
    for big in bigList:
        b+=1
        bigID = str(big[0])
        bigName = str(big[1])
        querysheet.write(b,0,"Project",regForm)
        querysheet.write(b,1,bigID + "; " + bigName,regForm)

if len(smallList) > 0:
    if b == 11:
        querysheet.write(b,1,"FAILED CHECKS: The following project(s) are smaller than the MINIMUM size threshold (1 Ha) and were not analyzed:",redForm)
    else:
        b+=1
        querysheet.write(b,1,"FAILED CHECKS: The following project(s) are smaller than the MINIMUM size threshold (1 Ha) and were not analyzed:",redForm)

    for small in smallList:
        b+=1
        smallID = str(small[0])
        smallName = str(small[1])
        querysheet.write(b,0,"Project",regForm)
        querysheet.write(b,1,smallID + "; " + smallName,regForm)


#add new worksheet and write reference information
refsheet = workbook.add_worksheet("References")
refsheet.set_column(0,7,11.0) #start column index, end column index, width
refsheetList = [[0,0,'Landcover percents reflect the land cover composition of each project based on the SAT landcover dataset (hybrid of NLCD and NETWHC datasets).',boldForm],[1,0,'Range: 0-100',italForm],[5,0,'Mean scores reflect average attribute senstivities in project areas (plus average function, beneficiary population, and beneficiary vulnerability values for people attributes).',boldForm],[6,0,'Range: 0-20 (Null: project falls outside scope of places that provide the function)',italForm],[10,0,'Efficiency scores reflect the percent rank of the project as compared to other similarly-sized areas across the state.',boldForm],[11,1,'Score',italForm],[11,2,'Interpretation',italForm],[12,1,'100',blueForm],[13,1,'50',greenForm],[14,1,'0',yellForm],[15,1,'Null',greyForm],[16,0,'Colors in the results spreadsheet and maps have the same directionality (yellow = low scores, blue = high scores), but the specific score/color classifications may not match exactly due to differences in scale of the calculations.',redForm],[19,0,'Effectiveness scores reflect how much of the attribute statewide total values are captured by projects, scaled so high value 1,000 acre (~400 Ha) projects have a score of 1.',boldForm],[20,1,'Score',italForm],[20,2,'Interpretation',italForm],[21,1,'>100',dblueForm],[22,1,'100',blueForm],[23,1,'10',greenForm],[24,1,'1',ygreenForm],[25,1,'0.1',yellForm],[26,0,'Colors in the results spreadsheet and maps have the same directionality (yellow = low scores, blue = high scores), but the specific score/color classifications may not match exactly due to differences in scale of the calculations.',redForm]]

for refItem in refsheetList:
    refsheet.write(refItem[0],refItem[1],refItem[2],refItem[3])

refsheetmergeList = [[12,0,15,0,'Efficiency score key',italwrapForm],[12,2,12,6,'project score is >= 100% of similarly sized areas across the state',blueForm],[13,2,13,6,'project score is >= 50% of similarly sized areas across the state',greenForm],[14,2,14,6,'project score is >= 0% of similarly sized areas across the state',yellForm],[15,2,15,6,'project is outside scope that currently provides function and/or service',greyForm],[21,0,25,0,'Effectiveness score key',italwrapForm],[21,2,21,6,'only achievable by projects larger than ~40,000 Ha',dblueForm],[22,2,22,6,'highest attainable score for projects of ~40,000 Ha',blueForm],[23,2,23,6,'highest attainable score for projects of ~4,000 Ha',greenForm],[24,2,24,6,'highest attainable score for projects of ~400 Ha',ygreenForm],[25,2,25,6,'highest attainable score for projects of ~40 Ha',yellForm],[28,0,28,1,'Attribute name',boldForm],[28,2,28,11,'Description of function/service/threat mapped',boldForm],[29,0,29,1,'Floodplain function',lalBordForm],[29,2,29,11,'Supporting current and future functional floodplains.',lalBordForm],[30,0,30,1,'Freshwater flows',lalBordForm],[30,2,30,11,'Supporting ecologically sufficient flows.',lalBordForm],[31,0,31,1,'Freshwater NPS mitigation',lalBordForm],[31,2,31,11,'Mitigating NPS pollution.',lalBordForm],[32,0,32,1,'Freshwater NPS prevention',lalBordForm],[32,2,32,11,'Preventing new sources of NPS pollution.',lalBordForm],[33,0,33,1,'Riparian function',lalBordForm],[33,2,33,11,'Providing material inputs and shading to in-stream habitats.',lalBordForm],[34,0,34,1,'Aquatic diversity',lalBordForm],[34,2,34,11,'Supporting aquatic habitat for conservation species.',lalBordForm],[35,0,35,1,'Stream physical variety',lalBordForm],[35,2,35,11,'Supporting the full geophysical variety of stream habitats.',lalBordForm],[36,0,36,1,'Floodplain habitat',lalBordForm],[36,2,36,11,'Providing floodplain habitat.',lalBordForm],[37,0,37,1,'Wetland habitat',lalBordForm],[37,2,37,11,'Providing freshwater wetland habitat.',lalBordForm],[38,0,38,1,'Stream resilience',lalBordForm],[38,2,38,11,'Supporting high resilience streams.',lalBordForm],[39,0,39,1,'Carbon sequestration',lalBordForm],[39,2,39,11,'Carbon sequestration rate, contributing to climate change mitigation.',lalBordForm],[40,0,40,1,'Carbon storage',lalBordForm],[40,2,40,11,'Storage of carbon, preventing increases in atmospheric carbon.',lalBordForm],[41,0,41,1,'Terrestrial climate flow',lalBordForm],[41,2,41,11,'Supporting range shifts and species movement with climate change.',lalBordForm],[42,0,42,1,'Terrestrial connectivity',lalBordForm],[42,2,42,11,'Supporting current-day movement of wildlife.',lalBordForm],[43,0,43,1,'Terrestrial habitat quality',lalBordForm],[43,2,43,11,'Preventing fragmentation and habitat degredation.',lalBordForm],[44,0,44,1,'Terrestrial diversity',lalBordForm],[44,2,44,11,'Provision of habitat for conservation species.',lalBordForm],[45,0,45,1,'Terrestrial habitat variety',lalBordForm],[45,2,45,11,'Lands with high habitat heterogeneity contributing to diversity and resilience.',lalBordForm],[46,0,46,1,'Terrestrial physical variety',lalBordForm],[46,2,46,11,'Providing the full variety of geophysical habitats.',lalBordForm],[47,0,47,1,'Terrestrial resilience',lalBordForm],[47,2,47,11,'Supporting climate change adaptation.',lalBordForm],[48,0,48,1,'Marine N prevention',lalBordForm],[48,2,48,11,'Preventing new nitrogen inputs to the marine system.',lalBordForm],[49,0,49,1,'Marine NPS reduction',lalBordForm],[49,2,49,11,'Lands preventing and mitigating NPS pollution to marine ecosystems.',lalBordForm],[50,0,50,1,'Shoreline dynamics',lalBordForm],[50,2,50,11,'Supporting shoreline dynamics.',lalBordForm],[51,0,51,1,'Future tidal wetlands',lalBordForm],[51,2,51,11,'Providing and supporting movement of tidal wetlands with climate change.',lalBordForm],[52,0,52,1,'Present tidal wetlands',lalBordForm],[52,2,52,11,'Providing and supporting current tidal wetlands.',lalBordForm],[53,0,53,1,'Flood damage prevention',lalBordForm],[53,2,53,11,'Reducing flood damage by excluding infrastructure from flood-prone areas.',lalBordForm],[54,0,54,1,'Marine storm surge mitigation',lalBordForm],[54,2,54,11,'Reducing coastal flood damage by capturing /slowing floodwaters and reducing flood extent and depth.',lalBordForm],[55,0,55,1,'Riverine flood mitigation',lalBordForm],[55,2,55,11,'Mitigating riverine flood damage by capturing/slowing floodwaters in the floodplain and lower peak flood stages.',lalBordForm],[56,0,56,1,'Riverine flood reduction',lalBordForm],[56,2,56,11,'Reducing riverine flood damage by capturing stormwater runoff throughout the watershed and reducing storm-related streamflow.',lalBordForm],[57,0,57,1,'Freshwater aquatic recreation',lalBordForm],[57,2,57,11,'Lands contributing to the condition of freshwaterways used for swimming, boating, and fishing.',lalBordForm],[58,0,58,1,'Marine aquatic recreation',lalBordForm],[58,2,58,11,'Lands contributing to the condition of marine waterways used for swimming, boating, and fishing.',lalBordForm],[59,0,59,1,'Terrestrial recreation',lalBordForm],[59,2,59,11,'Providing the potential for hiking and scenery or wildlife viewing.',lalBordForm],[60,0,60,1,'Heat mitigation',lalBordForm],[60,2,60,11,'Supporting temperature regulation through shade production and transpiration/reduced albedo.',lalBordForm],[61,0,61,1,'Ground water supply',lalBordForm],[61,2,61,11,'Lands providing purification and recharge of groundwater for public use.',lalBordForm],[62,0,62,1,'Surface water supply',lalBordForm],[62,2,62,11,'Lands supporting clean and dependable surface water supplies for public use.',lalBordForm],[63,0,63,1,'Land use intensification',lalBordForm],[63,2,63,11,'Threat of new or intensified anthropogenic land uses (developed and/or agricultural).',lalBordForm]]

for refmerge in refsheetmergeList:
    refsheet.merge_range(refmerge[0],refmerge[1],refmerge[2],refmerge[3],refmerge[4],refmerge[5])

#add new worksheet for writing land cover percent results and write template cells
lcsheet = workbook.add_worksheet("LandcoverPercents")
lcsheet.set_column(0,1,11.0)
lcsheet.set_column(2,36,13.57)
lcsheetList = [[0,2,'Land cover group>',mergeForm],[0,3,'Water',bordForm],[0,4,'Developed',bordForm],[0,5,'Developed',bordForm],[0,6,'Developed',bordForm],[0,7,'Developed',bordForm],[0,8,'Developed',bordForm],[0,9,'Agricultural',bordForm],[0,10,'Agricultural',bordForm],[0,11,'Agricultural',bordForm],[0,12,'Forest',bordForm],[0,13,'Forest',bordForm],[0,14,'Forest',bordForm],[0,15,'Forest',bordForm],[0,16,'Shrub/grassland',bordForm],[0,17,'Shrub/grassland',bordForm],[0,18,'Shrub/grassland',bordForm],[0,19,'Wetland',bordForm],[0,20,'Wetland',bordForm],[0,21,'Wetland',bordForm],[0,22,'Wetland',bordForm],[0,23,'Wetland',bordForm],[0,24,'Wetland',bordForm],[0,25,'Wetland',bordForm],[0,26,'Wetland',bordForm],[0,27,'Wetland',bordForm],[0,28,'Barren',bordForm],[0,29,'Barren',bordForm],[0,30,'Barren',bordForm],[0,31,'Barren',bordForm],[0,32,'Coastal upland',bordForm],[0,33,'Coastal wetland',bordForm],[0,34,'Coastal wetland',bordForm],[0,35,'Coastal upland',bordForm],[0,36,'Coastal wetland',bordForm],[0,37,'Coastal wetland',bordForm],[1,0,'Project Name',bordForm],[1,1,'Project ID',bordForm],[1,2,'Project area (Ha) *GIS calc.',bordForm],[1,3,'% Water',colForm],[1,4,'% Open Space Developed',colForm],[1,5,'% Low Intensity Developed',colForm],[1,6,'% Medium Intensity Developed',colForm],[1,7,'% High Intensity Developed',colForm],[1,8,'% Undetermined Developed',colForm],[1,9,'% Pasture/Hay',colForm],[1,10,'% Cultivated Crops',colForm],[1,11,'% Undetermined Agriculture',colForm],[1,12,'% Central Oak-Pine',colForm],[1,13,'% Undetermined Forest',colForm],[1,14,'% Northern Hardwood-Conifer',colForm],[1,15,'% Boreal Upland Forest',colForm],[1,16,'% Ruderal Shrubland / Grassland',colForm],[1,17,'% Undetermined Shrub / Grassland',colForm],[1,18,'% Glade, Barren and Savanna',colForm],[1,19,'% Large River Floodplain',colForm],[1,20,'% Coastal Plain Swamp',colForm],[1,21,'% Northern Swamp',colForm],[1,22,'% Northern Peatland',colForm],[1,23,'% Wet Meadow/Shrub Marsh',colForm],[1,24,'% Central Hardwood Swamp',colForm],[1,25,'% Emergent Marsh',colForm],[1,26,'% Undetermined Emergent Wetlands',colForm],[1,27,'% Undetermined Woody Wetlands',colForm],[1,28,'% Alpine',colForm],[1,29,'% Cliff/Talus',colForm],[1,30,'% Outcrop / Summit Scrub',colForm],[1,31,'% Undetermined Barren',colForm],[1,32,'% Coastal Grassland / Shrubland',colForm],[1,33,'% Coastal Plain Peatland',colForm],[1,34,'% Coastal Plain Peat Swamp',colForm],[1,35,'% Rocky Coast',colForm],[1,36,'% Tidal Swamp',colForm],[1,37,'% Tidal Marsh',colForm]]

for lcItem in lcsheetList:
    lcsheet.write(lcItem[0], lcItem[1], lcItem[2], lcItem[3])

#write lulc results to lulc sheet lulcResults = {projID: {lulcclass, pct, lulcclass, pct ...} ...}
for lulcResult in lulcResults:
    currRow = projDict[lulcResult][1]-1 #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}
    currName = projDict[lulcResult][0]
    lcsheet.write(currRow,0,currName,regForm) #projName
    lcsheet.write(currRow,1,lulcResult,regForm) #projID
    currVecHa = projDict[lulcResult][2] #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}
    lcsheet.write(currRow,2,currVecHa,regForm)
    lulcList = lulcResults[lulcResult]

    for lulc in lulcDictionary:
        lulcCol = lulcDictionary[lulc]
        currlulcPct = lulcList[lulc]
        lcsheet.write(currRow,lulcCol,currlulcPct,regForm)

#add new worksheet for writing mean score results and write template cells
meansheet = workbook.add_worksheet("MeanScores")

meansheet.set_column(0,2,11.0) #start column index, end column index, width
meansheet.set_column(3,67,11.71)
meansheetList = [[0,2,'Target System>',italralForm],[1,2,'Attribute Group>',italralForm],[0,4,'Freshwater',bordForm],[0,5,'Freshwater',bordForm],[0,6,'Freshwater',bordForm],[0,7,'Freshwater',bordForm],[0,8,'Freshwater',bordForm],[0,9,'Freshwater',bordForm],[0,10,'Freshwater',bordForm],[0,11,'Freshwater',bordForm],[0,12,'Freshwater',bordForm],[0,13,'Freshwater',bordForm],[0,14,'Terrestrial',bordForm],[0,15,'Terrestrial',bordForm],[0,16,'Terrestrial',bordForm],[0,17,'Terrestrial',bordForm],[0,18,'Terrestrial',bordForm],[0,19,'Terrestrial',bordForm],[0,20,'Terrestrial',bordForm],[0,21,'Terrestrial',bordForm],[0,22,'Terrestrial',bordForm],[0,23,'Marine',bordForm],[0,24,'Marine',bordForm],[0,25,'Marine',bordForm],[0,26,'Marine',bordForm],[0,27,'Marine',bordForm],[1,3,'Threat',bordForm],[1,4,'Condition',bordForm],[1,5,'Condition',bordForm],[1,6,'Condition',bordForm],[1,7,'Condition',bordForm],[1,8,'Condition',bordForm],[1,9,'Diversity',bordForm],[1,10,'Diversity',bordForm],[1,11,'Habitat',bordForm],[1,12,'Habitat',bordForm],[1,13,'Resilience',bordForm],[1,14,'Climate mitigation',bordForm],[1,15,'Climate mitigation',bordForm],[1,16,'Condition',bordForm],[1,17,'Condition',bordForm],[1,18,'Condition',bordForm],[1,19,'Diversity',bordForm],[1,20,'Diversity',bordForm],[1,21,'Diversity',bordForm],[1,22,'Reslience',bordForm],[1,23,'Condition',bordForm],[1,24,'Condition',bordForm],[1,25,'Condition',bordForm],[1,26,'Habitat',bordForm],[1,27,'Habitat',bordForm],[2,0,'Project Name',bordForm],[2,1,'Project ID',bordForm],[2,2,'Project area (Ha) *GIS calc.',bordForm],[2,3,'Land use intensification',colForm],[2,4,'Floodplain function',colForm],[2,5,'Freshwater flows',colForm],[2,6,'Freshwater NPS mitigation',colForm],[2,7,'Freshwater NPS prevention',colForm],[2,8,'Riparian function',colForm],[2,9,'Aquatic diversity',colForm],[2,10,'Stream physical variety',colForm],[2,11,'Floodplain habitat',colForm],[2,12,'Wetland habitat',colForm],[2,13,'Stream resilience',colForm],[2,14,'Carbon sequestration',colForm],[2,15,'Carbon storage',colForm],[2,16,'Terrestrial climate flow',colForm],[2,17,'Terrestrial connectivity',colForm],[2,18,'Terrestrial habitat quality',colForm],[2,19,'Terrestrial diversity',colForm],[2,20,'Terrestrial habitat variety',colForm],[2,21,'Terrestrial physical variety',colForm],[2,22,'Terrestrial resilience',colForm],[2,23,'Marine N prevention',colForm],[2,24,'Marine NPS reduction',colForm],[2,25,'Shoreline dynamics',colForm],[2,26,'Future tidal wetlands',colForm],[2,27,'Present tidal wetlands',colForm],[2,28,'Flood damage prevention',colForm],[2,29,'Function',colitalForm],[2,30,'Beneficiary population',colitalForm],[2,31,'Beneficiary vulnerability',colitalForm],[2,32,'Marine storm surge mitigation',colForm],[2,33,'Function',colitalForm],[2,34,'Beneficiary population',colitalForm],[2,35,'Beneficiary vulnerability',colitalForm],[2,36,'Riverine flood mitigation',colForm],[2,37,'Function',colitalForm],[2,38,'Beneficiary population',colitalForm],[2,39,'Beneficiary vulnerability',colitalForm],[2,40,'Riverine flood reduction',colForm],[2,41,'Function',colitalForm],[2,42,'Beneficiary population',colitalForm],[2,43,'Beneficiary vulnerability',colitalForm],[2,44,'Freshwater aquatic recreation',colForm],[2,45,'Function',colitalForm],[2,46,'Beneficiary population',colitalForm],[2,47,'Beneficiary vulnerability',colitalForm],[2,48,'Marine aquatic recreation',colForm],[2,49,'Function',colitalForm],[2,50,'Beneficiary population',colitalForm],[2,51,'Beneficiary vulnerability',colitalForm],[2,52,'Terrestrial recreation',colForm],[2,53,'Function',colitalForm],[2,54,'Beneficiary population',colitalForm],[2,55,'Beneficiary vulnerability',colitalForm],[2,56,'Heat mitigation',colForm],[2,57,'Function',colitalForm],[2,58,'Beneficiary population',colitalForm],[2,59,'Beneficiary vulnerability',colitalForm],[2,60,'Ground water supply',colForm],[2,61,'Function',colitalForm],[2,62,'Beneficiary population',colitalForm],[2,63,'Beneficiary vulnerability',colitalForm],[2,64,'Surface water supply',colForm],[2,65,'Function',colitalForm],[2,66,'Beneficiary population',colitalForm],[2,67,'Beneficiary vulnerability',colitalForm]
]

for meanItem in meansheetList:
    meansheet.write(meanItem[0], meanItem[1], meanItem[2], meanItem[3])

meanmergeList = [[0,28,0,31,'People',bordForm],[0,32,0,35,'People',bordForm],[0,36,0,39,'People',bordForm],[0,40,0,43,'People',bordForm],[0,44,0,47,'People',bordForm],[0,48,0,51,'People',bordForm],[0,52,0,55,'People',bordForm],[0,56,0,59,'People',bordForm],[0,60,0,63,'People',bordForm],[0,64,0,67,'People',bordForm],[1,28,1,31,'Flooding',bordForm],[1,32,1,35,'Flooding',bordForm],[1,36,1,39,'Flooding',bordForm],[1,40,1,43,'Flooding',bordForm],[1,44,1,47,'Recreation',bordForm],[1,48,1,51,'Recreation',bordForm],[1,52,1,55,'Recreation',bordForm],[1,56,1,59,'Temperature regulation',bordForm],[1,60,1,63,'Water supply',bordForm],[1,64,1,67,'Water supply',bordForm]
]

for meanmergeItem in meanmergeList:
    meansheet.merge_range(meanmergeItem[0],meanmergeItem[1],meanmergeItem[2],meanmergeItem[3],meanmergeItem[4],meanmergeItem[5])

#add new worksheet for writing size-relative results and write template cells
sizersheet = workbook.add_worksheet("EfficiencyScores")

sizersheet.set_column(0,2,11.0)
sizersheet.set_column(3,37,11.71)
sizersheetList = [[0,2,'Target System>',italralForm],[1,2,'Attribute Group>',italralForm],[0,4,'Freshwater',bordForm],[0,5,'Freshwater',bordForm],[0,6,'Freshwater',bordForm],[0,7,'Freshwater',bordForm],[0,8,'Freshwater',bordForm],[0,9,'Freshwater',bordForm],[0,10,'Freshwater',bordForm],[0,11,'Freshwater',bordForm],[0,12,'Freshwater',bordForm],[0,13,'Freshwater',bordForm],[0,14,'Terrestrial',bordForm],[0,15,'Terrestrial',bordForm],[0,16,'Terrestrial',bordForm],[0,17,'Terrestrial',bordForm],[0,18,'Terrestrial',bordForm],[0,19,'Terrestrial',bordForm],[0,20,'Terrestrial',bordForm],[0,21,'Terrestrial',bordForm],[0,22,'Terrestrial',bordForm],[0,23,'Marine',bordForm],[0,24,'Marine',bordForm],[0,25,'Marine',bordForm],[0,26,'Marine',bordForm],[0,27,'Marine',bordForm],[0,28,'People',bordForm],[0,29,'People',bordForm],[0,30,'People',bordForm],[0,31,'People',bordForm],[0,32,'People',bordForm],[0,33,'People',bordForm],[0,34,'People',bordForm],[0,35,'People',bordForm],[0,36,'People',bordForm],[0,37,'People',bordForm],[1,3,'Threat',bordForm],[1,4,'Condition',bordForm],[1,5,'Condition',bordForm],[1,6,'Condition',bordForm],[1,7,'Condition',bordForm],[1,8,'Condition',bordForm],[1,9,'Diversity',bordForm],[1,10,'Diversity',bordForm],[1,11,'Habitat',bordForm],[1,12,'Habitat',bordForm],[1,13,'Resilience',bordForm],[1,14,'Climate mitigation',bordForm],[1,15,'Climate mitigation',bordForm],[1,16,'Condition',bordForm],[1,17,'Condition',bordForm],[1,18,'Condition',bordForm],[1,19,'Diversity',bordForm],[1,20,'Diversity',bordForm],[1,21,'Diversity',bordForm],[1,22,'Reslience',bordForm],[1,23,'Condition',bordForm],[1,24,'Condition',bordForm],[1,25,'Condition',bordForm],[1,26,'Habitat',bordForm],[1,27,'Habitat',bordForm],[1,28,'Flooding',bordForm],[1,29,'Flooding',bordForm],[1,30,'Flooding',bordForm],[1,31,'Flooding',bordForm],[1,32,'Recreation',bordForm],[1,33,'Recreation',bordForm],[1,34,'Recreation',bordForm],[1,35,'Temperature regulation',bordForm],[1,36,'Water supply',bordForm],[1,37,'Water supply',bordForm],[2,0,'Project Name',bordForm],[2,1,'Project ID',bordForm],[2,2,'Project area (Ha) *GIS calc.',bordForm],[2,3,'Land use intensification',colForm],[2,4,'Floodplain function',colForm],[2,5,'Freshwater flows',colForm],[2,6,'Freshwater NPS mitigation',colForm],[2,7,'Freshwater NPS prevention',colForm],[2,8,'Riparian function',colForm],[2,9,'Aquatic diversity',colForm],[2,10,'Stream physical variety',colForm],[2,11,'Floodplain habitat',colForm],[2,12,'Wetland habitat',colForm],[2,13,'Stream resilience',colForm],[2,14,'Carbon sequestration',colForm],[2,15,'Carbon storage',colForm],[2,16,'Terrestrial climate flow',colForm],[2,17,'Terrestrial connectivity',colForm],[2,18,'Terrestrial habitat quality',colForm],[2,19,'Terrestrial diversity',colForm],[2,20,'Terrestrial habitat variety',colForm],[2,21,'Terrestrial physical variety',colForm],[2,22,'Terrestrial resilience',colForm],[2,23,'Marine N prevention',colForm],[2,24,'Marine NPS reduction',colForm],[2,25,'Shoreline dynamics',colForm],[2,26,'Future tidal wetlands',colForm],[2,27,'Present tidal wetlands',colForm],[2,28,'Flood damage prevention',colForm],[2,29,'Marine storm surge mitigation',colForm],[2,30,'Riverine flood mitigation',colForm],[2,31,'Riverine flood reduction',colForm],[2,32,'Freshwater aquatic recreation',colForm],[2,33,'Marine aquatic recreation',colForm],[2,34,'Terrestrial recreation',colForm],[2,35,'Heat mitigation',colForm],[2,36,'Ground water supply',colForm],[2,37,'Surface water supply',colForm]
]

for sizerItem in sizersheetList:
    sizersheet.write(sizerItem[0],sizerItem[1],sizerItem[2],sizerItem[3])

#add new worksheet for writing share of state results and write template cells
sharesheet = workbook.add_worksheet("EffectivenessScores")

sharesheet.set_column(0,2,11.0)
sharesheet.set_column(3,37,11.71)
sharesheetList = [[0,2,'Target System>',italralForm],[1,2,'Attribute Group>',italralForm],[0,4,'Freshwater',bordForm],[0,5,'Freshwater',bordForm],[0,6,'Freshwater',bordForm],[0,7,'Freshwater',bordForm],[0,8,'Freshwater',bordForm],[0,9,'Freshwater',bordForm],[0,10,'Freshwater',bordForm],[0,11,'Freshwater',bordForm],[0,12,'Freshwater',bordForm],[0,13,'Freshwater',bordForm],[0,14,'Terrestrial',bordForm],[0,15,'Terrestrial',bordForm],[0,16,'Terrestrial',bordForm],[0,17,'Terrestrial',bordForm],[0,18,'Terrestrial',bordForm],[0,19,'Terrestrial',bordForm],[0,20,'Terrestrial',bordForm],[0,21,'Terrestrial',bordForm],[0,22,'Terrestrial',bordForm],[0,23,'Marine',bordForm],[0,24,'Marine',bordForm],[0,25,'Marine',bordForm],[0,26,'Marine',bordForm],[0,27,'Marine',bordForm],[0,28,'People',bordForm],[0,29,'People',bordForm],[0,30,'People',bordForm],[0,31,'People',bordForm],[0,32,'People',bordForm],[0,33,'People',bordForm],[0,34,'People',bordForm],[0,35,'People',bordForm],[0,36,'People',bordForm],[0,37,'People',bordForm],[1,3,'Threat',bordForm],[1,4,'Condition',bordForm],[1,5,'Condition',bordForm],[1,6,'Condition',bordForm],[1,7,'Condition',bordForm],[1,8,'Condition',bordForm],[1,9,'Diversity',bordForm],[1,10,'Diversity',bordForm],[1,11,'Habitat',bordForm],[1,12,'Habitat',bordForm],[1,13,'Resilience',bordForm],[1,14,'Climate mitigation',bordForm],[1,15,'Climate mitigation',bordForm],[1,16,'Condition',bordForm],[1,17,'Condition',bordForm],[1,18,'Condition',bordForm],[1,19,'Diversity',bordForm],[1,20,'Diversity',bordForm],[1,21,'Diversity',bordForm],[1,22,'Reslience',bordForm],[1,23,'Condition',bordForm],[1,24,'Condition',bordForm],[1,25,'Condition',bordForm],[1,26,'Habitat',bordForm],[1,27,'Habitat',bordForm],[1,28,'Flooding',bordForm],[1,29,'Flooding',bordForm],[1,30,'Flooding',bordForm],[1,31,'Flooding',bordForm],[1,32,'Recreation',bordForm],[1,33,'Recreation',bordForm],[1,34,'Recreation',bordForm],[1,35,'Temperature regulation',bordForm],[1,36,'Water supply',bordForm],[1,37,'Water supply',bordForm],[2,0,'Project Name',bordForm],[2,1,'Project ID',bordForm],[2,2,'Project area (Ha) *GIS calc.',bordForm],[2,3,'Land use intensification',colForm],[2,4,'Floodplain function',colForm],[2,5,'Freshwater flows',colForm],[2,6,'Freshwater NPS mitigation',colForm],[2,7,'Freshwater NPS prevention',colForm],[2,8,'Riparian function',colForm],[2,9,'Aquatic diversity',colForm],[2,10,'Stream physical variety',colForm],[2,11,'Floodplain habitat',colForm],[2,12,'Wetland habitat',colForm],[2,13,'Stream resilience',colForm],[2,14,'Carbon sequestration',colForm],[2,15,'Carbon storage',colForm],[2,16,'Terrestrial climate flow',colForm],[2,17,'Terrestrial connectivity',colForm],[2,18,'Terrestrial habitat quality',colForm],[2,19,'Terrestrial diversity',colForm],[2,20,'Terrestrial habitat variety',colForm],[2,21,'Terrestrial physical variety',colForm],[2,22,'Terrestrial resilience',colForm],[2,23,'Marine N prevention',colForm],[2,24,'Marine NPS reduction',colForm],[2,25,'Shoreline dynamics',colForm],[2,26,'Future tidal wetlands',colForm],[2,27,'Present tidal wetlands',colForm],[2,28,'Flood damage prevention',colForm],[2,29,'Marine storm surge mitigation',colForm],[2,30,'Riverine flood mitigation',colForm],[2,31,'Riverine flood reduction',colForm],[2,32,'Freshwater aquatic recreation',colForm],[2,33,'Marine aquatic recreation',colForm],[2,34,'Terrestrial recreation',colForm],[2,35,'Heat mitigation',colForm],[2,36,'Ground water supply',colForm],[2,37,'Surface water supply',colForm]
]

for shareItem in sharesheetList:
    sharesheet.write(shareItem[0],shareItem[1],shareItem[2],shareItem[3])

n=3
#write project mean sensitivity, efficiency (legacy name = sizer) and effectiveness (legacy name = share) results to appropriate worksheets
for projResult in projResults: #projResults = {projID: {attribute: [mean sens, size-relative, share of state] ...} ...}

    currRow = projDict[projResult][1] #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}
    currName = projDict[projResult][0]
    meansheet.write(currRow,0,currName,regForm)
    meansheet.write(currRow,1,projResult,regForm)
    sizersheet.write(currRow,0,currName,regForm)
    sizersheet.write(currRow,1,projResult,regForm)
    sharesheet.write(currRow,0,currName,regForm)
    sharesheet.write(currRow,1,projResult,regForm)
    currVecHa = projDict[projResult][2] #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}
    meansheet.write(currRow,2,currVecHa,regForm)
    sizersheet.write(currRow,2,currVecHa,regForm)
    sharesheet.write(currRow,2,currVecHa,regForm)

    #retrieve a list of results for the current project
    projAttList = projResults[projResult]

    #iterate through the results for the current project
    for att in projAttList:

        #get the last three characters in att name and check if they are in list indicating people components
        if att[-2:] in ['fn','bp','bv']:
            #if current att is one of people components (fn: function, bp: beneficiary population, bv: beneficiary vulnerability), proceed with writing only mean results
            meanCol = meanColDictionary[att]
            currMean = projAttList[att][0]
            meansheet.write(currRow,meanCol,currMean,regForm)

        else:
            #if current att is not one of the people components, proceed with writing mean, size-relative, and share of state results
            meanCol = meanColDictionary[att]
            currMean = projAttList[att][0]
            meansheet.write(currRow,meanCol,currMean,regForm)

            attCol = colDictionary[att]
            currSizer = projAttList[att][1]
            sizersheet.write(currRow,attCol,currSizer,regForm)

            currShare = projAttList[att][2]
            sharesheet.write(currRow,attCol,currShare,regForm)

    n+=1

#apply conditional formatting to size-relative, and share of state results worksheets
sizersheet.conditional_format(3,3,n,37, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
sizersheet.conditional_format(3,3,n,37, {'type': 'cell', 'criteria': 'equal to', 'value': '"Null"', 'format': greycentForm}) #first row, first column, last row, last column, {conditional format dictionary}
sizersheet.conditional_format(3,3,n,37, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num', 'min_value': 0, 'mid_value': 50, 'max_value': 100, 'min_color': yellow, 'mid_color': green, 'max_color': blue})

sharesheet.conditional_format(3,3,n,37, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
sharesheet.conditional_format(3,3,n,37, {'type': 'cell', 'criteria': 'greater than', 'value': 100, 'format': dblueForm})
sharesheet.conditional_format(3,3,n,37, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num','min_value': 0, 'mid_value': .01, 'max_value': 1, 'min_color': yellow, 'mid_color': green, 'max_color': blue})

#add a new worksheet to the results workbook called scattersheet and write header and legend text
scattersheet = workbook.add_worksheet("ScatterPlot")

scatterList = [[0,0,'Interact with the scatter plots by filtering the table below in any combination of Project ID, Target system, Attribute group, Attribute, and/or Scores',lgboldForm],[13,0,'Scatter plot colors vary by project',italForm],[1,0,'Shapes vary by system/threat',italForm],[2,1,'Freshwater ',regForm],[3,1,'Terrestrial',regForm],[4,1,'Marine',regForm],[5,1,'People',regForm],[6,1,'Threat',regForm]]

for scatterItem in scatterList:
    scattersheet.write(scatterItem[0],scatterItem[1],scatterItem[2],scatterItem[3])

#add legend shape images to scattersheet
scattersheet.insert_image(2,0,circle,{'x_scale':0.7,'y_scale':0.7,'x_offset':20, 'y_offset':4})
scattersheet.insert_image(3,0,square,{'x_scale':0.7,'y_scale':0.7,'x_offset':20, 'y_offset':4})
scattersheet.insert_image(4,0,triangle,{'x_scale':0.7,'y_scale':0.7,'x_offset':20, 'y_offset':4})
scattersheet.insert_image(5,0,diamond,{'x_scale':0.7,'y_scale':0.7,'x_offset':19, 'y_offset':3})
scattersheet.insert_image(6,0,star,{'x_scale':0.7,'y_scale':0.7,'x_offset':20, 'y_offset':4})

#set scattersheet column widths
scattersheet.set_column(0,0,15)
scattersheet.set_column(1,1,15)
scattersheet.set_column(2,2,22)
scattersheet.set_column(3,3,15)
scattersheet.set_column(4,4,22)
scattersheet.set_column(5,5,25)
scattersheet.set_column(6,6,20)
scattersheet.set_column(7,7,20)
scattersheet.set_column(8,8,43)

#define start and end rows for scatter plot table
startRow = 38
if numProj <= 23:
    endRow = startRow+numProj*35
else:
    endRow = startRow+len(topProjs)*35

#create scatter plot table
scattersheet.add_table(startRow,2,endRow,8,{'columns':[{'header':'Project ID'},
                                                    {'header':'Target system','format':regForm},
                                                    {'header':'Attribute group', 'format': regForm},
                                                    {'header':'Attribute', 'format': regForm},
                                                    {'header':'Effectiveness score', 'format': regForm},
                                                    {'header':'Efficiency score', 'format': regForm},
                                                    {'header':'PLOT REFERENCE: Point name', 'format': regForm}],
                                         'style': 'Table Style Medium 15'})

#add freshwater scatter chart to scattersheet
fwscatter = workbook.add_chart({'type': 'scatter'})
scattersheet.insert_chart('C3',fwscatter,{'x_scale':1.27,'y_scale':1.25})

#add terrestrial scatter chart to scattersheet
trscatter = workbook.add_chart({'type': 'scatter'})
scattersheet.insert_chart('C21',trscatter,{'x_scale':1.27,'y_scale':1.25})

#add marine scatter chart to scattersheet
mrscatter = workbook.add_chart({'type': 'scatter'})
scattersheet.insert_chart('G3',mrscatter,{'x_scale':1.27,'y_scale':1.25})

#add people scatter chart to scattersheet
plscatter = workbook.add_chart({'type': 'scatter'})
scattersheet.insert_chart('G21',plscatter,{'x_scale':1.27,'y_scale':1.25})

#add a new worksheet to the results workbook called radarsheet and write the worksheet header line
radarsheet = workbook.add_worksheet("RadarPlots")
radarsheet.write(0,0,"Interact with the radar plots by filtering the table below by Project ID", lgboldForm)

#add radar plot table headers for indicating size-relative and share of state columns
radarsheet.merge_range('B39:AJ39', 'Efficiency scores',lgboldForm)
radarsheet.merge_range('AK39:BS39', 'Effectiveness scores',lgboldForm)

#add radar charts to radarsheet
sizerRad = workbook.add_chart({'type':'radar'})
radarsheet.insert_chart('A3',sizerRad,{'x_scale':2.0,'y_scale':2.5})
sosRad = workbook.add_chart({'type':'radar'})
radarsheet.insert_chart('P3',sosRad,{'x_scale':2.0,'y_scale':2.5})

#define start and end rows for radar plot table
radStart = 39
if numProj <= 23:
    radEnd = radStart+numProj
else:
    radEnd = radStart+len(topProjs)

#create radar plots table with headers
radarsheet.add_table(radStart,0,radEnd,70,{'columns':[{'header':'Project ID'},{'header':'Floodplain function', 'format': regForm},{'header':'Freshwater flows', 'format': regForm},{'header':'Freshwater NPS mitigation', 'format': regForm},
{'header':'Freshwater NPS prevention', 'format': regForm},{'header':'Riparian function', 'format': regForm},{'header':'Aquatic diversity', 'format': regForm},{'header':'Stream physical variety', 'format': regForm},
{'header':'Floodplain habitat', 'format': regForm},{'header':'Wetland habitat', 'format': regForm},{'header':'Stream resilience', 'format': regForm},{'header':'Carbon sequestration', 'format': regForm},{'header':'Carbon storage', 'format': regForm},
{'header':'Terrestrial climate flow', 'format': regForm},{'header':'Terrestrial connectivity', 'format': regForm},{'header':'Terrestrial habitat quality', 'format': regForm},{'header':'Terrestrial diversity', 'format': regForm},
{'header':'Terrestrial habitat variety', 'format': regForm},{'header':'Terrestrial physical variety', 'format': regForm},{'header':'Terrestrial resilience', 'format': regForm},{'header':'Marine N prevention', 'format': regForm},
{'header':'Marine NPS reduction', 'format': regForm},{'header':'Shoreline dynamics', 'format': regForm},{'header':'Future tidal wetlands', 'format': regForm},{'header':'Present tidal wetlands', 'format': regForm},
{'header':'Flood damage prevention', 'format': regForm},{'header':'Marine storm surge mitigation', 'format': regForm},{'header':'Riverine flood mitigation', 'format': regForm},{'header':'Riverine flood reduction','format':regForm},
{'header':'Freshwater aquatic recreation', 'format': regForm},{'header':'Marine aquatic recreation', 'format': regForm},{'header':'Terrestrial recreation', 'format': regForm},
{'header':'Heat mitigation', 'format': regForm},{'header':'Ground water suply', 'format': regForm},{'header':'Surface water supply', 'format': regForm},{'header':'Land use intensification','format':regForm},
{'header':'Floodplain function.', 'format': regForm},{'header':'Freshwater flows.', 'format': regForm},{'header':'Freshwater NPS mitigation.', 'format': regForm},
{'header':'Freshwater NPS prevention.', 'format': regForm},{'header':'Riparian function.', 'format': regForm},{'header':'Aquatic diversity.', 'format': regForm},{'header':'Stream physical variety.', 'format': regForm},
{'header':'Floodplain habitat.', 'format': regForm},{'header':'Wetland habitat.', 'format': regForm},{'header':'Stream resilience.', 'format': regForm},{'header':'Carbon sequestration.', 'format': regForm},{'header':'Carbon storage.', 'format': regForm},
{'header':'Terrestrial climate flow.', 'format': regForm},{'header':'Terrestrial connectivity.', 'format': regForm},{'header':'Terrestrial habitat quality.', 'format': regForm},{'header':'Terrestrial diversity.', 'format': regForm},
{'header':'Terrestrial habitat variety.', 'format': regForm},{'header':'Terrestrial physical variety.', 'format': regForm},{'header':'Terrestrial resilience.', 'format': regForm},{'header':'Marine N prevention.', 'format': regForm},
{'header':'Marine NPS reduction.', 'format': regForm},{'header':'Shoreline dynamics.', 'format': regForm},{'header':'Future tidal wetlands.', 'format': regForm},{'header':'Present tidal wetlands.', 'format': regForm},
{'header':'Flood damage prevention.', 'format': regForm},{'header':'Marine storm surge mitigation.', 'format': regForm},{'header':'Riverine flood mitigation.', 'format': regForm},{'header':'Riverine flood reduction.','format':regForm},
{'header':'Freshwater aquatic recreation.', 'format': regForm},{'header':'Marine aquatic recreation.', 'format': regForm},{'header':'Terrestrial recreation.', 'format': regForm},
{'header':'Heat mitigation.', 'format': regForm},{'header':'Ground water supply.', 'format': regForm},{'header':'Surface water supply.', 'format': regForm},{'header':'Land use intensification.','format':regForm}],'style': 'Table Style Medium 15'})

#define project color formats
col1Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#FF0000'})
col2Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#00FF00'})
col3Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#0000FF'})
col4Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#FF00FF'})
col5Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#00FFFF'})
col6Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#FFCC00'})
col7Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#FF9900'})
col8Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#99CC00'})
col9Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#800000'})
col10Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#3366FF'})
col11Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#9999FF'})
col12Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#800080'})
col13Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#339966'})
col14Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#008000'})
col15Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#000080'})
col16Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#808000'})
col17Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#808080'})
col18Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#008080'})
col19Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#003300'})
col20Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#FF6600'})
col21Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#C0C0C0'})
col22Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#993366'})
col23Form=workbook.add_format({'font_name':'Calibri','font_size':10,'bold':True,'font_color':'#660066'})


#define project color list [hex color, color format]
colorList = (['#FF0000',col1Form],['#00FF00',col2Form],['#0000FF',col3Form],['#FF00FF',col4Form],['#00FFFF',col5Form],['#FFCC00',col6Form],['#FF9900',col7Form],['#99CC00',col8Form],['#800000',col9Form],['#3366FF',col10Form],['#9999FF',col11Form],['#800080',col12Form],['#339966',col13Form],['#008000',col14Form],['#000080',col15Form],['#808000',col16Form],['#808080',col17Form],['#008080',col18Form],['#003300',col19Form],['#FF6600',col20Form],['#C0C0C0',col21Form],['#993366',col22Form],['#660066',col23Form])

#set counter for tracking projects
p = 0

#if the query is a batch of more than 23 projects, proceed with iterating through the top 23 projects in the topProjs dictionary for writing to scatter and radar plot tables, and graphing
if numProj > 23:
    for topProj in topProjs:
        #get id of current top proj and define a project text
        proj = topProjs[p][1]
        projText = "Project " + str(proj)

        #retrieve results for current top proj
        projAtts = projResults[proj]

        #retrieve color for current project
        currColor = colorList[p][0] #colorList = ([hex color, color form], ...)
        currForm = colorList[p][1]

        #write current project name to scattersheet legend
        scattersheet.write(14+p,0,"Project "+str(proj),currForm)

        #write current project name to radarsheet table
        radarsheet.write(40+p,0,proj,regForm)

        #iterate through results for current project
        for att in projAtts:
            #skip the people component results
            if att[-2:] in ['fn','bp','bv']:
                pass
            else:
                #retrieve attribute long name, target system, and attribute group
                attName = attDictionary[att]
                attTarget = classDict[att][0]
                attGroup = classDict[att][1]

                #retrieve attribute share of state and size-relative results
                attSos = projAtts[att][2]
                attSizer = projAtts[att][1]

                #retrieve attribute row number for scatter plot and write results into table
                attRow = p*35+scatDict[att]
                scattersheet.write_row('C'+str(attRow),[proj,attTarget,attGroup,attName,attSos,attSizer,projText+'; '+attName],regForm)

                #retrieve radar table column numbers for size-relative and share of state results for current attribute and write results into table
                sizerCol = radDict[att][0]
                sosCol = radDict[att][1]
                radarsheet.write(40+p,sizerCol,attSizer,regForm)
                radarsheet.write(40+p,sosCol,attSos,regForm)

                #retrieve current symbol for attribute
                currSym = attShapes[att]

                #skip plotting results on the scatter plot that are still under development
                if attSos == 'Under development' or attSizer == 'Under development':
                    pass

                #if the current target system is freshwater, add results to the fwscatter plot
                elif attTarget == 'Freshwater':
                    fwscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current target system is terrestrial, add results to the trscatter plot
                elif attTarget == 'Terrestrial':
                    trscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current target system is marine, add results to the trscatter plot
                elif attTarget == 'Marine':
                    mrscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current target system is people, add results to the trscatter plot
                elif attTarget == 'People':
                    plscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current result is the threat data, add results to all four scatter plots
                elif att == 'landuseintensificationrisk':
                    fwscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})
                    trscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})
                    mrscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})
                    plscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})


        sizerRad.add_series({'name':'=RadarPlots!$A$'+str(41+p),
                             'categories':'=RadarPlots!$B$40:$AJ$40',
                             'values':'=RadarPlots!$B$'+str(p+41)+':AJ$'+str(p+41),
                             'line':{'color':currColor, 'transparency':25}})

        sosRad.add_series({'name':'=RadarPlots!$A$'+str(41+p),
                             'categories':'=RadarPlots!$AK$40:$BS$40',
                             'values':'=RadarPlots!$AK$'+str(p+41)+':BS$'+str(p+41),
                             'line':{'color':currColor, 'transparency':25}})

        p+=1



#if the query is NOT a batch of more than 23 projects, proceed with iterating through all projects in project dictionary, writing to scatter and radar plot tables, and graphing
elif numProj <= 23:
    for proj in projResults:
        projText = "Project " + str(proj)

        #retrieve results for current proj
        projAtts = projResults[proj]

        #retrieve color for current project
        currColor = colorList[p][0] #colorList = ([hex color, color form], ...)
        currForm = colorList[p][1]

        #write current project name to scattersheet legend
        scattersheet.write(14+p,0,"Project "+str(proj),currForm)

        #write current project name to radarsheet table
        radarsheet.write(40+p,0,proj,regForm)

        #iterate through results for current project
        for att in projAtts:
            #skip the people component results
            if att[-2:] in ['fn','bp','bv']:
                pass
            else:
                #retrieve attribute long name, target system, and attribute group
                attName = attDictionary[att]
                attTarget = classDict[att][0]
                attGroup = classDict[att][1]

                #retrieve attribute share of state and size-relative results
                attSos = projAtts[att][2]
                attSizer = projAtts[att][1]

                #retrieve attribute row number for scatter plot and write results into table
                attRow = p*35+scatDict[att]
                scattersheet.write_row('C'+str(attRow),[proj,attTarget,attGroup,attName,attSos,attSizer,projText+'; '+attName],regForm)

                #retrieve radar table column numbers for size-relative and share of state results for current attribute and write results into table
                sizerCol = radDict[att][0]
                sosCol = radDict[att][1]
                radarsheet.write(40+p,sizerCol,attSizer,regForm)
                radarsheet.write(40+p,sosCol,attSos,regForm)

                #retrieve current symbol for attribute
                currSym = attShapes[att]

                #skip plotting results on the scatter plot that are still under development
                if attSos == 'Under development' or attSizer == 'Under development':
                    pass

                #if the current target system is freshwater, add results to the fwscatter plot
                elif attTarget == 'Freshwater':
                    fwscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current target system is terrestrial, add results to the trscatter plot
                elif attTarget == 'Terrestrial':
                    trscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current target system is marine, add results to the mrscatter plot
                elif attTarget == 'Marine':
                    mrscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current target system is people, add results to the plscatter plot
                elif attTarget == 'People':
                    plscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1},
                                            'fill': {'color': currColor, 'transparency': 33}}})

                #if the current result is the threat data, add results to the all scatter plots
                elif att == 'landuseintensificationrisk':
                    fwscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})
                    trscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})
                    mrscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})
                    plscatter.add_series({'name': '=ScatterPlot!$I$'+str(attRow),
                                            'categories': '=ScatterPlot!$G$'+str(attRow),
                                            'values': '=ScatterPlot!$H$'+str(attRow),
                                            'marker': {'type': currSym,
                                                        'size': 6,
                                                        'border': {'color': 'black', 'width':1.5},
                                            'fill': {'color': currColor, 'transparency': 33}}})


        sizerRad.add_series({'name':'=RadarPlots!$A$'+str(41+p),
                             'categories':'=RadarPlots!$B$40:$AJ$40',
                             'values':'=RadarPlots!$B$'+str(p+41)+':AJ$'+str(p+41),
                             'line':{'color':currColor, 'transparency':25}})

        sosRad.add_series({'name':'=RadarPlots!$A$'+str(41+p),
                             'categories':'=RadarPlots!$AK$40:$BS$40',
                             'values':'=RadarPlots!$AK$'+str(p+41)+':BS$'+str(p+41),
                             'line':{'color':currColor, 'transparency':25}})

        p+=1

#get the maximum effectiveness score (legacy name = share of state score) for determining x axis maximum for the scatter plots
shareMax = max(effectScores)

fwscatter.set_x_axis({'min':0.0,'max':shareMax,'name': "project effectiveness score", 'name_font': {'size': 10}, 'num_font': {'size': 9}})
fwscatter.set_y_axis({'name': "project efficiency score", 'name_font': {'size': 10}, 'num_font': {'size': 9}, 'min': 0, 'max': 100, 'major_gridlines': {'visible': False}})
fwscatter.set_plotarea({'border':{'color':'gray','width':1}})
fwscatter.set_legend({'none': True})
fwscatter.set_title({'name':'Target system: Freshwater','name_font':{'size':10}})

trscatter.set_x_axis({'min':0.0,'max':shareMax,'name': "project effectiveness score", 'name_font': {'size': 10}, 'num_font': {'size': 9}})
trscatter.set_y_axis({'name': "project efficiency score", 'name_font': {'size': 10}, 'num_font': {'size': 9}, 'min': 0, 'max': 100, 'major_gridlines': {'visible': False}})
trscatter.set_plotarea({'border':{'color':'gray','width':1}})
trscatter.set_legend({'none': True})
trscatter.set_title({'name':'Target system: Terrestrial','name_font':{'size':10}})

mrscatter.set_x_axis({'min':0.0,'max':shareMax,'name': "project effectiveness score", 'name_font': {'size': 10}, 'num_font': {'size': 9}})
mrscatter.set_y_axis({'name': "project efficiency score", 'name_font': {'size': 10}, 'num_font': {'size': 9}, 'min': 0, 'max': 100, 'major_gridlines': {'visible': False}})
mrscatter.set_plotarea({'border':{'color':'gray','width':1}})
mrscatter.set_legend({'none': True})
mrscatter.set_title({'name':'Target system: Marine','name_font':{'size':10}})

plscatter.set_x_axis({'min':0.0,'max':shareMax,'name': "project effectiveness score", 'name_font': {'size': 10}, 'num_font': {'size': 9}})
plscatter.set_y_axis({'name': "project efficiency score", 'name_font': {'size': 10}, 'num_font': {'size': 10}, 'min': 0, 'max': 100, 'major_gridlines': {'visible': False}})
plscatter.set_plotarea({'border':{'color':'gray','width':1}})
plscatter.set_legend({'none': True})
plscatter.set_title({'name':'Target system: People','name_font':{'size':10}})

sizerRad.set_title({'name': "Efficiency scores", 'name_font': {'size':14}})
sizerRad.set_x_axis({'name': 'Attribute','name_font':{'size':9}, 'num_font':{'size':9}})
sizerRad.set_y_axis({'name': 'Score','name_font':{'size':9}, 'num_font':{'size':9}})
sosRad.set_title({'name': "Effectiveness scores", 'name_font': {'size':14}})
sosRad.set_x_axis({'name': 'Attribute','name_font':{'size':9}, 'num_font':{'size':9}})
sosRad.set_y_axis({'name': 'Score','name_font':{'size':9}, 'num_font':{'size':9}})

#close workbook to save
workbook.close()


#############################################################################################################################################################################################################################
#Write individual reports if desired
if report == "true":
    print "Writing data to reports ..."
    arcpy.AddMessage("Writing data to reports ...")

    #create new workbook with xlswriter
    with xlsxwriter.Workbook(outReport) as resultbook:

        #define workbook colors
        dblue = '#004dce'
        blue = '#0071c6'
        green = '#00b252'
        ygreen = 'effb94'
        yellow = '#ffff9c'
        grey = '#bdbebd'
        dgrey = '#808080'

        #define workbook style formats
        titleForm = resultbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 12, 'align': 'center', 'valign': 'vcenter'})
        boldCentForm = resultbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bottom': True})
        boldWrapForm = resultbook.add_format({'bold':True,'font_name':'Calibri','font_size':9,'text_wrap':True})
        redForm = resultbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 9, 'font_color': 'red', 'align': 'center', 'valign': 'vcenter', 'bottom': True})
        redNolineForm = resultbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 9, 'font_color': 'red', 'align': 'center', 'valign': 'vcenter','text_wrap':True})
        headForm = resultbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'bottom': True, 'text_wrap': True})
        cellForm = resultbook.add_format({'font_name': 'Calibri', 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': True, 'text_wrap': True})
        attForm = resultbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 9, 'font_color': 'white', 'bg_color': dgrey, 'border': True, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
        pattForm = resultbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 9, 'font_color': 'white', 'bg_color': dgrey, 'border': True, 'align': 'left', 'valign': 'vcenter'})
        italralForm = resultbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 9, 'align': 'right', 'valign': 'vcenter'})
        italWrapForm = resultbook.add_format({'italic':True,'font_name':'Calibri','font_size':9,'align':'left','valign':'vcenter','text_wrap':True})
        regForm = resultbook.add_format({'font_name': 'Calibri', 'font_size':9})
        ralForm = resultbook.add_format({'font_name': 'Calibri', 'font_size':9, 'align': 'right'})
        lalForm = resultbook.add_format({'font_name': 'Calibri', 'font_size':9, 'align': 'left'})
        wrapForm = resultbook.add_format({'font_name': 'Calibri', 'font_size':8, 'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
        centForm = resultbook.add_format({'font_name': 'Calibri', 'font_size':9, 'align': 'center'})
        dblueForm = resultbook.add_format({'font_name': 'Calibri','font_size':9,'bg_color': dblue})
        blueForm = resultbook.add_format({'font_name':'Calibri','font_size':9,'bg_color':blue})
        greenForm = resultbook.add_format({'font_name':'Calibri','font_size':9,'bg_color':green})
        ygreenForm = resultbook.add_format({'font_name':'Calibri','font_size':9,'bg_color':ygreen})
        yellForm = resultbook.add_format({'font_name':'Calibri','font_size':9,'bg_color':yellow})
        greyForm = resultbook.add_format({'font_name':'Calibri','font_size':9,'bg_color':grey})
        greycentForm = resultbook.add_format({'font_name':'Calibri','font_size':10,'bg_color':grey,'align':'center'})
        whiteForm = resultbook.add_format({'font_name':'Calibri','font_size':9,'align':'center','bg_color':'white'})
        boldForm = resultbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 9})
        italForm = resultbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 9})
        whiteBordForm = resultbook.add_format({'border':True,'font_name': 'Calibri', 'font_size': 9, 'align':'left','text_wrap':True})
        greyBordForm = resultbook.add_format({'border':True,'font_name': 'Calibri', 'font_size': 9, 'align':'left','text_wrap':True, 'bg_color':grey})

        #add new worksheet for writing reports and set column widths and margins
        resultsheet = resultbook.add_worksheet("SATProjectReport")
        resultsheet.set_column(0,0,11.43)
        resultsheet.set_column(1,1,15.00)
        resultsheet.set_column(2,2,20.00)
        resultsheet.set_column(3,5,16.57)
        resultsheet.set_margins(top=0.3,bottom=0.3,left=0.3,right=0.3)

        #set project and page counters
        p=0
        pg=1
        #iterate through projects
        for proj in projDict: #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}

            #set row heights for current project report
            tallRows = [(p*84) + 0, (p*84) + 4, (p*84) + 8, (p*84) + 37, (p*84) + 80]
            for tall in tallRows:
                resultsheet.set_row(tall,30.6)
            midRows = [(p*84) +9, (p*84) +10, (p*84) +11, (p*84) +12, (p*84) +13, (p*84) +14, (p*84) +15, (p*84) +16, (p*84) +17, (p*84) +18, (p*84) +19, (p*84) +20, (p*84) +21, (p*84) +22, (p*84) +23, (p*84) +24, (p*84) +25, (p*84) +26, (p*84) +27, (p*84) +28, (p*84) +29, (p*84) +30, (p*84) +31, (p*84) +32, (p*84) +38, (p*84) +42, (p*84) +46, (p*84) +50, (p*84) +54, (p*84) +58, (p*84) +62, (p*84) +66, (p*84) +70, (p*84) +74, (p*84) + 81]
            for mid in midRows:
                resultsheet.set_row(mid,23.73)
            shortRows = [(p*84) + 1, (p*84) + 2, (p*84) + 3, (p*84) + 5, (p*84) + 6, (p*84) + 7, (p*84) + 33, (p*84) + 34, (p*84) + 35, (p*84) + 36, (p*84) + 78, (p*84) + 79, (p*84) + 82, (p*84) + 83]
            for short in shortRows:
                resultsheet.set_row(short,14.16)
            vshortRows = [(p*84) + 39, (p*84) + 40, (p*84) + 41, (p*84) + 43, (p*84) + 44, (p*84) + 45, (p*84) + 47, (p*84) + 48, (p*84) + 49, (p*84) + 51, (p*84) + 52, (p*84) + 53, (p*84) + 55, (p*84) + 56, (p*84) + 57, (p*84) + 59, (p*84) + 60, (p*84) + 61, (p*84) + 63, (p*84) + 64, (p*84) + 65, (p*84) + 67, (p*84) + 68, (p*84) + 69, (p*84) + 71, (p*84) + 72, (p*84) + 73, (p*84) + 75, (p*84) + 76, (p*84) + 77]
            for vshort in vshortRows:
                resultsheet.set_row(vshort,12.18)

            #define template cells to write for current project
            template = [[((p*84)+1),0,'Query name:',italralForm],[((p*84)+2),0,'Submitted by:',italralForm],[((p*84)+3),0,'Generated on:',italralForm],[((p*84)+4),0,'Input file:',italralForm],[((p*84)+5),0,'Unique ID field:',italralForm],[((p*84)+1),3,'Project shape ID:',italralForm],[((p*84)+2),3,'Project area*:',italralForm],[((p*84)+3),3,'*GIS calculated',italralForm],[((p*84)+8),0,'Target System',headForm],[((p*84)+8),1,'Attribute Group',headForm],[((p*84)+8),2,'Attribute',headForm],[((p*84)+8),3,'Mean score for project (range 0-20)',headForm],[((p*84)+8),4,'Efficiency score for project (range 0-100)',headForm],[((p*84)+8),5,'Effectiveness score for project (range 0-100+)',headForm],[((p*84)+9),1,'Condition',cellForm],[((p*84)+9),2,'Floodplain function',attForm],[((p*84)+10),1,'Condition',cellForm],[((p*84)+10),2,'Freshwater flows',attForm],[((p*84)+11),1,'Condition',cellForm],[((p*84)+11),2,'Freshwater NPS mitigation',attForm],[((p*84)+12),1,'Condition',cellForm],[((p*84)+12),2,'Freshwater NPS prevention',attForm],[((p*84)+13),1,'Condition',cellForm],[((p*84)+13),2,'Riparian function',attForm],[((p*84)+14),1,'Diversity',cellForm],[((p*84)+14),2,'Aquatic diversity',attForm],[((p*84)+15),1,'Diversity',cellForm],[((p*84)+15),2,'Stream physical variety',attForm],[((p*84)+16),1,'Habitat',cellForm],[((p*84)+16),2,'Floodplain habitat',attForm],[((p*84)+17),1,'Habitat',cellForm],[((p*84)+17),2,'Wetland habitat',attForm],[((p*84)+18),1,'Resilience',cellForm],[((p*84)+18),2,'Stream resilience',attForm],[((p*84)+19),1,'Climate mitigation',cellForm],[((p*84)+19),2,'Carbon sequestration',attForm],[((p*84)+20),1,'Climate mitigation',cellForm],[((p*84)+20),2,'Carbon storage',attForm],[((p*84)+21),1,'Condition',cellForm],[((p*84)+21),2,'Terrestrial climate flow',attForm],[((p*84)+22),1,'Condition',cellForm],[((p*84)+22),2,'Terrestrial connectivity',attForm],[((p*84)+23),1,'Condition',cellForm],[((p*84)+23),2,'Terrestrial habitat quality',attForm],[((p*84)+24),1,'Diversity',cellForm],[((p*84)+24),2,'Terrestrial diversity',attForm],[((p*84)+25),1,'Diversity',cellForm],[((p*84)+25),2,'Terrestrial habitat variety',attForm],[((p*84)+26),1,'Diversity',cellForm],[((p*84)+26),2,'Terrestrial physical variety',attForm],[((p*84)+27),1,'Resilience',cellForm],[((p*84)+27),2,'Terrestrial resilience',attForm],[((p*84)+28),1,'Condition',cellForm],[((p*84)+28),2,'Marine N prevention',attForm],[((p*84)+29),1,'Condition',cellForm],[((p*84)+29),2,'Marine NPS reduction',attForm],[((p*84)+30),1,'Condition',cellForm],[((p*84)+30),2,'Shoreline dynamics',attForm],[((p*84)+31),1,'Habitat',cellForm],[((p*84)+31),2,'Future tidal wetlands',attForm],[((p*84)+32),1,'Habitat',cellForm],[((p*84)+32),2,'Present tidal wetlands',attForm],[((p*84)+37),0,'Target System',headForm],[((p*84)+37),1,'Attribute Group',headForm],[((p*84)+37),2,'Attribute',headForm],[((p*84)+37),3,'Mean score for project (range 0-20)',headForm],[((p*84)+37),4,'Efficiency score for project (range 0-100)',headForm],[((p*84)+37),5,'Effectiveness score for project (range 0-100+)',headForm],[((p*84)+38),2,'Flood damage prevention',attForm],[((p*84)+39),2,'       Function',pattForm],[((p*84)+40),2,'       Beneficiary population',pattForm],[((p*84)+41),2,'       Beneficiary vulnerability',pattForm],[((p*84)+42),2,'Marine storm surge mitigation',attForm],[((p*84)+43),2,'       Function',pattForm],[((p*84)+44),2,'       Beneficiary population',pattForm],[((p*84)+45),2,'       Beneficiary vulnerability',pattForm],[((p*84)+46),2,'Riverine flood mitigation',attForm],[((p*84)+47),2,'       Function',pattForm],[((p*84)+48),2,'       Beneficiary population',pattForm],[((p*84)+49),2,'       Beneficiary vulnerability',pattForm],[((p*84)+50),2,'Riverine flood reduction',attForm],[((p*84)+51),2,'       Function',pattForm],[((p*84)+52),2,'       Beneficiary population',pattForm],[((p*84)+53),2,'       Beneficiary vulnerability',pattForm],[((p*84)+54),2,'Freshwater aquatic recreation',attForm],[((p*84)+55),2,'       Function',pattForm],[((p*84)+56),2,'       Beneficiary population',pattForm],[((p*84)+57),2,'       Beneficiary vulnerability',pattForm],[((p*84)+58),2,'Marine aquatic recreation',attForm],[((p*84)+59),2,'       Function',pattForm],[((p*84)+60),2,'       Beneficiary population',pattForm],[((p*84)+61),2,'       Beneficiary vulnerability',pattForm],[((p*84)+62),2,'Terrestrial recreation',attForm],[((p*84)+63),2,'       Function',pattForm],[((p*84)+64),2,'       Beneficiary population',pattForm],[((p*84)+65),2,'       Beneficiary vulnerability',pattForm],[((p*84)+66),2,'Heat mitigation',attForm],[((p*84)+67),2,'       Function',pattForm],[((p*84)+68),2,'       Beneficiary population',pattForm],[((p*84)+69),2,'       Beneficiary vulnerability',pattForm],[((p*84)+70),2,'Ground water supply',attForm],[((p*84)+71),2,'       Function',pattForm],[((p*84)+72),2,'       Beneficiary population',pattForm],[((p*84)+73),2,'       Beneficiary vulnerability',pattForm],[((p*84)+74),2,'Surface water supply',attForm],[((p*84)+75),2,'       Function',pattForm],[((p*84)+76),2,'       Beneficiary population',pattForm],[((p*84)+77),2,'       Beneficiary vulnerability',pattForm],[((p*84)+80),2,'Threat type',headForm],[((p*84)+80),3,'Mean score for project (range 0-20)',headForm],[((p*84)+80),4,'Efficiency score for project (range 0-100)',headForm],[((p*84)+80),5,'Effectiveness score for project (range 0-100+)',headForm],[((p*84)+81),2,'Land use intensification',attForm]]

            tempMerge = [[(p*84)+0,0,(p*84)+0,5,'SAT PROJECT REPORT',titleForm],[(p*84)+6,0,(p*84)+6,5,'Results are subject to field verification.  Scores are calculated as summaries for project areas - see maps to evaluate sub-project variation.',redForm],[(p*84)+7,0,(p*84)+7,5,'Results for Ecosystem Attributes:',boldCentForm],[(p*84)+9,0,(p*84)+18,0,'Freshwater',cellForm],[(p*84)+19,0,(p*84)+27,0,'Terrestrial ',cellForm],[(p*84)+28,0,(p*84)+32,0,'Marine',cellForm],[(p*84)+36,0,(p*84)+36,5,'Results for People Attributes:',boldCentForm],[(p*84)+38,0,(p*84)+77,0,'People',cellForm],[(p*84)+38,1,(p*84)+41,1,'Flooding ',cellForm],[(p*84)+42,1,(p*84)+45,1,'Flooding ',cellForm],[(p*84)+46,1,(p*84)+49,1,'Flooding ',cellForm],[(p*84)+50,1,(p*84)+53,1,'Flooding ',cellForm],[(p*84)+54,1,(p*84)+57,1,'Recreation',cellForm],[(p*84)+58,1,(p*84)+61,1,'Recreation',cellForm],[(p*84)+62,1,(p*84)+65,1,'Recreation',cellForm],[(p*84)+66,1,(p*84)+69,1,'Temperature regulation',cellForm],[(p*84)+70,1,(p*84)+73,1,'Water supply',cellForm],[(p*84)+74,1,(p*84)+77,1,'Water supply',cellForm],[(p*84)+79,0,(p*84)+79,5,'Results for Threat Analysis:',boldCentForm],[(p*84)+81,0,(p*84)+81,1,'Threat',cellForm]]

            #write template cells for current project
            for item in template:
                resultsheet.write(item[0],item[1],item[2],item[3])
            for mItem in tempMerge:
                resultsheet.merge_range(mItem[0],mItem[1],mItem[2],mItem[3],mItem[4],mItem[5])

            #retrieve size of current project
            currVecHa = projDict[proj][2] #projDict = {projID: [projName, outRow, vectorhectares, rasterhectares, number cells]}
            currVecAc = round(currVecHa * 2.47105,2)

            #define project data cells and write for current project
            data = [[(p*84)+1,1,queName,regForm],[(p*84)+2,1,user,regForm],[(p*84)+3,1,fooTime,regForm],[(p*84)+5,1,projIDField,regForm],[(p*84)+1,4,str(proj),lalForm],[(p*84)+2,4,str(currVecHa) + " Ha (" + str(currVecAc) + " acres)",regForm],[(p*84)+35,0,"Query Name: " + str(queName),regForm],[(p*84)+35,5,"Project Shape ID: " + str(proj),ralForm]]

            dataMerge = [[(p*84)+4,1,(p*84)+4,5,projects,wrapForm],[(p*84)+34,0,(p*84)+34,5,"page " + str(pg),centForm],[(p*84)+83,0,(p*84)+83,5,"page " + str(pg+1),centForm]]

            for datum in data:
                resultsheet.write(datum[0],datum[1],datum[2],datum[3])
            for mDatum in dataMerge:
                resultsheet.merge_range(mDatum[0],mDatum[1],mDatum[2],mDatum[3],mDatum[4],mDatum[5])

            #retrieve senstivity results for current project
            currResults = projResults[proj]
            for att in currResults:
                #if current att is one of the people components, proceed with writing mean results only
                if att[-2:] in ['fn', 'bp', 'bv']:
                    currRow = (p*84)+rowDictionary[att]
                    currMean = currResults[att][0]
                    resultsheet.write(currRow,3,currMean,cellForm)

                #if current att is a people sensitivity result, proceed with writing mean results in single cells, and size-rel and share of shate results as merges
                elif att in ['heatmitigation', 'groundwatersupply', 'surfacewatersupply', 'freshwateraquaticrecreation', 'marineaquaticrecreation', 'terrestrialrecreation', 'marinestormsurgemitigation', 'flooddamageprevention', 'riverinelandscapefloodreduction', 'riverinefpfloodmitigation']:
                    currRow = (p*84)+rowDictionary[att]
                    currMean = currResults[att][0]
                    resultsheet.write(currRow,3,currMean,cellForm)

                    firstRow = (p*84)+mergeDictionary[att][0]
                    lastRow = (p*84)+mergeDictionary[att][1]
                    currSizer = currResults[att][1]
                    resultsheet.merge_range(firstRow,4,lastRow,4,currSizer,cellForm) #first row, first column, last row, last column, data, cell format
                    currSos = currResults[att][2]
                    resultsheet.merge_range(firstRow,5,lastRow,5,currSos,cellForm)

                #if current att is an ecosystem sensitivity result or the landuseintensification result, proceed with writing results directly
                else:
                    currRow = (p*84)+rowDictionary[att]
                    currMean = currResults[att][0]
                    resultsheet.write(currRow,3,currMean,cellForm)
                    currSizer = currResults[att][1]
                    resultsheet.write(currRow,4,currSizer,cellForm)
                    currSos = currResults[att][2]
                    resultsheet.write(currRow,5,currSos,cellForm)


            #apply conditional formatting to size-relative, and share of state results for current project
            #for ecosystem attribute results
            resultsheet.conditional_format((p*84)+9,4,(p*84)+32,4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+9,4,(p*84)+32,4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Null"', 'format': greycentForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+9,4,(p*84)+32,4, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num', 'min_value': 0, 'mid_value': 50, 'max_value': 100, 'min_color': yellow, 'mid_color': green, 'max_color': blue})
            resultsheet.conditional_format((p*84)+9,5,(p*84)+32,5, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+9,5,(p*84)+32,5, {'type': 'cell', 'criteria': 'greater than', 'value': 100, 'format': dblueForm})
            resultsheet.conditional_format((p*84)+9,5,(p*84)+32,5, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num','min_value': 0, 'mid_value': 0.01, 'max_value': 1, 'min_color': yellow, 'mid_color': green, 'max_color': blue})

            #for people attribute results
            resultsheet.conditional_format((p*84)+38,4,(p*84)+77,4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+38,4,(p*84)+77,4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Null"', 'format': greycentForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+38,4,(p*84)+77,4, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num', 'min_value': 0, 'mid_value': 50, 'max_value': 100, 'min_color': yellow, 'mid_color': green, 'max_color': blue})
            resultsheet.conditional_format((p*84)+38,5,(p*84)+77,5, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+38,5,(p*84)+77,5, {'type': 'cell', 'criteria': 'greater than', 'value': 100, 'format': dblueForm})
            resultsheet.conditional_format((p*84)+38,5,(p*84)+77,5, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num','min_value': 0, 'mid_value': 0.01, 'max_value': 1, 'min_color': yellow, 'mid_color': green, 'max_color': blue})

            #for landuseintensification results
            resultsheet.conditional_format((p*84)+81,4,(p*84)+81,4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+81,4,(p*84)+81,4, {'type': 'cell', 'criteria': 'equal to', 'value': '"Null"', 'format': greycentForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+81,4,(p*84)+81,4, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num', 'min_value': 0, 'mid_value': 50, 'max_value': 100, 'min_color': yellow, 'mid_color': green, 'max_color': blue})
            resultsheet.conditional_format((p*84)+81,5,(p*84)+81,5, {'type': 'cell', 'criteria': 'equal to', 'value': '"Under development"', 'format': whiteForm}) #first row, first column, last row, last column, {conditional format dictionary}
            resultsheet.conditional_format((p*84)+81,5,(p*84)+81,5, {'type': 'cell', 'criteria': 'greater than', 'value': 100, 'format': dblueForm})
            resultsheet.conditional_format((p*84)+81,5,(p*84)+81,5, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num','min_value': 0, 'mid_value': 0.01, 'max_value': 10, 'min_color': yellow, 'mid_color': green, 'max_color': blue})

            p+=1
            pg+=2

        #write SAT project report appendix
        #set row heights for report appendix
        appShortRows = [(p*84)+0,(p*84)+1,(p*84)+2,(p*84)+3,(p*84)+4,(p*84)+5,(p*84)+6,(p*84)+7,(p*84)+8,(p*84)+9,(p*84)+10,(p*84)+11,(p*84)+12,(p*84)+13,(p*84)+14,(p*84)+15,(p*84)+16,(p*84)+17,(p*84)+18,(p*84)+19,(p*84)+20,(p*84)+21,(p*84)+22,(p*84)+23,(p*84)+24,(p*84)+25,(p*84)+26,(p*84)+27,(p*84)+28,(p*84)+29,(p*84)+30,(p*84)+31,(p*84)+32,(p*84)+33,(p*84)+34,(p*84)+35,(p*84)+36,(p*84)+37,(p*84)+38,(p*84)+39,(p*84)+40,(p*84)+41,(p*84)+42,(p*84)+43,(p*84)+44,(p*84)+45,(p*84)+46,(p*84)+47,(p*84)+48,(p*84)+49,(p*84)+50,(p*84)+53,(p*84)+54,(p*84)+55,(p*84)+56,(p*84)+57,(p*84)+58,(p*84)+59]

        for appShort in appShortRows:
            resultsheet.set_row(appShort,12)

        appTallRows = [(p*84)+51,(p*84)+52]

        for appTall in appTallRows:
            resultsheet.set_row(appTall,22.5)

        #write appendix cells
        appendList = [[(p*84)+5,1,'Score',italForm],[(p*84)+15,1,'Score',italForm],[(p*84)+6,1,'100',blueForm],[(p*84)+7,1,'50',greenForm],[(p*84)+8,1,'0',yellForm],[(p*84)+9,1,'Null',greyForm],[(p*84)+16,1,'>100',dblueForm],[(p*84)+17,1,'100',blueForm],[(p*84)+18,1,'10',greenForm],[(p*84)+19,1,'1',ygreenForm],[(p*84)+20,1,'0.1',yellForm]]

        for appendL in appendList:
            resultsheet.write(appendL[0],appendL[1],appendL[2],appendL[3])

        appendMerge = [[(p*84)+0,0,(p*84)+0,5,'SAT Project Report References',boldCentForm],[(p*84)+1,0,(p*84)+1,5,'Mean scores reflect average attribute senstivities in project areas (plus components for people attributes).',boldForm],[(p*84)+2,0,(p*84)+2,5,'Range: 0-20 (Null: project is outside scope that currently provides the function and/or service)',italForm],[(p*84)+4,0,(p*84)+4,5,'Efficiency scores reflect the percent rank of the project as compared to other similarly-sized areas across the state.',boldForm],[(p*84)+6,0,(p*84)+9,0,'Efficiency score key',italWrapForm],[(p*84)+5,2,(p*84)+5,4,'Interpretation',italForm],[(p*84)+6,2,(p*84)+6,4,'project score is >= 100% of similarly sized areas across the state',blueForm],[(p*84)+7,2,(p*84)+7,4,'project score is >= 50% of similarly sized areas across the state',greenForm],[(p*84)+8,2,(p*84)+8,4,'project score is >= 0% of similarly sized areas across the state',yellForm],[(p*84)+9,2,(p*84)+9,4,'project is outside scope that currently provides function and/or service',greyForm],[(p*84)+10,0,(p*84)+11,5,'Colors in the report and maps have the same directionality (yellow = low scores, blue = high scores), but the specific score/color classifications may not match exactly due to differences in scale of the calculations.',redNolineForm],[(p*84)+13,0,(p*84)+14,5,'Effectiveness scores reflect how much of the attribute statewide total values are captured by projects, scaled so high value 1,000 acre (~400 Ha) projects have a score of 1.',boldWrapForm],[(p*84)+16,0,(p*84)+20,0,'Effectiveness score key',italWrapForm],[(p*84)+15,2,(p*84)+15,4,'Interpretation',italForm],[(p*84)+16,2,(p*84)+16,4,'only achievable by projects larger than ~40,000 Ha',dblueForm],[(p*84)+17,2,(p*84)+17,4,'highest attainable score for projects of ~40,000 Ha',blueForm],[(p*84)+18,2,(p*84)+18,4,'highest attainable score for projects of ~4,000 Ha',greenForm],[(p*84)+19,2,(p*84)+19,4,'highest attainable score for projects of ~400 Ha',ygreenForm],[(p*84)+20,2,(p*84)+20,4,'highest attainable score for projects of ~40 Ha',yellForm],[(p*84)+21,0,(p*84)+22,5,'Colors in the report and maps have the same directionality (yellow = low scores, blue = high scores), but the specific score/color classifications may not match exactly due to differences in scale of the calculations.',redNolineForm],[(p*84)+24,0,(p*84)+24,1,'Attribute name',boldForm],[(p*84)+24,2,(p*84)+24,5,'Description of function/service/threat mapped',boldForm],[(p*84)+25,0,(p*84)+25,1,'Floodplain function',whiteBordForm],[(p*84)+26,0,(p*84)+26,1,'Freshwater flows',greyBordForm],[(p*84)+27,0,(p*84)+27,1,'Freshwater NPS mitigation',whiteBordForm],[(p*84)+28,0,(p*84)+28,1,'Freshwater NPS prevention',greyBordForm],[(p*84)+29,0,(p*84)+29,1,'Riparian function',whiteBordForm],[(p*84)+30,0,(p*84)+30,1,'Aquatic diversity',greyBordForm],[(p*84)+31,0,(p*84)+31,1,'Stream physical variety',whiteBordForm],[(p*84)+32,0,(p*84)+32,1,'Floodplain habitat',greyBordForm],[(p*84)+33,0,(p*84)+33,1,'Wetland habitat',whiteBordForm],[(p*84)+34,0,(p*84)+34,1,'Stream resilience',greyBordForm],[(p*84)+35,0,(p*84)+35,1,'Carbon sequestration',whiteBordForm],[(p*84)+36,0,(p*84)+36,1,'Carbon storage',greyBordForm],[(p*84)+37,0,(p*84)+37,1,'Terrestrial climate flow',whiteBordForm],[(p*84)+38,0,(p*84)+38,1,'Terrestrial connectivity',greyBordForm],[(p*84)+39,0,(p*84)+39,1,'Terrestrial habitat quality',whiteBordForm],[(p*84)+40,0,(p*84)+40,1,'Terrestrial diversity',greyBordForm],[(p*84)+41,0,(p*84)+41,1,'Terrestrial habitat variety',whiteBordForm],[(p*84)+42,0,(p*84)+42,1,'Terrestrial physical variety',greyBordForm],[(p*84)+43,0,(p*84)+43,1,'Terrestrial resilience',whiteBordForm],[(p*84)+44,0,(p*84)+44,1,'Marine N prevention',greyBordForm],[(p*84)+45,0,(p*84)+45,1,'Marine NPS reduction',whiteBordForm],[(p*84)+46,0,(p*84)+46,1,'Shoreline dynamics',greyBordForm],[(p*84)+47,0,(p*84)+47,1,'Future tidal wetlands',whiteBordForm],[(p*84)+48,0,(p*84)+48,1,'Present tidal wetlands',greyBordForm],[(p*84)+49,0,(p*84)+49,1,'Flood damage prevention',whiteBordForm],[(p*84)+50,0,(p*84)+50,1,'Marine storm surge mitigation',greyBordForm],[(p*84)+51,0,(p*84)+51,1,'Riverine flood mitigation',whiteBordForm],[(p*84)+52,0,(p*84)+52,1,'Riverine flood reduction',greyBordForm],[(p*84)+53,0,(p*84)+53,1,'Freshwater aquatic recreation',whiteBordForm],[(p*84)+54,0,(p*84)+54,1,'Marine aquatic recreation',greyBordForm],[(p*84)+55,0,(p*84)+55,1,'Terrestrial recreation',whiteBordForm],[(p*84)+56,0,(p*84)+56,1,'Heat mitigation',greyBordForm],[(p*84)+57,0,(p*84)+57,1,'Ground water supply',whiteBordForm],[(p*84)+58,0,(p*84)+58,1,'Surface water supply',greyBordForm],[(p*84)+59,0,(p*84)+59,1,'Land use intensification',whiteBordForm],[(p*84)+25,2,(p*84)+25,5,'Supporting current and future functional floodplains.',whiteBordForm],[(p*84)+26,2,(p*84)+26,5,'Supporting ecologically sufficient flows.',greyBordForm],[(p*84)+27,2,(p*84)+27,5,'Mitigating NPS pollution.',whiteBordForm],[(p*84)+28,2,(p*84)+28,5,'Preventing new sources of NPS pollution.',greyBordForm],[(p*84)+29,2,(p*84)+29,5,'Providing material inputs and shading to in-stream habitats.',whiteBordForm],[(p*84)+30,2,(p*84)+30,5,'Supporting aquatic habitat for conservation species.',greyBordForm],[(p*84)+31,2,(p*84)+31,5,'Supporting the full geophysical variety of stream habitats.',whiteBordForm],[(p*84)+32,2,(p*84)+32,5,'Providing floodplain habitat.',greyBordForm],[(p*84)+33,2,(p*84)+33,5,'Providing freshwater wetland habitat.',whiteBordForm],[(p*84)+34,2,(p*84)+34,5,'Supporting high-resilience streams.',greyBordForm],[(p*84)+35,2,(p*84)+35,5,'Carbon sequestration rate, contributing to climate change mitigation.',whiteBordForm],[(p*84)+36,2,(p*84)+36,5,'Storage of carbon, preventing increases in atmospheric carbon.',greyBordForm],[(p*84)+37,2,(p*84)+37,5,'Supporting range shifts and species movement with climate change.',whiteBordForm],[(p*84)+38,2,(p*84)+38,5,'Supporting current-day movement of wildlife.',greyBordForm],[(p*84)+39,2,(p*84)+39,5,'Preventing fragmentation and habitat degredation.',whiteBordForm],[(p*84)+40,2,(p*84)+40,5,'Provision of habitat for conservation species.',greyBordForm],[(p*84)+41,2,(p*84)+41,5,'Lands with high habitat heterogeneity contributing to diversity and resilience.',whiteBordForm],[(p*84)+42,2,(p*84)+42,5,'Providing the full variety of geophysical habitats.',greyBordForm],[(p*84)+43,2,(p*84)+43,5,'Supporting climate change adaptation.',whiteBordForm],[(p*84)+44,2,(p*84)+44,5,'Preventing new nitrogen inputs to the marine system.',greyBordForm],[(p*84)+45,2,(p*84)+45,5,'Lands preventing and mitigating NPS pollution to marine ecosystems.',whiteBordForm],[(p*84)+46,2,(p*84)+46,5,'Supporting shoreline dynamics.',greyBordForm],[(p*84)+47,2,(p*84)+47,5,'Providing and supporting movement of tidal wetlands with climate change.',whiteBordForm],[(p*84)+48,2,(p*84)+48,5,'Providing and supporting current tidal wetlands.',greyBordForm],[(p*84)+49,2,(p*84)+49,5,'Reducing flood damage by excluding infrastructure from flood-prone areas.',whiteBordForm],[(p*84)+50,2,(p*84)+50,5,'Reducing coastal flood damage by capturing/slowing floodwaters and reducing flood extent and depth.',greyBordForm],[(p*84)+51,2,(p*84)+51,5,'Mitigating riverine flood damage by capturing/slowing floodwaters in the floodplain and lower peak flood stages.',whiteBordForm],[(p*84)+52,2,(p*84)+52,5,'Reducing riverine flood damage by capturing stormwater runoff throughout the watershed and reducing storm-related streamflow.',greyBordForm],[(p*84)+53,2,(p*84)+53,5,'Lands contributing to the condition of freshwaterways used for swimming, boating, and fishing.',whiteBordForm],[(p*84)+54,2,(p*84)+54,5,'Lands contributing to the condition of marine waterways used for swimming, boating, and fishing.',greyBordForm],[(p*84)+55,2,(p*84)+55,5,'Providing the potential for hiking and scenery/wildlife viewing.',whiteBordForm],[(p*84)+56,2,(p*84)+56,5,'Supporting temperature regulation through shade production and transpiration/reduced albedo.',greyBordForm],[(p*84)+57,2,(p*84)+57,5,'Lands providing purification and recharge of groundwater for public use.',whiteBordForm],[(p*84)+58,2,(p*84)+58,5,'Lands supporting clean and dependable surface water supplies for public use.',greyBordForm],[(p*84)+59,2,(p*84)+59,5,'Threat of new or intensified anthropogenic land uses (developed and/or agricultural).',whiteBordForm]]

        for appendM in appendMerge:
            resultsheet.merge_range(appendM[0],appendM[1],appendM[2],appendM[3],appendM[4],appendM[5])

        #close resultbook to save
        resultbook.close()

        #try to print excel report as pdf, and then delete excel report
        try:
            excel = client.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(outReport)
            ws = wb.Worksheets[0]
            ws.Visible = 1
            ws.ExportAsFixedFormat(0, outPdf)
            wb.Close()
            os.remove(outReport)

        #if printing excel as pdf fails, print warning message to check for report as excel file instead
        except:
            print "NOTE: Couldn't save report as PDF. Check output directory for Excel file of the report instead."
            arcpy.AddMessage("NOTE: Couldn't save report as PDF. Check output directory for Excel file of the report instead.")

else:
    pass

###########################################################################################################################################################################################
#create maps if desired
if maps == 'true':
    print "Maps are currently in beta production and cannot be created automatically by the query tool.  See the 'SAT Guidance for Using Map Templates' document for help preparing maps using existing templates created for this purpose."
    arcpy.AddMessage("Maps are currently in beta production and cannot be created automatically by the query tool.  See the 'SAT Guidance for Using Map Templates' document for help preparing maps using existing templates created for this purpose.")
else:
    pass

############################################################################################################################################################################################
#create spatial data if desired
if spatial == 'true':
    print "Saving spatial data ..."

    spatialOrder = ("floodplainfunction","freshwaterprovision","freshwaternpsmitigation", "freshwaternpsprevention","riparianfunction", "aquaticdiversity", "streamphysicalvariety", "floodplainhabitat", "wetlandhabitat",
    "streamresilience","carbonsequestration", "carbonstorage","terrestrialclimateflow", "terrestrialconnectivity", "terrestrialhabitatquality", "terrestrialdiversity", "terrestrialhabitatvariety", "terrestrialphysicalvariety",
    "terrestrialresilience","marinenprevention", "marinenpsreduction","shorelinedynamics","futuretidalwetlands","presenttidalwetlands","flooddamageprevention", "marinestormsurgemitigation", "riverinefpfloodmitigation",
    "riverinelandscapefloodreduction", "freshwateraquaticrecreation","marineaquaticrecreation","terrestrialrecreation", "heatmitigation", "groundwatersupply", "surfacewatersupply", "landuseintensificationrisk")

    #Create a new geodatabase in the output directory and copy the projWork feature class into this gdb
    outSpatial = outDir + "\\" + "SATResults_" + str(queName) + "_" + str(runStamp) + ".gdb\\" + queName
    arcpy.CreateFileGDB_management(outDir, "SATResults_" + str(queName) + "_" + str(runStamp))
    arcpy.CopyFeatures_management(projWork, outSpatial)

    #add fields to the outSpatial dataset for storing efficiency scores
    efyUpdate = list()
    for attribute in spatialOrder:
        efyField = fieldReverse[attribute] + "_efy"

        if attribute in devList:
            arcpy.AddField_management(outSpatial, efyField, "TEXT")
        else:
            arcpy.AddField_management(outSpatial, efyField, "DOUBLE")
        efyUpdate.append(efyField)

    efyUpdate.insert(0,projIDField)
    #update efficiency fields in the outSpatial dataset
    with arcpy.da.UpdateCursor(outSpatial,efyUpdate) as cursor:
        for row in cursor:
            currId = row[0]
            idResults = projResults[currId]

            efy = 1
            #iterate through results for currId
            for attribute in spatialOrder:
                attEfy = idResults[attribute][1]
                if attEfy == 'Null':
                    pass
                else:
                    row[efy] = attEfy
                efy+=1
            cursor.updateRow(row)

    #add fields to the outSpatial dataset for storing effectiveness scores
    eftUpdate = list()
    for attribute in spatialOrder:
        eftField = fieldReverse[attribute] + "_eft"
        if attribute in devList:
            arcpy.AddField_management(outSpatial, eftField, "TEXT")
        else:
            arcpy.AddField_management(outSpatial, eftField, "DOUBLE")
        eftUpdate.append(eftField)

    eftUpdate.insert(0,projIDField)
    #update effectiveness fields in the outSpatial dataset
    with arcpy.da.UpdateCursor(outSpatial,eftUpdate) as cursor:
        for row in cursor:
            currId = row[0]
            idResults = projResults[currId]

            eft = 1
            #iterate through results for currId
            for attribute in spatialOrder:
                attEft = idResults[attribute][2]
                if attEfy == 'Null':
                    pass
                else:
                    row[eft] = attEft
                eft+=1
            cursor.updateRow(row)

    #add fields to the outSpatial dataset for storing efficiency*effectiveness scores
    exeUpdate = list()
    for attribute in spatialOrder:
        exeField = fieldReverse[attribute] + "_exe"
        if attribute in devList:
            arcpy.AddField_management(outSpatial, exeField, "TEXT")
        else:
            arcpy.AddField_management(outSpatial, exeField, "DOUBLE")
        exeUpdate.append(exeField)

    exeUpdate.insert(0,projIDField)
    #update efficiency*effectiveness fields in the outSpatial dataset
    with arcpy.da.UpdateCursor(outSpatial,exeUpdate) as cursor:
        for row in cursor:
            currId = row[0]
            idResults = projResults[currId]

            exe = 1
            #iterate through results for currId
            for attribute in spatialOrder:
                attExe = idResults[attribute][3]
                if attExe == 'Null':
                    pass
                else:
                    row[exe] = attExe
                exe+=1
            cursor.updateRow(row)
else:
    pass

###################################################################################################################################################################################################################################
#Clean up in_memory workspace and intermediate files
arcpy.Delete_management("in_memory")

for item in toDelete:
    try:
        arcpy.Delete_management(item)
    except:
        print "Couldn't delete intermediate dataset in default geodatabase " + scratch + "; please clean up manually!"
        arcpy.AddMessage("Couldn't delete intermediate dataset in default geodatabase " + scratch + "; please clean up manually!")

#create copy of project shapes and results for the query tool archive
arcpy.CopyFeatures_management(projects,shapeArchive+"\\SATQuery_"+str(queName)+"_"+str(runStamp))
shutil.copy(outDir + "\\SATResults_" + str(queName) + "_" + str(runStamp) + ".xlsx", resultsArchive+"\\SATResults_"+str(queName)+"_"+str(runStamp)+".xlsx")

#print status and caution messages, if applicable
print "All done!  Check results in " + outDir
arcpy.AddMessage("All done! Check results in " + outDir)

if len(bigList) > 0:
    print "   REMINDER: Project(s) " + str(bigList) + " are larger than the MAXIMUM size threshold (24,600 Ha) and were not analyzed ..."
    arcpy.AddMessage("   REMINDER: Project(s) " + str(bigList) + " are larger than the MAXIMUM size threshold (24,600 Ha) and were not analyzed ...")

if len(smallList) > 0:
    print "   REMINDER: Project(s) " + str(smallList) + " are smaller than the MINIMUM size threshold (1 Ha) and were not analyzed ..."
    arcpy.AddMessage("   REMINDER: Project(s) " + str(smallList) + " are smaller than the MINIMUM size threshold (1 Ha) and were not analyzed ...")

elapsed = round((time.time() - start)/60.0,2)
print "took " + str(elapsed) + " minutes for " + str(len(projDict)) + " projects"


