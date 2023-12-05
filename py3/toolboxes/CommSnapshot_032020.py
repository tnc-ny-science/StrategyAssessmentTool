#-------------------------------------------------------------------------------
# Name:        SAT Community Snapshot
# Purpose:
#
# Author:      shannon.thol
#
# Created:     03/20/2020
# Copyright:   (c) shannon.thol 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import arcpy
print "Loading packages ..."
arcpy.AddMessage("Loading packages ...")

import shutil, re, sys #shutil used for making copy of spreadsheet to save in archives,
##sys.path.append(r"D:\gisdata\Resources\conda\arcgispro-py3\Lib\site-packages")
#sys.path.append(r"E:\Python27\ArcGIS10.7\Lib\site-packages")
#sys.path.append(r"E:\Python27\ArcGIS10.7\Lib\site-packages\win32")
#sys.path.append(r"E:\Python27\ArcGIS10.7\Lib\site-packages\win32\lib")
##from win32com import client

import xlsxwriter, os, csv, time, datetime, xlrd
import numpy as np
from arcpy import env
from arcpy.sa import *
arcpy.env.overwriteOutput = True
arcpy.CheckOutExtension("Spatial")

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

#Get path to ALICE data
alice = r"D:\gisdata\Projects\Regional\StrategyAssessmentTool\RESTRICTED_Data4ComSnapshot\CommSnapshotQueryData.gdb\SAT_vuln_vec_ALICE"

#get path to enhanced EPA EJSCREEN data
ejscreen = r'D:\gisdata\Projects\Regional\StrategyAssessmentTool\RESTRICTED_Data4ComSnapshot\CommSnapshotQueryData.gdb\SAT_vuln_vec_EJSCREEN'

#paths to shape and results archives for saving a copy of the spatial input and results spreadsheet
shapeArchive = 'D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_ComSnapshotArchive\\ShapeArchive.gdb'
resultsArchive = 'D:\\gisdata\\Projects\\Regional\\StrategyAssessmentTool\\RESTRICTED_ComSnapshotArchive\\ResultsArchive'

############################################################################################################################################################################################################################################################
#set path to default SAT workspace
##scratch = 'D:\\Users\\'+ os.getenv('username') + '\\Documents\\ArcGIS\\SATworkspace.gdb'
scratch = 'E:\\Users\\'+ os.getenv('username') + '\\Documents\\ArcGIS\\SATworkspace.gdb'

#if the SATworkspace gdb already exists, set is as the workspace
if arcpy.Exists(scratch):
    arcpy.env.workspace = scratch

#if the SATworkspace gdb doesn't already exist, create it and set it as the workspace
else:
    arcpy.CreateFileGDB_management('D:\\Users\\'+ os.getenv('username') + '\\Documents\\ArcGIS', 'SATworkspace.gdb')
    arcpy.env.workspace = scratch

#create empty list for storing paths of temp data that need to be deleted at end of run
toDelete = list()

##############################################################################################################################################################################################################################################################
#Get path for proposed project shapes and name of unique ID field from user input
projects = arcpy.GetParameterAsText(0)
##projects = r'D:\GIS_Data\Personal\sthol\CommunitySnapshotDev\data\CommunitySnapshotTESTS.gdb\fourParcels_Albers'

#Get query name from user input
queEntName = arcpy.GetParameterAsText(1)
##queEntName = 'fourParcelsTest'
queName = re.sub('\W+','',queEntName.lower())

#Get name of unique numeric ID field (MUST BE INTEGER FIELD) from user input
projIDField = arcpy.GetParameterAsText(2)
##projIDField = 'SAT_ID'

#Get optional name of project "name" field from user input
projNameField = arcpy.GetParameterAsText(3)
##projNameField = 'MAILING'

#Get path for directory to write individual reports from user input
outDir = arcpy.GetParameterAsText(4)
##outDir = r'D:\GIS_Data\Personal\sthol\CommunitySnapshotDev\tests'

############################################################################################################################################################################################################################################################
#Derive path for output excel spreadsheet based on user supplied output directory, query name, and name/time run stamp
outPath = outDir + "\\SATCommunitySnapshot_" + str(queName) + "_" + str(runStamp) + ".xlsx"

############################################################################################################################################################################################################################################################
#define dictionaries of field names
ejFields = {u'ID': 'Census tract ID', u'ACSTOTPOP': 'Estimated population', u'MINORPCT': '% minority', u'LESSHSPCT': '% < high school education', u'LINGISOPCT': '% linguistic isolation',
u'UNDER5PCT': '% under age 5', u'OVER64PCT': '% over age 64', u'DISABLPCT': '% disabled'}

alFields = {u'GEOID_Mod': 'Census unit ID', u'Households_Est': 'Estimated number of households', u'ALPOV_ICEPCT': '% poverty + ALICE'}

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


#Dissolve the input shapes on the projIDField field, and reproject the input shapes to use NAD_1983_Contiguous_USA_Albers coordinate system (standard for SAT)
projDiss = arcpy.Dissolve_management(projects, "in_memory\\projectsdiss", projIDField, [[projNameField, 'FIRST']])

projWork = arcpy.Project_management(projDiss, scratch + "\\" + str(queName) + "_" + str(runStamp) + "_allprojects", coorSystem)
arcpy.Delete_management(projDiss)
toDelete.append(projWork)

#If a field named "SAT_Hectares" already exists in the project attribute table, overwrite the values with a new geometry calculation
if 'SAT_Hectares' in [n.name for n in arcpy.ListFields(projWork)]:
    arcpy.CalculateField_management(projWork, "SAT_Hectares", "!SHAPE.AREA@HECTARES!", "PYTHON")
#If a field named "SAT_Hectares" doesn't already exist in the project attribute table, proceed with adding it and calculating area
else:
    arcpy.AddField_management(projWork, "SAT_Hectares", "DOUBLE")
    arcpy.CalculateField_management(projWork, "SAT_Hectares", "!SHAPE.AREA@HECTARES!", "PYTHON")

#Create empty dictionaries for storing results
idsDict= dict() #idsDict = {'OBJECITD': projID}
projDict = dict() #projDict = {projID: [projName, vector hectares, raster hectares, number cells] ...}
ejResults = dict() #ejResults = {projID: {unitID: [est pop, % minority, % < high school: 4, % ling iso, % under 5, % over 64, % disabled] ...} ...}
alResults = dict() #alResults = {projID: {unitID: [est pop, % ALICE] ...} ...}

#Create lists for storing size flagged projects (don't meet the minimum or maximum size thresholds)
smallList = list()
bigList = list()

#process input polygon data
print "Checking project polygons ..."
arcpy.AddMessage("Checking project polygons ...")

#Iterate through the projects in the dissolved project feature class and retrieve the IDs and vector area in hectares
with arcpy.da.UpdateCursor(projWork,[projIDField,'FIRST_' + str(projNameField),"SAT_Hectares", "OBJECTID"]) as cursor:
    for row in cursor:
        #check project size and add items to small or big list as appropriate
        if row[2] < 1:
            print "   FAILED CHECKS: Project " + str(row[0]) + " is smaller than the MINIMUM size threshold (1 Ha) and will not be analyzed ..."
            arcpy.AddMessage("   FAILED CHECKS: Project " + str(row[0]) + " is smaller than the MINIMUM size threshold (1 Ha) and will not be analyzed ...")
            cursor.deleteRow()
            smallList.append([row[0], row[1]])
##        elif row[2] > 24600:
##            print "   FAILED CHECK: Project " + str(row[0]) + " is larger than the MAXIMUM size threshold (24,600 Ha) and will not be analyzed ..."
##            arcpy.AddMessage("   FAILED CHECK: Project " + str(row[0]) + " is larger than the MAXIMUM size threshold (24,600 Ha) and will not be analyzed ...")
##            cursor.deleteRow()
##            bigList.append([row[0], row[1]])
        else:
            #populate projDict information for current project polygon
            projDict[row[0]] = ['']*2 #projDict = {projID: [projName, vectorhectares] ...}
            projDict[row[0]][0] = str(row[1]) #postion 0 in projDict = projName
            projDict[row[0]][1] = round(float(row[2]),2) #position 1 in projDict = vectorhectares

            #populate idsDict information for current project polygon
            idsDict[row[3]] = row[0]

            #prepare results dictionaries for current project polygon (to populate later)
            ejResults[row[0]] = dict() #ejResults = {projID: {unitID: [est pop, % minority, % < high school: 4, % ling iso, % under 5, % over 64, % disabled] ...} ...}
            alResults[row[0]] = dict() #alResults = {projID: {unitID: [est pop, % ALICE] ...} ...}

#check number of features remaining in projWork, if no features remain after removing too big and too small projects, terminate analysis
if int(arcpy.GetCount_management(projWork)[0]) == 0:
    print "   FAILED CHECKS: No projects met the size requirements, terminating analysis ..."
    arcpy.AddMessage("   FAILED CHECKS: No projects met the size requirements, terminating analysis ...")
    arcpy.Delete_management(projWork)
    sys.exit(0)

#############################################################################################################################################################################################################################################################
#Create dictionaries of ALICE and EJSCREEN data
print "Importing reference data ..."
arcpy.AddMessage("Importing reference data ...")

aliceData = dict() #aliceData = {'OBJECTID': ['GEOID10_Mod','Households_Est', 'ALICEPOV_PCT'] ... }
ejscreenData = dict() #ejscreenData = {'OBJECTID': ['ID', 'ACSTOTPOP', 'MINORPCT', 'LESSHSPCT', 'LINGISOPCT', 'UNDER5PCT', 'OVER64PCT', 'DISABLPCT'] ... }

with arcpy.da.SearchCursor(alice, ['OBJECTID', 'GEOID_Mod', 'Households_Est', 'ALICEPOV_PCT']) as cursor:
    for row in cursor:
        aliceData[row[0]] = [row[1], row[2], row[3]]

with arcpy.da.SearchCursor(ejscreen, ['OBJECTID', 'ID', 'ACSTOTPOP', 'MINORPCT', 'LESSHSPCT', 'LINGISOPCT', 'UNDER5PCT', 'OVER64PCT', 'DISABLPCT']) as cursor:
    for row in cursor:
        ejscreenData[row[0]] = [row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8]]

#############################################################################################################################################################################################################################################################
#Generate near tables between projects and ALICE and EJSCREEN data data using 1/2 mile as the search radius
print "Identifying census units near projects ..."
arcpy.AddMessage("Finding nearby census units ...")

nearAlice = arcpy.GenerateNearTable_analysis(projWork, alice, scratch + "\\" + str(queName) + "_" + str(runStamp) + "_allprojects_near_alice", "0.5 Miles", "NO_LOCATION", "NO_ANGLE", "ALL")
toDelete.append(nearAlice)

nearEjscreen = arcpy.GenerateNearTable_analysis(projWork, ejscreen, scratch + "\\" + str(queName) + "_" + str(runStamp) + "_allprojects_near_ejscreen", "0.5 Miles", "NO_LOCATION", "NO_ANGLE", "ALL")
toDelete.append(nearEjscreen)

#Iterate through records in near table results and populate ej and alice results dictionaries

#idsDict = {'OBJECITD': projID}
#projDict = {projID: [projName, vector hectares] ...}

#ejResults = {projID: {unitID: [est pop, % minority, % < high school: 4, % ling iso, % under 5, % over 64, % disabled] ...} ...}
#alResults = {projID: {unitID: [est pop, % ALICE] ...} ...}

#aliceData = {'OBJECTID': ['GEOID_Mod', 'Households_Est', 'ALICEPOV_PCT'] ... }
#ejscreenData = {'OBJECTID': ['ID', 'ACSTOTPOP', 'MINORPCT', 'LESSHSPCT', 'LINGISOPCT', 'UNDER5PCT', 'OVER64PCT', 'DISABLPCT'] ... }

print "Reading results ..."
arcpy.AddMessage("Reading results ...")
with arcpy.da.SearchCursor(scratch + "\\" + str(queName) + "_" + str(runStamp) + "_allprojects_near_alice", ['IN_FID', 'NEAR_FID']) as cursor:
    for row in cursor:
        projOid = row[0]
        projId = idsDict[projOid]
        aliceOid = row[1]
        currAlice = aliceData[aliceOid]
        aliceId = currAlice[0]
        estPop = currAlice[1]
        alicePct = int(round(currAlice[2]*100.0,0))

        #if the current projId is already in the alResults dictionary retrieve current results for the projId before adding new data
        if projId in alResults:
            currResults = alResults[projId]
            currResults[aliceId] = [estPop, alicePct]
            alResults[projId] = currResults

        #if the current projId is not already in the alResults dictionary, add the new data directly
        else:
            alResults[projId] = {aliceId: [estPop, alicePct]}

#idsDict = {'OBJECITD': projID}
#projDict = {projID: [projName, vector hectares] ...}

#ejResults = {projID: {unitID: [est pop, % minority, % < high school: 4, % ling iso, % under 5, % over 64, % disabled] ...} ...}
#alResults = {projID: {unitID: [est pop, % ALICE] ...} ...}

#aliceData = {'OBJECTID': ['GEOID_Mod', 'Households_Est', 'ALICEPOVPCT'] ... }
#ejscreenData = {'OBJECTID': ['ID', 'ACSTOTPOP', 'MINORPCT', 'LESSHSPCT', 'LINGISOPCT', 'UNDER5PCT', 'OVER64PCT', 'DISABLPCT'] ... }

with arcpy.da.SearchCursor(scratch + "\\" + str(queName) + "_" + str(runStamp) + "_allprojects_near_ejscreen", ['IN_FID', 'NEAR_FID']) as cursor:
    for row in cursor:
        projOid = row[0]
        projId = idsDict[projOid]
        ejscreenOid = row[1]
        currEjscreen = ejscreenData[ejscreenOid]
        ejscreenId = currEjscreen[0]
        estPop = currEjscreen[1]
        minPct = int(round(currEjscreen[2]*100.0,0))
        lesshsPct = int(round(currEjscreen[3]*100.0,0))
        lingisoPct = int(round(currEjscreen[4]*100.0,0))
        under5Pct = int(round(currEjscreen[5]*100.0,0))
        over64Pct = int(round(currEjscreen[6]*100.0,0))
        disablPct = int(round(currEjscreen[7]*100.0,0))

        #if the current projId is already in the ejResults dictionary retrieve current results for the projId before adding new data
        if projId in ejResults:
            currResults = ejResults[projId]
            currResults[ejscreenId] = [estPop, minPct, lesshsPct, lingisoPct, under5Pct, over64Pct, disablPct]
            ejResults[projId] = currResults

        #if the current projId is not already in the aejResults dictionary, add the new data directly
        else:
            alResults[projId] = {ejscreenId: [estPop, minPct, lesshsPct, lingisoPct, under5Pct, over64Pct, disablPct]}

#############################################################################################################################################################################################################################################################
#Write results to excel spreadsheet
print "Writing results to Excel spreadsheet ..."
arcpy.AddMessage("Writing results to Excel spreadsheet ...")

#create new workbook with xlswriter
workbook = xlsxwriter.Workbook(outPath)

#define workbook colors
lgtOrange = '#fff3e6'
medOrange = '#ffb566'
drkOrange= '#e67700'
grey = '#bdbebd'
dgrey = '#808080'

#define workbook style formats
boldForm = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 10})
italForm = workbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 10})
ralitalForm = workbook.add_format({'italic': True, 'font_name': 'Calibri', 'font_size': 10, 'align': 'right'})
colForm = workbook.add_format({'font_name': 'Calibri', 'font_size': 10, 'bg_color': dgrey, 'font_color': 'white', 'bold': True, 'align': 'center', 'border': True, 'text_wrap': True})
regForm = workbook.add_format({'font_name': 'Calibri', 'font_size': 10})
redForm = workbook.add_format({'bold': True, 'font_name': 'Calibri', 'font_size': 10, 'font_color': 'red'})
numForm = workbook.add_format({'font_name': 'Calibri', 'font_size': 10, 'num_format': '#,##0'})
bordForm = workbook.add_format({'border': True, 'font_name': 'Calibri', 'font_size': 10, 'text_wrap': True, 'align': 'center'})

#add new worksheet for writing information about the query run and write template cells
querysheet = workbook.add_worksheet("QueryInformation")
querysheet.set_column(0,0,14.0)
querysheetList = [[0,0,'SAT Community Snapshot Query Results',boldForm],[1,0,'NOTE: All results are estimates derived from the most recent US Census Bureau data available for each value.  ALICE and EJSCREEN values are reported for different spatial units.  Please consult the spatial data and exercise good judgement when interpreting these results.',boldForm],[3,0,'Query name:',italForm],[4,0,'Submitted by:',italForm],[5,0,'Generated on:',italForm],[6,0,'Input file:',italForm],[7,0,'Unique ID field:',italForm],[8,0,'Name field:',italForm]]

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
refsheetList = [[0,0,u'EJSCREEN DATA:',boldForm],
[1,1,u'Sources:',ralitalForm],
[1,2,u'Environmental Protection Agency Environmental Justice Screening and Mapping Tool (https://www.epa.gov/ejscreen/what-ejscreen) and American Community Survey (https://www.census.gov/programs-surveys/acs)',regForm],
[2,1,u'Spatial Units:',ralitalForm],
[2,2,u'Census Tracts',regForm],
[3,1,u'Vintage:',ralitalForm],
[3,2,u'Data is from 2015-2011 American Community Survey 5-Year estimates',regForm],
[4,1,u'Notes:',ralitalForm],
[4,2,u'All data is from the EPA EJSCREEN index (source 2011-2015 American Community Survey 5-Year estimates), except % Disability, which was obtained separately from the American Community Survey and added based on TNC-NY staff input.',regForm],
[6,0,u'ALICE DATA:',boldForm],
[7,1,u'Sources:',ralitalForm],
[7,2,u'United Way Asset Limited, Income Constrained, Employed (ALICE) Project (https://www.unitedforalice.org/new-york)',regForm],
[8,1,u'Spatial Units:',ralitalForm],
[8,2,u'Aggregation of Census Places, Public Use Microdata Areas (PUMA), and County Subdivisions',regForm],
[9,1,u'Vintage:',ralitalForm],
[9,2,u'Data is from 2012-2016 American Community Survey 5-Year estimates',regForm],
[10,1,u'Notes:',ralitalForm],
[10,2,u'ALICE and poverty data used in the SAT Community Snapshot is of a more recent vintage and mapped to smaller spatial units in NYC than those used in analyses to generate SAT ecosystem service sensitivity grids.',regForm],
[12,0,u'METHOD AND USE NOTES:',boldForm],
[13,1,u'All census units that are within 1/2 mile (straightline distance from edge of project polygon to edge of census unit polygon) of a proposed project are included in the Community Snapshot results.',regForm],
[14,1,u'Percentages and population/household numbers should be treated as estimates from a single point in time (i.e. we did not attempt to capture trends over time).',regForm],
[15,1,u'Percentages and population/household estimates are reported for the entire census unit - please consult maps and spatial data and exercise good judgement when interpreting the results.',regForm],
[16,1,u'These statistics are intended to help inform stakeholder and community engagement efforts, not replace them.',regForm],
[18,0,u'COLOR RAMP KEY:',boldForm],
[19,1,u'Percentage of population/households',regForm],
[20,1,100,regForm],
[21,1,75,regForm],
[22,1,50,regForm],
[23,1,25,regForm],
[24,1,0,regForm]]

refsheet.conditional_format(20,1,24,1, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num', 'min_value':0, 'mid_value':50, 'max_value':100, 'min_color': lgtOrange, 'mid_color': medOrange, 'max_color': drkOrange})

for refItem in refsheetList:
    refsheet.write(refItem[0],refItem[1],refItem[2],refItem[3])

#add new worksheet for writing land cover percent results and write template cells
alicesheet = workbook.add_worksheet("ALICE")
alicesheet.set_column(0,1,11.0)
alicesheet.set_column(2,6,13.57)
alicesheetList = [[0,0,'Project name',bordForm],[0,1,'Project ID',bordForm],[0,2,'Project area (Ha) *GIS calc.',bordForm],[0,3,'Project area (acres) * GIS calc.',bordForm],[0,4,'Census unit ID',colForm],[0,5,'Census unit estimated number of households',colForm],[0,6,'Census unit % households below poverty or ALICE thresholds',colForm]]

for aliceItem in alicesheetList:
    alicesheet.write(aliceItem[0], aliceItem[1], aliceItem[2], aliceItem[3])

#write ALICE results to ALICE sheet
#alResults = {projID: {unitID: [est households, % ALICE] ...} ...}
#Create counter for keeping track of row numbers
currRow = 1
for currProj in sorted(alResults):
    #retrieve name and area of current project
    currName = projDict[currProj][0] #projDict = {projID: [projName, vectorhectares]}
    currHectares = projDict[currProj][1]
    currAcres = round(projDict[currProj][1]*2.47105,2)

    #retrieve dictionary of results for current project
    currResults = alResults[currProj]

    #iterate through census subunits in current dictionary of results and write to appropriate line of ALICE shreadsheet
    for aliceId in sorted(currResults):
        #retrieve current ALICE results
        currAlice = currResults[aliceId]

        alicesheet.write(currRow,0,currName,regForm) #projName
        alicesheet.write(currRow,1,currProj,regForm) #projID
        alicesheet.write(currRow,2,currHectares,regForm) #vectorhectares
        alicesheet.write(currRow,3,currAcres,regForm) #vectoracres
        alicesheet.write(currRow,4,aliceId,regForm) #censusId
        alicesheet.write(currRow,5,currAlice[0],numForm) #est households
        alicesheet.write(currRow,6,currAlice[1],regForm) #% ALICE

        #advance row counter by 1
        currRow+=1

#apply conditional formatting to alice results worksheets  #(first_row, first_col, last_row, last_col)
alicesheet.conditional_format(1,6,currRow-1,6, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num', 'min_value':0, 'mid_value':50, 'max_value':100, 'min_color': lgtOrange, 'mid_color': medOrange, 'max_color': drkOrange})

#add new worksheet for writing land cover percent results and write template cells
ejscreensheet = workbook.add_worksheet("EJSCREEN")
ejscreensheet.set_column(0,1,11.0)
ejscreensheet.set_column(2,11,13.57)
ejscreensheetList = [[0,0,'Project name',bordForm],[0,1,'Project ID',bordForm],[0,2,'Project area (Ha) *GIS calc.',bordForm],[0,3,'Project area (acres) *GIS calc.',bordForm],[0,4,'Census unit ID',colForm],[0,5,'Census unit estimated population',colForm],[0,6,'Census unit % population minority',colForm], [0,7,'Census unit % population < high school education',colForm],[0,8,'Census unit % population linguistic isolation',colForm],[0,9,'Census unit % population under age 5',colForm],[0,10,'Census unit % population over age 64',colForm],[0,11,'Census unit % population disabled',colForm]]

for ejscreenItem in ejscreensheetList:
    ejscreensheet.write(ejscreenItem[0], ejscreenItem[1], ejscreenItem[2], ejscreenItem[3])

#write EJSCREEN results to EJSCREEN sheet
#ejResults = {projID: {unitID: [est pop, % minority, % < high school: 4, % ling iso, % under 5, % over 64, % disabled] ...} ...}
#Create counter for keeping track of row numbers
currRow = 1
for currProj in sorted(ejResults):
    #retrieve name and area of current project
    currName = projDict[currProj][0] #projDict = {projID: [projName, vectorhectares]}
    currHectares = projDict[currProj][1]
    currAcres= round(projDict[currProj][1]*2.47105,2)

    #retrieve dictionary of results for current project
    currResults = ejResults[currProj]

    #iterate through census subunits in current dictionary of results and write to appropriate line of ALICE shreadsheet
    for ejscreenId in sorted(currResults):
        #retrieve current ALICE results
        currEjscreen = currResults[ejscreenId]

        ejscreensheet.write(currRow,0,currName,regForm) #projName
        ejscreensheet.write(currRow,1,currProj,regForm) #projID
        ejscreensheet.write(currRow,2,currHectares,regForm) #vectorhectares
        ejscreensheet.write(currRow,3,currAcres,regForm) #vectoracres
        ejscreensheet.write(currRow,4,ejscreenId,regForm) #censusId
        ejscreensheet.write(currRow,5,currEjscreen[0],numForm) #est pop
        ejscreensheet.write(currRow,6,currEjscreen[1],regForm) #% minority
        ejscreensheet.write(currRow,7,currEjscreen[2],regForm) #% < high school
        ejscreensheet.write(currRow,8,currEjscreen[3],regForm) #% linguistic isolation
        ejscreensheet.write(currRow,9,currEjscreen[4],regForm) #% under age 5
        ejscreensheet.write(currRow,10,currEjscreen[5],regForm) #% over age 64
        ejscreensheet.write(currRow,11,currEjscreen[6],regForm) #% disabled

        #advance row counter by 1
        currRow+=1

#apply conditional formatting to ejscreen results worksheets  #(first_row, first_col, last_row, last_col)
ejscreensheet.conditional_format(1,6,currRow-1,11, {'type': '3_color_scale', 'min_type': 'num', 'mid_type': 'num', 'max_type': 'num', 'min_value':0, 'mid_value':50, 'max_value':100, 'min_color': lgtOrange, 'mid_color': medOrange, 'max_color': drkOrange})

#close workbook to save
workbook.close()

#############################################################################################################################################################################################################################################################
#Clean up in_memory workspace and intermediate files
print "Cleaning up temporary data files ..."
arcpy.AddMessage("Cleaning up temporary data files ...")

arcpy.Delete_management("in_memory")

for item in toDelete:
    try:
        arcpy.Delete_management(item)
    except:
        print "Couldn't delete intermediate dataset in default geodatabase " + scratch + "; please clean up manually!"
        arcpy.AddMessage("Couldn't delete intermediate dataset in default geodatabase " + scratch + "; please clean up manually!")

#create copy of project shapes and results for the query tool archive
arcpy.CopyFeatures_management(projects,shapeArchive+"\\SATCommunitySnapshot_"+str(queName)+"_"+str(runStamp))
shutil.copy(outDir + "\\SATCommunitySnapshot_" + str(queName) + "_" + str(runStamp) + ".xlsx", resultsArchive+"\\SATCommunitySnapshot_"+str(queName)+"_"+str(runStamp)+".xlsx")


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
arcpy.AddMessage("took " + str(elapsed) + " minutes for " + str(len(projDict)) + " projects")