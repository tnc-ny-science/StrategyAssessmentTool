#-------------------------------------------------------------------------------
# Name:        PercentileRanks
# Purpose:     Script for calculating percentile ranks (reference data for calculating SAT efficiency scores by the SAT query tool)
#
# Author:      shannon.thol
#
# Created:     15/09/2017
# Copyright:   (c) shannon.thol 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#############################################################################################################################################
#specify name of attribute
attribute = 'terrestrialconnectivity'

#Enter information about neighborhood size and radius

sizeDict = {'pt5':1, 1:2, 2:3, 5:4, 8:5, 10:6, 20:8, 40:12, 50:13, 100:19, 200:27, 250:30, 300:33, 500:42, 1000:59, 2000:84, 3000:103, 5000:133, 8000:168, 10000:188, 20000:266, 30000:325}
colDict = {'pt5':1, 1:2, 2:3, 5:4, 8:5, 10:6, 20:7, 40:8, 50:9, 100:10, 200:11, 250:12, 300:13, 500:14, 1000:15, 2000:16, 3000:17, 5000:18, 8000:19, 10000:20, 20000:21, 30000:22}

#############################################################################################################################################
done =[]

#Import libraries and set up environment
print "Preparing libraries and setting up environment ..."

import time
import arcpy
import os
import xlrd
import xlwt
import numpy as np
from arcpy import env
from arcpy.sa import *
arcpy.CheckOutExtension("Spatial")
arcpy.env.overwriteOutput = True

nlcd = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\General\\LandCover.gdb\\NLCD_2011_landcover_NYR"
arcpy.env.cellSize = 30
arcpy.env.snapRaster = nlcd
arcpy.env.outputCoordinateSystem = nlcd
extent = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\General\\Extents.gdb\\NYS_EXTENT_zero"

scratchDir = "D:\\gisdata\\Temp\\scratch_for_sthol"
scratchName = "scratch_nghbr.gdb"
scratchPath = "D:\\gisdata\\Temp\\scratch_for_sthol\\scratch_nghbr.gdb"

gridDir = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\ZonalStatsTool\\SensitivityGrids_x10e2.gdb"



######################################################################################################################################################################################################

arcpy.env.workspace = scratchPath

for size in sizeDict:
    if size in done:
        pass
    else:
        hectare = size
        print "Working on " + str(hectare) + " Ha reference data ..."
        radius = sizeDict[size]
        nghbrGrids = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\ZonalStatsTool\\SensitivityGrids_nghbr\\SensitivityGrids_" + str(hectare) + "Ha.gdb"
        attFolder = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\ZonalStatsTool\\SensitivityTables\\SensitivityTables_" + str(hectare) + "Ha_mean"
        pctFolder = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\ZonalStatsTool\\PercentileTables\\PercentileTables_" + str(hectare) + "Ha"

        grid = "\\" + attribute + "_x10e2_scope"
        startTime = time.time()
        gridOut = grid.split("_")[0]
        currentGrid = Raster(gridDir + grid)
        print currentGrid
        print "  Calculating focal stats ..."
        nghbr = FocalStatistics(currentGrid, NbrCircle(radius, "CELL"), "MEAN", "DATA")
        print "  Extracting results by mask ..."
        nghbrExt = ExtractByMask(nghbr, extent)
        nghbrExtInt = Int(nghbrExt)
        nghbrExtInt.save(nghbrGrids + "\\" + gridOut)
        arcpy.BuildRasterAttributeTable_management(nghbrGrids + "\\" + gridOut, "Overwrite")
        arcpy.Delete_management(nghbr)
        arcpy.Delete_management(nghbrExt)

        ################################################################################################################################################################################################
        #Export tables of neighborhood grids
        print "  Exporting attribute table ..."
        arcpy.env.workspace = nghbrGrids
        inputGrids = arcpy.ListRasters()
        arcpy.env.workspace = scratchPath

        for grid in inputGrids:
            if grid == attribute:
                inputName = nghbrGrids + "\\" + grid
                outName = grid
                arcpy.TableToTable_conversion(inputName, attFolder, outName)
            else:
                pass

        ################################################################################################################################################################################################
        #Convert dbf tables to excel workbooks
        print "  Converting dbf table to excel table ..."
        arcpy.env.workspace = attFolder
        inputTables = arcpy.ListTables()
        arcpy.env.workspace = scratchPath

        for table in inputTables:
            if table == attribute + '.dbf':
                inputName = attFolder + "\\" + table
                outName = attFolder + "\\" + table.split(".",1)[0] + ".xls"
                arcpy.TableToExcel_conversion(inputName, outName)
            else:
                pass

        #############################################################################################################################################

        #Create percentile tables from exported tables
        print "  Calculating percentiles ..."

##        try:
        inputTable = attFolder + "\\" + attribute + ".xls"
        outputTable = pctFolder + "\\" + attribute + ".xls"

        book = xlrd.open_workbook(inputTable)
        sheet = book.sheet_by_index(0)
        #Create list and array for values
        valueList = sheet.col_values(1)
        valuePop = valueList[1:]
        length = len(valuePop)
        valueArray = np.asarray(valuePop)
        adjvalArray = valueArray/100.0
        adjvalList = adjvalArray.tolist()

        #Create list and array for counts
        countList = sheet.col_values(2)
        countPop = countList[1:]
        countArray = np.asarray(countPop)

        #Create list of cumulative frequencies
        cumlFreqList = list()
        i = 0
        while i < length:
            cumlFreq = sum(countPop[:i+1])
            cumlFreqList.append(cumlFreq)
            i += 1
        cumlFreqArray = np.asarray(cumlFreqList)
        totalCount = cumlFreqList[length-1]

        #Calculate percentile rank for first row
        firstVal = adjvalList[0]
        firstCount = countPop[0]
        firstCumlFreq = cumlFreqList[0]

        firstPct = 100.0*(0.0+firstCount)/totalCount
        percentileList = list()
        percentileList.append(firstPct)

        #Calculate percentile ranks for subsequent rows
        j = 1
        while j < length:
            currentVal = adjvalList[j]
            currentCount = countPop[j]
            prevCumlFreq = cumlFreqList[j-1]
            j += 1
            currentPct = 100.0*(prevCumlFreq+currentCount)/totalCount
            percentileList.append(currentPct)

        #Write results to spreadsheet
        outBook = xlwt.Workbook()
        outSheet = outBook.add_sheet(attribute, cell_overwrite_ok = True)

        k = 0
        while k < length:
            currentVal = adjvalList[k]
            currentPct = percentileList[k]
            outSheet.write(k,0,currentVal) # row, column, value
            outSheet.write(k,1,currentPct)
            k += 1

        outBook.save(outputTable)
##        except:
##            print "  **couldn't calculate percentiles for " + attribute

    endTime = time.time()
    print "Elapsed time = " + str(int(endTime - startTime)/60) + " min"

#############################################################################################################################################
#Interpolate standardized percentiles
print "Standardizing percentiles ..."

#Set up Excel workbook for writing results
stdBook = xlwt.Workbook()
stdSheet = stdBook.add_sheet(attribute, cell_overwrite_ok = True)

for size in colDict:
    area = size
    currCol = colDict[size]

    pctFolder = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\ZonalStatsTool\\PercentileTables\\PercentileTables_" + str(area) + "Ha"

    masterTable = "D:\\gisdata\\Projects\\Regional\\ConservationDimensions\\ZonalStatsTool\\PercentileTables\\PercentileTables_MASTER_" + attribute + ".xls"

    stdVal = [0.0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08, 0.09, 0.1, 0.11, 0.12, 0.13, 0.14, 0.15, 0.16, 0.17, 0.18, 0.19, 0.2, 0.21, 0.22, 0.23, 0.24, 0.25, 0.26, 0.27, 0.28, 0.29, 0.3, 0.31, 0.32, 0.33, 0.34, 0.35, 0.36, 0.37, 0.38, 0.39, 0.4, 0.41, 0.42, 0.43, 0.44, 0.45, 0.46, 0.47, 0.48, 0.49, 0.5, 0.51, 0.52, 0.53, 0.54, 0.55, 0.56, 0.57, 0.58, 0.59, 0.6, 0.61, 0.62, 0.63, 0.64, 0.65, 0.66, 0.67, 0.68, 0.69, 0.7, 0.71, 0.72, 0.73, 0.74, 0.75, 0.76, 0.77, 0.78, 0.79, 0.8, 0.81, 0.82, 0.83, 0.84, 0.85, 0.86, 0.87, 0.88, 0.89, 0.9, 0.91, 0.92, 0.93, 0.94, 0.95, 0.96, 0.97, 0.98, 0.99, 1.0, 1.01, 1.02, 1.03, 1.04, 1.05, 1.06, 1.07, 1.08, 1.09, 1.1, 1.11, 1.12, 1.13, 1.14, 1.15, 1.16, 1.17, 1.18, 1.19, 1.2, 1.21, 1.22, 1.23, 1.24, 1.25, 1.26, 1.27, 1.28, 1.29, 1.3, 1.31, 1.32, 1.33, 1.34, 1.35, 1.36, 1.37, 1.38, 1.39, 1.4, 1.41, 1.42, 1.43, 1.44, 1.45, 1.46, 1.47, 1.48, 1.49, 1.5, 1.51, 1.52, 1.53, 1.54, 1.55, 1.56, 1.57, 1.58, 1.59, 1.6, 1.61, 1.62, 1.63, 1.64, 1.65, 1.66, 1.67, 1.68, 1.69, 1.7, 1.71, 1.72, 1.73, 1.74, 1.75, 1.76, 1.77, 1.78, 1.79, 1.8, 1.81, 1.82, 1.83, 1.84, 1.85, 1.86, 1.87, 1.88, 1.89, 1.9, 1.91, 1.92, 1.93, 1.94, 1.95, 1.96, 1.97, 1.98, 1.99, 2.0, 2.01, 2.02, 2.03, 2.04, 2.05, 2.06, 2.07, 2.08, 2.09, 2.1, 2.11, 2.12, 2.13, 2.14, 2.15, 2.16, 2.17, 2.18, 2.19, 2.2, 2.21, 2.22, 2.23, 2.24, 2.25, 2.26, 2.27, 2.28, 2.29, 2.3, 2.31, 2.32, 2.33, 2.34, 2.35, 2.36, 2.37, 2.38, 2.39, 2.4, 2.41, 2.42, 2.43, 2.44, 2.45, 2.46, 2.47, 2.48, 2.49, 2.5, 2.51, 2.52, 2.53, 2.54, 2.55, 2.56, 2.57, 2.58, 2.59, 2.6, 2.61, 2.62, 2.63, 2.64, 2.65, 2.66, 2.67, 2.68, 2.69, 2.7, 2.71, 2.72, 2.73, 2.74, 2.75, 2.76, 2.77, 2.78, 2.79, 2.8, 2.81, 2.82, 2.83, 2.84, 2.85, 2.86, 2.87, 2.88, 2.89, 2.9, 2.91, 2.92, 2.93, 2.94, 2.95, 2.96, 2.97, 2.98, 2.99, 3.0, 3.01, 3.02, 3.03, 3.04, 3.05, 3.06, 3.07, 3.08, 3.09, 3.1, 3.11, 3.12, 3.13, 3.14, 3.15, 3.16, 3.17, 3.18, 3.19, 3.2, 3.21, 3.22, 3.23, 3.24, 3.25, 3.26, 3.27, 3.28, 3.29, 3.3, 3.31, 3.32, 3.33, 3.34, 3.35, 3.36, 3.37, 3.38, 3.39, 3.4, 3.41, 3.42, 3.43, 3.44, 3.45, 3.46, 3.47, 3.48, 3.49, 3.5, 3.51, 3.52, 3.53, 3.54, 3.55, 3.56, 3.57, 3.58, 3.59, 3.6, 3.61, 3.62, 3.63, 3.64, 3.65, 3.66, 3.67, 3.68, 3.69, 3.7, 3.71, 3.72, 3.73, 3.74, 3.75, 3.76, 3.77, 3.78, 3.79, 3.8, 3.81, 3.82, 3.83, 3.84, 3.85, 3.86, 3.87, 3.88, 3.89, 3.9, 3.91, 3.92, 3.93, 3.94, 3.95, 3.96, 3.97, 3.98, 3.99, 4.0, 4.01, 4.02, 4.03, 4.04, 4.05, 4.06, 4.07, 4.08, 4.09, 4.1, 4.11, 4.12, 4.13, 4.14, 4.15, 4.16, 4.17, 4.18, 4.19, 4.2, 4.21, 4.22, 4.23, 4.24, 4.25, 4.26, 4.27, 4.28, 4.29, 4.3, 4.31, 4.32, 4.33, 4.34, 4.35, 4.36, 4.37, 4.38, 4.39, 4.4, 4.41, 4.42, 4.43, 4.44, 4.45, 4.46, 4.47, 4.48, 4.49, 4.5, 4.51, 4.52, 4.53, 4.54, 4.55, 4.56, 4.57, 4.58, 4.59, 4.6, 4.61, 4.62, 4.63, 4.64, 4.65, 4.66, 4.67, 4.68, 4.69, 4.7, 4.71, 4.72, 4.73, 4.74, 4.75, 4.76, 4.77, 4.78, 4.79, 4.8, 4.81, 4.82, 4.83, 4.84, 4.85, 4.86, 4.87, 4.88, 4.89, 4.9, 4.91, 4.92, 4.93, 4.94, 4.95, 4.96, 4.97, 4.98, 4.99, 5.0, 5.01, 5.02, 5.03, 5.04, 5.05, 5.06, 5.07, 5.08, 5.09, 5.1, 5.11, 5.12, 5.13, 5.14, 5.15, 5.16, 5.17, 5.18, 5.19, 5.2, 5.21, 5.22, 5.23, 5.24, 5.25, 5.26, 5.27, 5.28, 5.29, 5.3, 5.31, 5.32, 5.33, 5.34, 5.35, 5.36, 5.37, 5.38, 5.39, 5.4, 5.41, 5.42, 5.43, 5.44, 5.45, 5.46, 5.47, 5.48, 5.49, 5.5, 5.51, 5.52, 5.53, 5.54, 5.55, 5.56, 5.57, 5.58, 5.59, 5.6, 5.61, 5.62, 5.63, 5.64, 5.65, 5.66, 5.67, 5.68, 5.69, 5.7, 5.71, 5.72, 5.73, 5.74, 5.75, 5.76, 5.77, 5.78, 5.79, 5.8, 5.81, 5.82, 5.83, 5.84, 5.85, 5.86, 5.87, 5.88, 5.89, 5.9, 5.91, 5.92, 5.93, 5.94, 5.95, 5.96, 5.97, 5.98, 5.99, 6.0, 6.01, 6.02, 6.03, 6.04, 6.05, 6.06, 6.07, 6.08, 6.09, 6.1, 6.11, 6.12, 6.13, 6.14, 6.15, 6.16, 6.17, 6.18, 6.19, 6.2, 6.21, 6.22, 6.23, 6.24, 6.25, 6.26, 6.27, 6.28, 6.29, 6.3, 6.31, 6.32, 6.33, 6.34, 6.35, 6.36, 6.37, 6.38, 6.39, 6.4, 6.41, 6.42, 6.43, 6.44, 6.45, 6.46, 6.47, 6.48, 6.49, 6.5, 6.51, 6.52, 6.53, 6.54, 6.55, 6.56, 6.57, 6.58, 6.59, 6.6, 6.61, 6.62, 6.63, 6.64, 6.65, 6.66, 6.67, 6.68, 6.69, 6.7, 6.71, 6.72, 6.73, 6.74, 6.75, 6.76, 6.77, 6.78, 6.79, 6.8, 6.81, 6.82, 6.83, 6.84, 6.85, 6.86, 6.87, 6.88, 6.89, 6.9, 6.91, 6.92, 6.93, 6.94, 6.95, 6.96, 6.97, 6.98, 6.99, 7.0, 7.01, 7.02, 7.03, 7.04, 7.05, 7.06, 7.07, 7.08, 7.09, 7.1, 7.11, 7.12, 7.13, 7.14, 7.15, 7.16, 7.17, 7.18, 7.19, 7.2, 7.21, 7.22, 7.23, 7.24, 7.25, 7.26, 7.27, 7.28, 7.29, 7.3, 7.31, 7.32, 7.33, 7.34, 7.35, 7.36, 7.37, 7.38, 7.39, 7.4, 7.41, 7.42, 7.43, 7.44, 7.45, 7.46, 7.47, 7.48, 7.49, 7.5, 7.51, 7.52, 7.53, 7.54, 7.55, 7.56, 7.57, 7.58, 7.59, 7.6, 7.61, 7.62, 7.63, 7.64, 7.65, 7.66, 7.67, 7.68, 7.69, 7.7, 7.71, 7.72, 7.73, 7.74, 7.75, 7.76, 7.77, 7.78, 7.79, 7.8, 7.81, 7.82, 7.83, 7.84, 7.85, 7.86, 7.87, 7.88, 7.89, 7.9, 7.91, 7.92, 7.93, 7.94, 7.95, 7.96, 7.97, 7.98, 7.99, 8.0, 8.01, 8.02, 8.03, 8.04, 8.05, 8.06, 8.07, 8.08, 8.09, 8.1, 8.11, 8.12, 8.13, 8.14, 8.15, 8.16, 8.17, 8.18, 8.19, 8.2, 8.21, 8.22, 8.23, 8.24, 8.25, 8.26, 8.27, 8.28, 8.29, 8.3, 8.31, 8.32, 8.33, 8.34, 8.35, 8.36, 8.37, 8.38, 8.39, 8.4, 8.41, 8.42, 8.43, 8.44, 8.45, 8.46, 8.47, 8.48, 8.49, 8.5, 8.51, 8.52, 8.53, 8.54, 8.55, 8.56, 8.57, 8.58, 8.59, 8.6, 8.61, 8.62, 8.63, 8.64, 8.65, 8.66, 8.67, 8.68, 8.69, 8.7, 8.71, 8.72, 8.73, 8.74, 8.75, 8.76, 8.77, 8.78, 8.79, 8.8, 8.81, 8.82, 8.83, 8.84, 8.85, 8.86, 8.87, 8.88, 8.89, 8.9, 8.91, 8.92, 8.93, 8.94, 8.95, 8.96, 8.97, 8.98, 8.99, 9.0, 9.01, 9.02, 9.03, 9.04, 9.05, 9.06, 9.07, 9.08, 9.09, 9.1, 9.11, 9.12, 9.13, 9.14, 9.15, 9.16, 9.17, 9.18, 9.19, 9.2, 9.21, 9.22, 9.23, 9.24, 9.25, 9.26, 9.27, 9.28, 9.29, 9.3, 9.31, 9.32, 9.33, 9.34, 9.35, 9.36, 9.37, 9.38, 9.39, 9.4, 9.41, 9.42, 9.43, 9.44, 9.45, 9.46, 9.47, 9.48, 9.49, 9.5, 9.51, 9.52, 9.53, 9.54, 9.55, 9.56, 9.57, 9.58, 9.59, 9.6, 9.61, 9.62, 9.63, 9.64, 9.65, 9.66, 9.67, 9.68, 9.69, 9.7, 9.71, 9.72, 9.73, 9.74, 9.75, 9.76, 9.77, 9.78, 9.79, 9.8, 9.81, 9.82, 9.83, 9.84, 9.85, 9.86, 9.87, 9.88, 9.89, 9.9, 9.91, 9.92, 9.93, 9.94, 9.95, 9.96, 9.97, 9.98, 9.99, 10.0, 10.01, 10.02, 10.03, 10.04, 10.05, 10.06, 10.07, 10.08, 10.09, 10.1, 10.11, 10.12, 10.13, 10.14, 10.15, 10.16, 10.17, 10.18, 10.19, 10.2, 10.21, 10.22, 10.23, 10.24, 10.25, 10.26, 10.27, 10.28, 10.29, 10.3, 10.31, 10.32, 10.33, 10.34, 10.35, 10.36, 10.37, 10.38, 10.39, 10.4, 10.41, 10.42, 10.43, 10.44, 10.45, 10.46, 10.47, 10.48, 10.49, 10.5, 10.51, 10.52, 10.53, 10.54, 10.55, 10.56, 10.57, 10.58, 10.59, 10.6, 10.61, 10.62, 10.63, 10.64, 10.65, 10.66, 10.67, 10.68, 10.69, 10.7, 10.71, 10.72, 10.73, 10.74, 10.75, 10.76, 10.77, 10.78, 10.79, 10.8, 10.81, 10.82, 10.83, 10.84, 10.85, 10.86, 10.87, 10.88, 10.89, 10.9, 10.91, 10.92, 10.93, 10.94, 10.95, 10.96, 10.97, 10.98, 10.99, 11.0, 11.01, 11.02, 11.03, 11.04, 11.05, 11.06, 11.07, 11.08, 11.09, 11.1, 11.11, 11.12, 11.13, 11.14, 11.15, 11.16, 11.17, 11.18, 11.19, 11.2, 11.21, 11.22, 11.23, 11.24, 11.25, 11.26, 11.27, 11.28, 11.29, 11.3, 11.31, 11.32, 11.33, 11.34, 11.35, 11.36, 11.37, 11.38, 11.39, 11.4, 11.41, 11.42, 11.43, 11.44, 11.45, 11.46, 11.47, 11.48, 11.49, 11.5, 11.51, 11.52, 11.53, 11.54, 11.55, 11.56, 11.57, 11.58, 11.59, 11.6, 11.61, 11.62, 11.63, 11.64, 11.65, 11.66, 11.67, 11.68, 11.69, 11.7, 11.71, 11.72, 11.73, 11.74, 11.75, 11.76, 11.77, 11.78, 11.79, 11.8, 11.81, 11.82, 11.83, 11.84, 11.85, 11.86, 11.87, 11.88, 11.89, 11.9, 11.91, 11.92, 11.93, 11.94, 11.95, 11.96, 11.97, 11.98, 11.99, 12.0, 12.01, 12.02, 12.03, 12.04, 12.05, 12.06, 12.07, 12.08, 12.09, 12.1, 12.11, 12.12, 12.13, 12.14, 12.15, 12.16, 12.17, 12.18, 12.19, 12.2, 12.21, 12.22, 12.23, 12.24, 12.25, 12.26, 12.27, 12.28, 12.29, 12.3, 12.31, 12.32, 12.33, 12.34, 12.35, 12.36, 12.37, 12.38, 12.39, 12.4, 12.41, 12.42, 12.43, 12.44, 12.45, 12.46, 12.47, 12.48, 12.49, 12.5, 12.51, 12.52, 12.53, 12.54, 12.55, 12.56, 12.57, 12.58, 12.59, 12.6, 12.61, 12.62, 12.63, 12.64, 12.65, 12.66, 12.67, 12.68, 12.69, 12.7, 12.71, 12.72, 12.73, 12.74, 12.75, 12.76, 12.77, 12.78, 12.79, 12.8, 12.81, 12.82, 12.83, 12.84, 12.85, 12.86, 12.87, 12.88, 12.89, 12.9, 12.91, 12.92, 12.93, 12.94, 12.95, 12.96, 12.97, 12.98, 12.99, 13.0, 13.01, 13.02, 13.03, 13.04, 13.05, 13.06, 13.07, 13.08, 13.09, 13.1, 13.11, 13.12, 13.13, 13.14, 13.15, 13.16, 13.17, 13.18, 13.19, 13.2, 13.21, 13.22, 13.23, 13.24, 13.25, 13.26, 13.27, 13.28, 13.29, 13.3, 13.31, 13.32, 13.33, 13.34, 13.35, 13.36, 13.37, 13.38, 13.39, 13.4, 13.41, 13.42, 13.43, 13.44, 13.45, 13.46, 13.47, 13.48, 13.49, 13.5, 13.51, 13.52, 13.53, 13.54, 13.55, 13.56, 13.57, 13.58, 13.59, 13.6, 13.61, 13.62, 13.63, 13.64, 13.65, 13.66, 13.67, 13.68, 13.69, 13.7, 13.71, 13.72, 13.73, 13.74, 13.75, 13.76, 13.77, 13.78, 13.79, 13.8, 13.81, 13.82, 13.83, 13.84, 13.85, 13.86, 13.87, 13.88, 13.89, 13.9, 13.91, 13.92, 13.93, 13.94, 13.95, 13.96, 13.97, 13.98, 13.99, 14.0, 14.01, 14.02, 14.03, 14.04, 14.05, 14.06, 14.07, 14.08, 14.09, 14.1, 14.11, 14.12, 14.13, 14.14, 14.15, 14.16, 14.17, 14.18, 14.19, 14.2, 14.21, 14.22, 14.23, 14.24, 14.25, 14.26, 14.27, 14.28, 14.29, 14.3, 14.31, 14.32, 14.33, 14.34, 14.35, 14.36, 14.37, 14.38, 14.39, 14.4, 14.41, 14.42, 14.43, 14.44, 14.45, 14.46, 14.47, 14.48, 14.49, 14.5, 14.51, 14.52, 14.53, 14.54, 14.55, 14.56, 14.57, 14.58, 14.59, 14.6, 14.61, 14.62, 14.63, 14.64, 14.65, 14.66, 14.67, 14.68, 14.69, 14.7, 14.71, 14.72, 14.73, 14.74, 14.75, 14.76, 14.77, 14.78, 14.79, 14.8, 14.81, 14.82, 14.83, 14.84, 14.85, 14.86, 14.87, 14.88, 14.89, 14.9, 14.91, 14.92, 14.93, 14.94, 14.95, 14.96, 14.97, 14.98, 14.99, 15.0, 15.01, 15.02, 15.03, 15.04, 15.05, 15.06, 15.07, 15.08, 15.09, 15.1, 15.11, 15.12, 15.13, 15.14, 15.15, 15.16, 15.17, 15.18, 15.19, 15.2, 15.21, 15.22, 15.23, 15.24, 15.25, 15.26, 15.27, 15.28, 15.29, 15.3, 15.31, 15.32, 15.33, 15.34, 15.35, 15.36, 15.37, 15.38, 15.39, 15.4, 15.41, 15.42, 15.43, 15.44, 15.45, 15.46, 15.47, 15.48, 15.49, 15.5, 15.51, 15.52, 15.53, 15.54, 15.55, 15.56, 15.57, 15.58, 15.59, 15.6, 15.61, 15.62, 15.63, 15.64, 15.65, 15.66, 15.67, 15.68, 15.69, 15.7, 15.71, 15.72, 15.73, 15.74, 15.75, 15.76, 15.77, 15.78, 15.79, 15.8, 15.81, 15.82, 15.83, 15.84, 15.85, 15.86, 15.87, 15.88, 15.89, 15.9, 15.91, 15.92, 15.93, 15.94, 15.95, 15.96, 15.97, 15.98, 15.99, 16.0, 16.01, 16.02, 16.03, 16.04, 16.05, 16.06, 16.07, 16.08, 16.09, 16.1, 16.11, 16.12, 16.13, 16.14, 16.15, 16.16, 16.17, 16.18, 16.19, 16.2, 16.21, 16.22, 16.23, 16.24, 16.25, 16.26, 16.27, 16.28, 16.29, 16.3, 16.31, 16.32, 16.33, 16.34, 16.35, 16.36, 16.37, 16.38, 16.39, 16.4, 16.41, 16.42, 16.43, 16.44, 16.45, 16.46, 16.47, 16.48, 16.49, 16.5, 16.51, 16.52, 16.53, 16.54, 16.55, 16.56, 16.57, 16.58, 16.59, 16.6, 16.61, 16.62, 16.63, 16.64, 16.65, 16.66, 16.67, 16.68, 16.69, 16.7, 16.71, 16.72, 16.73, 16.74, 16.75, 16.76, 16.77, 16.78, 16.79, 16.8, 16.81, 16.82, 16.83, 16.84, 16.85, 16.86, 16.87, 16.88, 16.89, 16.9, 16.91, 16.92, 16.93, 16.94, 16.95, 16.96, 16.97, 16.98, 16.99, 17.0, 17.01, 17.02, 17.03, 17.04, 17.05, 17.06, 17.07, 17.08, 17.09, 17.1, 17.11, 17.12, 17.13, 17.14, 17.15, 17.16, 17.17, 17.18, 17.19, 17.2, 17.21, 17.22, 17.23, 17.24, 17.25, 17.26, 17.27, 17.28, 17.29, 17.3, 17.31, 17.32, 17.33, 17.34, 17.35, 17.36, 17.37, 17.38, 17.39, 17.4, 17.41, 17.42, 17.43, 17.44, 17.45, 17.46, 17.47, 17.48, 17.49, 17.5, 17.51, 17.52, 17.53, 17.54, 17.55, 17.56, 17.57, 17.58, 17.59, 17.6, 17.61, 17.62, 17.63, 17.64, 17.65, 17.66, 17.67, 17.68, 17.69, 17.7, 17.71, 17.72, 17.73, 17.74, 17.75, 17.76, 17.77, 17.78, 17.79, 17.8, 17.81, 17.82, 17.83, 17.84, 17.85, 17.86, 17.87, 17.88, 17.89, 17.9, 17.91, 17.92, 17.93, 17.94, 17.95, 17.96, 17.97, 17.98, 17.99, 18.0, 18.01, 18.02, 18.03, 18.04, 18.05, 18.06, 18.07, 18.08, 18.09, 18.1, 18.11, 18.12, 18.13, 18.14, 18.15, 18.16, 18.17, 18.18, 18.19, 18.2, 18.21, 18.22, 18.23, 18.24, 18.25, 18.26, 18.27, 18.28, 18.29, 18.3, 18.31, 18.32, 18.33, 18.34, 18.35, 18.36, 18.37, 18.38, 18.39, 18.4, 18.41, 18.42, 18.43, 18.44, 18.45, 18.46, 18.47, 18.48, 18.49, 18.5, 18.51, 18.52, 18.53, 18.54, 18.55, 18.56, 18.57, 18.58, 18.59, 18.6, 18.61, 18.62, 18.63, 18.64, 18.65, 18.66, 18.67, 18.68, 18.69, 18.7, 18.71, 18.72, 18.73, 18.74, 18.75, 18.76, 18.77, 18.78, 18.79, 18.8, 18.81, 18.82, 18.83, 18.84, 18.85, 18.86, 18.87, 18.88, 18.89, 18.9, 18.91, 18.92, 18.93, 18.94, 18.95, 18.96, 18.97, 18.98, 18.99, 19.0, 19.01, 19.02, 19.03, 19.04, 19.05, 19.06, 19.07, 19.08, 19.09, 19.1, 19.11, 19.12, 19.13, 19.14, 19.15, 19.16, 19.17, 19.18, 19.19, 19.2, 19.21, 19.22, 19.23, 19.24, 19.25, 19.26, 19.27, 19.28, 19.29, 19.3, 19.31, 19.32, 19.33, 19.34, 19.35, 19.36, 19.37, 19.38, 19.39, 19.4, 19.41, 19.42, 19.43, 19.44, 19.45, 19.46, 19.47, 19.48, 19.49, 19.5, 19.51, 19.52, 19.53, 19.54, 19.55, 19.56, 19.57, 19.58, 19.59, 19.6, 19.61, 19.62, 19.63, 19.64, 19.65, 19.66, 19.67, 19.68, 19.69, 19.7, 19.71, 19.72, 19.73, 19.74, 19.75, 19.76, 19.77, 19.78, 19.79, 19.8, 19.81, 19.82, 19.83, 19.84, 19.85, 19.86, 19.87, 19.88, 19.89, 19.9, 19.91, 19.92, 19.93, 19.94, 19.95, 19.96, 19.97, 19.98, 19.99, 20.0]

    try:
        print "Working on " + str(size) + " ... "
        start = time.time()
        pctTable = pctFolder + "\\" + attribute + ".xls"

        #write header line in worksheet; set up integer for writing results
        stdSheet.write(0,0, "Sens")  #row, column, value
        stdSheet.write(0,currCol, str(area) + "_Ha")  #row, column, value
        i = 0

        #open percentile table for reading inputs for current attribute
        inBook = xlrd.open_workbook(pctTable)
        inSheet = inBook.sheet_by_index(0)

        valList = list()
        pctList = list()
        valsList = list()

        #iterate through rows in the table, getting current value and percentile rank and adding to lists
        for row in range(inSheet.nrows):
            rowVals = inSheet.row_values(row)
            value = round(rowVals[0],2)
            pct = round(rowVals[1],2)
            valList.append(value)
            pctList.append(pct)
            valsList.append([value, pct])

        #convert lists to tuples so they are immutable
        valTuple = tuple(valList)
        pctTuple = tuple(pctList)
        minVal = min(valTuple)
        maxVal = max(valTuple)
        number = len(valTuple)

        #iterate through sensitivities in standard sensitivity list
        for currVal in stdVal:
            i = i + 1

            #if current sensitivity in value list, get corresponding percentile from pct list and assign is as currPct
            if currVal in valList:
                currValIndex = valList.index(currVal)
                currPct = pctTuple[currValIndex]
            #if current sensitivity is less than minimum sensitivity from table, assign currPct as 0
            elif currVal < minVal:
                currPct = 0.00

            #if current sensitivity is greater than maximum sensitivity from table, assign currPct as 100
            elif currVal > maxVal:
                currPct = 100.00

            else:
                #calculate absolute different tuple and get minimum difference pct and val
                absDiffTuple = tuple([abs(currVal - x) for x in valTuple])
                minAd = min(absDiffTuple)
                minAdIn = absDiffTuple.index(minAd)
                minAdPct = pctTuple[minAdIn]
                minAdVal = valTuple[minAdIn]

                #if the minAdIn is the highest number in the lists, or minAdVal is greater than currVal extrapolate based on minus 1 indices values
                if minAdIn == number - 1 or minAdVal > currVal:
                    oneMinPct = pctTuple[minAdIn - 1]
                    oneMinVal = valTuple[minAdIn - 1]
                    slope = (minAdPct - oneMinPct)/(minAdVal - oneMinVal)
                    currPct = round(minAdPct - slope*(minAdVal - currVal),2)

                #if the minAdIn is the lowest number in the lists, or minAdVal is less than currVal extrapolate based on plus 1 indices values
                elif minAdIn == 0 or minAdVal < currVal:
                    onePlsPct = pctTuple[minAdIn + 1]
                    onePlsVal = valTuple[minAdIn + 1]
                    slope = (onePlsPct - minAdPct)/(onePlsVal - minAdVal)
                    currPct = round(slope*(currVal - minAdVal) + minAdPct,2)

            #write current value (sensitivity) for current stdPct to column 0, and write currPct for current stdPct to currCol
            stdSheet.write(i,0,currVal)
            stdSheet.write(i,currCol,currPct)
    except:
        print "  **couldn't standardize percentiles for " + attribute

stdBook.save(masterTable)

#############################################################################################################################################


print "All done!"
print ''

