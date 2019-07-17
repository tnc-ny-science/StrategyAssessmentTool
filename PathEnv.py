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
