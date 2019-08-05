#-------------------------------------------------------------------------------
# Name:        Neighboring parcel ownership aggregations
# Purpose:     Identify and group neighboring parcels with common ownership for query by the the SAT query tool
#
# Author:      shannon.thol
#
# Created:     28/05/2019
# Copyright:   (c) shannon.thol 2019
# Licence:     <your licence>
#-------------------------------------------------------------------------------

#load necessary packages
import arcpy
from arcpy import env
import fuzzywuzzy
from fuzzywuzzy import fuzz
import time

arcpy.env.overwriteOutput = True

########################################################################################################################################################################################################################
#Specify criteria for finding neighboring parcels
dist = "50 Meters"
count = 1000

#Specify criteria for aggregating owernships in fuzzy string matching
score = 93

#########################################################################################################################################################################################################################
#Get path to parcel dataset
parcels = r"C:\Users\shannon.thol\Box Sync\NYSIP\NYSIP_Projects\Conservation_Dimensions\Owasco\ParcelsAggOwnership_DRAFT.gdb\Owasco_merge_Albers"

#Get path to scratch gdb
scratch = "C:\\Users\\shannon.thol\\Box Sync\\NYSIP\\NYSIP_Projects\\Conservation_Dimensions\\Owasco\\ParcelsScratch.gdb"

#Get name of ID field
idField = "ID"

#Get name of owner name field
ownField = "OWNERNAME"

#Specify name of the group field
groupField = "DISSID2"

#Specify name of confidence field
confidence = "CONFIDENCE2"
#########################################################################################################################################################################################################################

#Retrieve type for idField
fields = arcpy.ListFields(parcels)
fieldNames = [x.name for x in fields]
fieldTypes = [y.type for y in fields]
idPos = fieldNames.index(idField)
idType = fieldTypes[idPos]

#Retrieve count of features in parcel dataset
num = arcpy.GetCount_management(parcels)

#########################################################################################################################################################################################################################
#Create empty dictionary for storing ID numbers and owner names
ownDict = dict()

start = time.time()
#Iterate through features in parcel dataset, populating ownDict dictionary with parcel ID numbers and owner names
print "Creating reference dictionary of parcel IDs and owner names ..."
with arcpy.da.SearchCursor(parcels, [idField, ownField]) as cursor:
    for row in cursor:
        ownDict[row[0]] = row[1]

#Get list of keys (parcel IDs) in the owner dictionary
ids = ownDict.keys()
maxId = max(ids)


end = time.time()
print "Done creating dictionary, took " + str(round((end-start),2)) + " seconds"

#########################################################################################################################################################################################################################

#Add new fields to parcel dataset
print "Adding fields to store group IDs ..."
if groupField in fieldNames:
    pass
else:
    arcpy.AddField_management(parcels, groupField, idType)

#Create empty dictionary for storing parcel grouping information in the form #groupPar = {group ID = [parcel ID1, parcel ID2, etc.]}
groupPar = dict()

#Create empty dictionary for storing parcel group assignments in the form #parGroup = {parcel ID: group ID}
parGroup = dict()

#Set up counter and group ID
i = 1
groupId = maxId + i

#Iterate through features in parcel dataset again, identifying near features and assessing similarity of their owner names
print "Determining group assignments ..."
with arcpy.da.SearchCursor(parcels, [idField, ownField]) as cursor:
    for row in cursor:

        #Define variables for fields
        focId = row[0]
        focOwn = row[1]

        print "Working on parcel ID: " + str(focId)
        start = time.time()

        #If the current focal parcel owner name is blank, give it a unique group id and move on (it's not possible to proceed with neighbor analysis because there is no name)
        if focOwn == '' or focOwn == ' ' or focOwn is None:
            parGroup[focId] = groupId
            groupPar[groupId] = [focId]
            groupId+=1

        #If the current focal parcel owner name is NOT blank, proceed with creating a feature layer and finding neighbor parcels
        else:
            #Create empty list for storing paths of datasets to delete
            toDel = list()

            #Create layer from current feature
            sql = '"' + "ID" + '" = ' + str(focId)
            focLyr = arcpy.MakeFeatureLayer_management(parcels, "focalPar", sql)
            focPar = scratch + "\\par" + str(focId)
            arcpy.CopyFeatures_management(focLyr, focPar)
            toDel.append(focPar)

            #Identify nearest parcels via generate near table operation using parameters specified above
            nearPar = scratch + "\\par" + str(focId) + "near"
            arcpy.GenerateNearTable_analysis(focPar, parcels, nearPar, dist, "NO_LOCATION", "NO_ANGLE", "ALL", count, "PLANAR")
            toDel.append(nearPar)

            #Create empty list for storing near features that have similar names
            currGroup = list()

            #Iterate through results in the near table, retrieving the near feature's FIDs
            with arcpy.da.SearchCursor(nearPar, 'NEAR_FID') as nearCursor:
                for nearRow in nearCursor:
                    nearId = nearRow[0]

                    #If the current near parcel ID is the same as the current focal parcel ID, pass as this is identifying the focal polygon in the parcel layer
                    if nearId == focId:
                        pass

                    #If the current near parcel ID is NOT the same as the current focal parcel ID, retrieve the owner name for the current near parcel ID from the ownership dictionary (ownDict)
                    else:
                        nearOwn = ownDict[nearId]

                        #If the current near parcel owner name is blank, pass as we cannot aggregate it based on ownership
                        if nearOwn == '' or nearOwn == ' ' or nearOwn is None:
                            pass

                        #If the current near parcel owner name is NOT blank, compare it to the focal parcel owner name using fuzzy string matching
                        else:
                            sortScore = fuzz.partial_token_sort_ratio(focOwn, nearOwn)
                            ratioScore = fuzz.partial_ratio(focOwn, nearOwn)

                            currScore = max([sortScore, ratioScore])

                            #If the current score is greater than the specified criteria score, proceed with adding the near parcel id to the current group list
                            if currScore > score:
                                currGroup.append(nearId)

                            #If the current score is not greater than the specified criteria score, pass
                            else:
                                pass

#####################################################################################################################################################################################################################
            #if the current group list has one or more entries (one or more neighbors), check to see if any of the group members have already been assigned to a group
            if len(currGroup) > 0:

##                print "neighbors of focal parcel (currGroup):"
##                print currGroup
                #create a new empty list for storing ids that have already been assigned to members of the current group
                groups = list()

                for currMem in currGroup:
                    #check to see if the current group member has already been assigned to a group in parGroup
                    if currMem in parGroup:
                        #if it has, retrieve the group number and add it to the groups list
                        currGroupId = parGroup[currMem]
                        groups.append(currGroupId)

                groups = list(set(groups))

                #if the group list is empty (no members of the current group have already been assigned a group id), proceed with assigning the focal parcel and all members of the group to a new group id
                if len(groups) == 0:
##                    print "none of the neighbors have already been assigned to a group - new group number assigned:"
##                    print groupId
                    parGroup[focId] = groupId
                    groupPar[groupId] = [focId]

                    for currMem in currGroup:
                        parGroup[currMem] = groupId
                        if groupId in groupPar:
                            if currMem in groupPar[groupId]:
                                pass
                            else:
                                groupPar[groupId].append(currMem)
                        else:
                            groupPar[groupId] = [currMem]
                        groupId+=1

                #if the group list has one or more items (members of the group have already been assigned to one or more group id), retrieve the smallest group id that has already been used and assign all members of the group to this group id
                else:
##                    print "neighbors have already been assigned to other group(s) of ids:"
##                    print groups

                    #set up new list for collecting group ids of other members of the group(s)
                    otherGroups = list()

                    #get group numbers of all members of these other group(s)
                    for group in groups:
                        members = groupPar[group]
                        for member in members:
                            other = parGroup[member]
                            otherGroups.append(other)

                    #retrieve the smallest group id that has already been used
                    assignedGroup = min(groups)

                    #assign the focal parcel to this group id
                    parGroup[focId] = assignedGroup
                    if focId in groupPar[assignedGroup]:
                        pass
                    else:
                        groupPar[assignedGroup].append(focId)

                    #assign all previously identified members of the group to this group id
                    for currMem in currGroup:
                        parGroup[currMem] = assignedGroup
                        if currMem in groupPar[assignedGroup]:
                            pass
                        else:
                            groupPar[assignedGroup].append(currMem)


                    #retrieve all members that have been assigned to the other group ids and assign their members to the smallest group id (called assignedGroup)
                    for other in groups:
                        if other == assignedGroup:
                            pass
                        else:
                            needReassigned = groupPar[other] #this is the list of members that have already been assigned to the current other group
                            for reassign in needReassigned:
                                parGroup[reassign] = assignedGroup
                                if reassign in groupPar[assignedGroup]:
                                    pass
                                else:
                                    groupPar[assignedGroup].append(reassign)

                            groupPar.pop(other)


            #If the current group list is empty (the focal parcel has no neighbors), proceed with assigning it a new group id
            elif len(currGroup) == 0:
                parGroup[focId] = groupId
                groupPar[groupId] = [focId]
##                print "the current parcel has no neighbors, assigning new group id:"
##                print groupId
                groupId+=1


        #Clean up intermediate datasets
        for item in toDel:
            arcpy.Delete_management(item)
        stop = time.time()
        elapsed = round(stop - start,1)
        print "  took " + str(elapsed) + " seconds"

##        print "parcel: group"
##        print parGroup
##        print "group: parcel"
##        print groupPar
##        print 'round done'

print "Writing results to attribute table ..."
with arcpy.da.UpdateCursor(parcels, [idField, groupField]) as upCursor:
    for upRow in upCursor:
        currParc = upRow[0]
        currParcGroup = parGroup[currParc]
        upRow[1] = currParcGroup
        upCursor.updateRow(upRow)



print 'All done!'
