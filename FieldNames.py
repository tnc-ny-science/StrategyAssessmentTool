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
u'marinestormsurgemitigation_fn', u'marinestormsurgemitigation', u'heatmitigation', u'riverinefpfloodmitigation', u'riverinefpfloodmitigation_fn', u'riverinefpfloodmitigation_bp', u'riverinefpfloodmitigation_bv', u'flooddamageprevention_bp', u'flooddamageprevention_bv', u'marineaquaticrecreation_bp', u'marineaquaticrecreation_bv',
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
