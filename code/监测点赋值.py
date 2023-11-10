import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import random

layer_input=arcpy.GetParameter(0)
field_type=arcpy.GetParameterAsText(1)
layer_addin=arcpy.GetParameter(2)
field_addin=arcpy.GetParameterAsText(3)

def random_name(name,lenth):
	return name+"_"+"".join(random.sample('zyxwvutsrqponmlkjihgfedcba1234567890',lenth))

for layer in layer_input:
	#get workspace
	layer_des=arcpy.Describe(layer)
	catalog_path=layer_des.path#图层路径
	if arcpy.Describe(catalog_path).dataType=='FeatureDataset':
		arcpy.env.workspace=arcpy.Describe(catalog_path).path#数据库路径
	else:
		arcpy.env.workspace=catalog_path
	arcpy.AddMessage("设定临时工作空间："+arcpy.env.workspace)

	# keywords=[]
	# with arcpy.da.SearchCursor(layer,["JCDBM"]) as cursor:
	# 	for row in cursor:
	# 		if row[0]==' ' or row[0]==None or row[0]==0:
	# 			continue
	# 		if row[0] not in keywords:
	# 			keywords.append(row[0])

	if field_type in ("SZXZQ","SZTDJB","SZQDBM"):
		join_name=random_name("join",5)
		arcpy.SpatialJoin_analysis(layer,layer_addin,join_name,"JOIN_ONE_TO_ONE","KEEP_ALL","#","WITHIN") 
		arcpy.AddJoin_management(layer,"JCDBM",join_name,"JCDBM")
		
		if field_type=="SZXZQ":
			arcpy.management.CalculateField(layer,str(layer)+".SZXZQ","!"+join_name+"."+field_addin+"!","PYTHON_9.3")
		elif field_type=="SZTDJB":
			arcpy.management.CalculateField(layer,str(layer)+".SZTDJB","!"+join_name+"."+field_addin+"!","PYTHON_9.3")
		elif field_type=="SZQDBM":
			arcpy.management.CalculateField(layer,str(layer)+".SZQDBM","!"+join_name+"."+field_addin+"!","PYTHON_9.3")
		
		arcpy.RemoveJoin_management(layer,join_name)
		arcpy.Delete_management(join_name)

	elif field_type=="JSZXJL":
		distance_name=random_name("distance",5)
		arcpy.PointDistance_analysis(layer,layer_addin,distance_name)
		distance_sort_name=random_name("distance_sort",5)
		arcpy.Sort_management(distance_name,distance_sort_name,[["DISTANCE","ASCENDING"]])
		arcpy.Delete_management(distance_name)
		arcpy.DeleteIdentical_management(distance_sort_name,"INPUT_FID")
		arcpy.AddJoin_management(layer,"OBJECTID",distance_sort_name,"INPUT_FID")
		arcpy.management.CalculateField(layer,str(layer)+".JSZXJL","round(!"+distance_sort_name+".DISTANCE!/1000,2)","PYTHON_9.3")#非精确四舍五入
		arcpy.RemoveJoin_management(layer,distance_sort_name)
		arcpy.Delete_management(distance_sort_name)