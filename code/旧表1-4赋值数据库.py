import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd,datetime

layer_input=arcpy.GetParameter(0)
excel_input=arcpy.GetParameterAsText(1)
excel_input=excel_input.replace("'","")
excel_input=excel_input.split(";")

title=["TDWZ","TDSYZ","TDSYQLX","TXQL","ZZSYRQ","PZSYNX","TDSYZBH","TDSJYT","TDMJ",
"SDYTFTMJ","XZJZMJ","GHJZMJ","XZRJL","GHRJL","XZKFCD","GHXZ","ZYJZW","JSZXJL","LJZK",
"ZWJTTJ","ZWHJTJ","DZTJ","ZDXZ","DMJ","LMJ","JYRQ"]

def cuttime(a):
	now_date=datetime.strptime(a,"%y/%m/%d %H:%M:%S.%F")
	return now_date.strftime("%y/%m/%d")

for layer in layer_input:
	keywords=[]
	with arcpy.da.SearchCursor(layer,["JCDBM"]) as cursor:
		for row in cursor:
			if row[0]==' ' or row[0]==None or row[0]==0:
				continue
			if row[0] not in keywords:
				keywords.append(row[0])

	for keyword in keywords:
		expression=arcpy.AddFieldDelimiters(layer,"JCDBM")+'=\''+keyword+'\''
		with arcpy.da.UpdateCursor(layer,title,where_clause=expression) as cursor:
			for row in cursor:
				excel_path=[x for i,x in enumerate(excel_input) if x.find(keyword) != -1]
				if excel_path:
					excel=xlrd.open_workbook(excel_path[0])
					table=excel.sheet_by_index(0)
					row[0]=table.cell_value(3,2)#土地位置
					row[1]=table.cell_value(4,2)#土地使用者
					row[2]=table.cell_value(5,2)#土地使用权类型
					row[3]=table.cell_value(5,6)#他项权利
					row[4]=table.cell_value(18,6)+table.cell_value(6,6)*365.25#终止使用日期
					row[5]=table.cell_value(6,2)#批准使用年限
					row[6]=table.cell_value(7,2)#土地使用证（不动产权证）编号
					row[7]=table.cell_value(8,2)#土地实际用途
					row[8]=table.cell_value(8,6)#土地面积（平方米）TDMJ
					row[9]=table.cell_value(8,6)#设定用途分摊面积（平方米）
					row[10]=table.cell_value(9,2)#现状建筑面积（平方米）
					row[11]=table.cell_value(9,6)#规划建筑面积（平方米）
					row[12]=table.cell_value(10,2)#现状容积率
					row[13]=table.cell_value(10,6)#规划容积率
					row[14]=table.cell_value(11,2)#现状开发程度
					row[15]=table.cell_value(12,2)#规划限制
					row[16]=table.cell_value(12,6)#主要建筑物
					row[17]=table.cell_value(13,2)#距市中心距离（公里）
					row[18]=table.cell_value(13,6)#临街状况LJZK
					row[19]=table.cell_value(14,2)#周围交通条件
					row[20]=table.cell_value(14,6)#周围环境条件
					row[21]=table.cell_value(15,2)#地质条件
					row[22]=table.cell_value(15,6)#宗地形状
					if table.cell_value(22,2)!="":
						row[23]=table.cell_value(22,2)
						row[24]=row[23]*row[13]
					row[25]=table.cell_value(18,6)-(row[5]-table.cell_value(6,6))*365.25
					cursor.updateRow(row)




