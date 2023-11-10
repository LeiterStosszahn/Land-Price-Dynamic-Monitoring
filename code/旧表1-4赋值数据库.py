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
					row[0]=table.cell_value(3,2)#����λ��
					row[1]=table.cell_value(4,2)#����ʹ����
					row[2]=table.cell_value(5,2)#����ʹ��Ȩ����
					row[3]=table.cell_value(5,6)#����Ȩ��
					row[4]=table.cell_value(18,6)+table.cell_value(6,6)*365.25#��ֹʹ������
					row[5]=table.cell_value(6,2)#��׼ʹ������
					row[6]=table.cell_value(7,2)#����ʹ��֤��������Ȩ֤�����
					row[7]=table.cell_value(8,2)#����ʵ����;
					row[8]=table.cell_value(8,6)#���������ƽ���ף�TDMJ
					row[9]=table.cell_value(8,6)#�趨��;��̯�����ƽ���ף�
					row[10]=table.cell_value(9,2)#��״���������ƽ���ף�
					row[11]=table.cell_value(9,6)#�滮���������ƽ���ף�
					row[12]=table.cell_value(10,2)#��״�ݻ���
					row[13]=table.cell_value(10,6)#�滮�ݻ���
					row[14]=table.cell_value(11,2)#��״�����̶�
					row[15]=table.cell_value(12,2)#�滮����
					row[16]=table.cell_value(12,6)#��Ҫ������
					row[17]=table.cell_value(13,2)#�������ľ��루���
					row[18]=table.cell_value(13,6)#�ٽ�״��LJZK
					row[19]=table.cell_value(14,2)#��Χ��ͨ����
					row[20]=table.cell_value(14,6)#��Χ��������
					row[21]=table.cell_value(15,2)#��������
					row[22]=table.cell_value(15,6)#�ڵ���״
					if table.cell_value(22,2)!="":
						row[23]=table.cell_value(22,2)
						row[24]=row[23]*row[13]
					row[25]=table.cell_value(18,6)-(row[5]-table.cell_value(6,6))*365.25
					cursor.updateRow(row)




