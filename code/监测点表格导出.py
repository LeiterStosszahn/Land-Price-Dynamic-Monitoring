import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd,xlwt,datetime

layer_input=arcpy.GetParameter(0)
sample_excel=arcpy.GetParameterAsText(1)
Filling_agency=arcpy.GetParameterAsText(2)
Filling_person=arcpy.GetParameterAsText(3)
Filling_date=arcpy.GetParameterAsText(4)
Save_route=arcpy.GetParameterAsText(5)
quot=arcpy.GetParameterAsText(6)

#Style
def sheet_style(font_name,height,bold,align,*borders):
	style=xlwt.XFStyle()
	
	font=xlwt.Font()
	font.name=font_name
	font.height=20*height
	font.bold=bold
	style.font=font
	
	alignment=xlwt.Alignment()
	alignment.wrap=1
	alignment.vert=0x01
	if align:
		alignment.horz=0x02
	style.alignment=alignment

	if len(borders):
		borders_style=xlwt.Borders()
		borders_style.left=borders[0]
		borders_style.right=borders[1]
		borders_style.top=borders[2]
		borders_style.bottom=borders[3]
		style.borders=borders_style
	
	return style

sample=xlrd.open_workbook(sample_excel)
table=sample.sheet_by_index(0)
new_worksheet=xlwt.Workbook(encoding="utf-8")
worksheet=new_worksheet.add_sheet(u'监测点登记表',cell_overwrite_ok=True)

#合并单元
worksheet.write_merge(0,0,0,6,table.cell_value(0,0),sheet_style(u"仿宋_GB2312",22,True,True))#标题
worksheet.write_merge(1,7,0,0,table.cell_value(1,0),sheet_style(u"宋体",14,False,True,2,1,2,1))#基本情况
worksheet.write_merge(8,14,0,0,table.cell_value(8,0),sheet_style(u"宋体",14,False,True,2,1,1,1))#权利状况
worksheet.write_merge(15,24,0,0,table.cell_value(15,0),sheet_style(u"宋体",14,False,True,2,1,1,1))#利用状况
worksheet.write_merge(25,30,0,0,table.cell_value(25,0),sheet_style(u"宋体",14,False,True,2,1,1,1))#影响因素
worksheet.write_merge(31,33,0,0,table.cell_value(25,0),sheet_style(u"宋体",14,False,True,2,1,1,1))#价格状况
worksheet.write(34,0,table.cell_value(34,0),sheet_style(u"宋体",14,False,True,2,1,1,2))#备注
worksheet.write_merge(36,36,0,6,table.cell_value(36,0),sheet_style(u"宋体",11,False,False))#填表说明
worksheet.write(1,1,table.cell_value(1,1),sheet_style(u"Nimbus Roman No9 L",14,False,True,1,1,2,1))#标号
for i in range(2,31):
	worksheet.write(i,1,table.cell_value(i,1),sheet_style(u"Nimbus Roman No9 L",14,False,True,1,1,1,1))
worksheet.write_merge(31,33,1,1,table.cell_value(31,1),sheet_style(u"Nimbus Roman No9 L",14,False,True,1,1,1,1))
worksheet.write(34,1,table.cell_value(34,1),sheet_style(u"Nimbus Roman No9 L",14,False,True,1,1,1,2))
worksheet.write_merge(1,1,2,5,table.cell_value(1,2),sheet_style(u"宋体",14,False,False,1,1,2,1))#名称
for i in range(2,31):
	worksheet.write_merge(i,i,2,5,table.cell_value(i,2),sheet_style(u"宋体",14,False,False,1,1,1,1))
worksheet.write_merge(31,33,2,2,table.cell_value(31,2),sheet_style(u"宋体",14,False,False,1,1,1,1))#曾发生的交易价格
for i in range(31,34):
	worksheet.write_merge(i,i,3,5,table.cell_value(i,3),sheet_style(u"宋体",14,False,False,1,1,1,1))#地面价等
worksheet.write_merge(34,34,2,6)#备注后面的空行
for i in [0,3,5]:
	worksheet.write(35,i,table.cell_value(35,i),sheet_style(u"宋体",12,False,False))#填表单位

sample.release_resources()

#宽度
worksheet.col(0).width=16*272
worksheet.col(1).width=12*272
worksheet.col(2).width=24*272
worksheet.col(3).width=int(8.25*272)
worksheet.col(4).width=int(11.25*272)
worksheet.col(5).width=int(10.13*272)
worksheet.col(6).width=50*272

#行高
worksheet.row(0).height_mismatch=True
worksheet.row(0).height=int(33*20)
for i in range(1,34):
	worksheet.row(i).height_mismatch=True
	worksheet.row(i).height=int(18.75*20)
worksheet.row(35).height_mismatch=True
worksheet.row(35).height=int(28.5*20)
worksheet.row(36).height_mismatch=True
worksheet.row(36).height=int(341.25*20)

title=["JCDBM","CS","SZXZQ","YTLX","SZTDJB","SZQDBM","TDWZ","TDSYZ","TDSYQLX","ZZSYYT","TXQL","ZZSYRQ","PZSYNX","TDSYZBH","TDSJYT","TDMJ",
"SDYTFTMJ","XZJZMJ","GHJZMJ","XZRJL","GHRJL","XZKFCD","GHXZ","ZYJZW","JSZXJL","LJZK","ZWJTTJ","ZWHJTJ","DZTJ","ZDXZ","DMJ","LMJ","JYRQ","BZ"]

for layer in layer_input:
	with arcpy.da.SearchCursor(layer,title) as cursor:
		for row in cursor:
			worksheet.write(1,6,row[0],sheet_style(u"宋体",14,False,False,1,2,2,1))
			for i in range(1,11)+range(12,32):
				worksheet.write(i+1,6,row[i],sheet_style(u"宋体",14,False,False,1,2,1,1))
			if row[11]!=None:
				worksheet.write(12,6,row[11].strftime('%Y/%m/%d'),sheet_style(u"宋体",14,False,False,1,2,1,1))
			else:
				worksheet.write(12,6,"",sheet_style(u"宋体",14,False,False,1,2,1,1))
			if row[32]!=None:
				worksheet.write(33,6,row[32].strftime('%Y/%m/%d'),sheet_style(u"宋体",14,False,False,1,2,1,1))
			else:
				worksheet.write(33,6,"",sheet_style(u"宋体",14,False,False,1,2,1,1))
			worksheet.write_merge(34,34,2,6,row[33],sheet_style(u"宋体",14,False,False,1,2,1,2))
			worksheet.write_merge(35,35,1,2,Filling_agency,sheet_style(u"宋体",12,False,False))
			worksheet.write(35,4,Filling_person,sheet_style(u"宋体",12,False,False))
			worksheet.write(35,6,Filling_date,sheet_style(u"宋体",12,False,False))

			new_worksheet.save(Save_route+'\\'+row[1]+quot+u'监测点登记表('+row[0]+').xls')





