import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd,xlwt

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
worksheet=new_worksheet.add_sheet(u'地价区段登记表',cell_overwrite_ok=True)

#合并单元
worksheet.write_merge(0,0,0,7,table.cell_value(0,0),sheet_style(u"仿宋_GB2312",22,True,True))#标题
worksheet.write_merge(1,8,0,0,table.cell_value(1,0),sheet_style(u"宋体",14,False,True,2,1,2,1))#区段基本情况
worksheet.write_merge(9,10,0,0,table.cell_value(9,0),sheet_style(u"宋体",14,False,True,2,1,1,1))#区段设定条件
worksheet.write(11,0,table.cell_value(11,0),sheet_style(u"宋体",14,False,True,2,1,1,2))#备注
worksheet.write_merge(13,13,0,7,table.cell_value(13,0),sheet_style(u"宋体",11,False,False))#填表说明
worksheet.write(1,1,table.cell_value(1,1),sheet_style(u"Nimbus Roman No9 L",14,False,True,1,1,2,1))#标号
for i in range(2,11):
	worksheet.write(i,1,table.cell_value(i,1),sheet_style(u"Nimbus Roman No9 L",14,False,True,1,1,1,1))
worksheet.write(11,1,table.cell_value(11,1),sheet_style(u"Nimbus Roman No9 L",14,False,True,1,1,1,2))
worksheet.write(1,2,table.cell_value(1,2),sheet_style(u"宋体",14,False,False,1,1,2,1))#名称
for i in range(2,11):
	worksheet.write(i,2,table.cell_value(i,2),sheet_style(u"宋体",14,False,False,1,1,1,1))
worksheet.write(11,2,table.cell_value(11,2),sheet_style(u"宋体",14,False,False,1,1,1,2))
worksheet.write_merge(11,11,3,7)#备注后面的空行
worksheet.write_merge(12,12,1,2,"",sheet_style(u"宋体",12,False,False))#填表单位
worksheet.write_merge(12,12,4,5,"",sheet_style(u"宋体",12,False,False))#填表人
for i in [0,3,6]:
	worksheet.write(12,i,table.cell_value(12,i),sheet_style(u"宋体",12,False,False))
sample.release_resources()

#宽度
worksheet.col(0).width=12*272
worksheet.col(1).width=int(8.38*272)
worksheet.col(2).width=29*272
for i in range(3,8):
	worksheet.col(i).width=int(11.75*272)
#行高
worksheet.row(0).height_mismatch=True
worksheet.row(0).height=int(35.25*20)
for i in range(1,13):
	worksheet.row(i).height_mismatch=True
	worksheet.row(i).height=24*20
worksheet.row(13).height_mismatch=True
worksheet.row(13).height=150*20

for layer in layer_input:
	with arcpy.da.SearchCursor(layer,['QDBM','CS','SZXZQ','QDLX','ZDKFLYMS','SZTDJB','TZQY','QDZMJ','SDRJL','SDKFCD','BZ']) as cursor:
		for row in cursor:
			worksheet.write_merge(1,1,3,7,row[0],sheet_style(u"宋体",14,False,False,1,2,2,1))
			for i in range(1,4)+range(5,10):
				worksheet.write_merge(i+1,i+1,3,7,row[i],sheet_style(u"宋体",14,False,False,1,2,1,1))
			worksheet.write_merge(11,11,3,7,row[10],sheet_style(u"宋体",14,False,False,1,2,1,2))
			#区段类型
			split=row[4].split("、")
			split_len=len(split)
			if split_len<=4:
				for i in range(0,split_len):
					worksheet.write(5,3+i,split[i],sheet_style(u"宋体",14,False,False,1,1,1,1))
				for i in range(split_len,4):
					worksheet.write(5,3+i,"",sheet_style(u"宋体",14,False,False,1,1,1,1))
				worksheet.write(5,7,"",sheet_style(u"宋体",14,False,False,1,2,1,1))
			else:
				for i in range(0,4):
					worksheet.write(5,3+i,split[i],sheet_style(u"宋体",14,False,False,1,1,1,1))
				for i in range(4,split_len):
					worksheet.write(5,7,"",sheet_style(u"宋体",14,False,False,1,2,1,1))
			worksheet.write(12,1,Filling_agency,sheet_style(u"宋体",12,False,False))
			worksheet.write(12,4,Filling_person,sheet_style(u"宋体",12,False,False))
			worksheet.write(12,7,Filling_date,sheet_style(u"宋体",12,False,False))

			new_worksheet.save(Save_route+'\\'+row[1]+quot+u'地价区段登记表('+row[0]+').xls')



