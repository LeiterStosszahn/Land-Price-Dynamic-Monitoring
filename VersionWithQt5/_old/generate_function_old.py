import os,time,configparser
import pandas as pd
from docxtpl import DocxTemplate
from xlsxtpl.writerx import BookWriter
from concurrent.futures import ThreadPoolExecutor

from check_function import verify_form_col

#Config info
time_now=time.localtime()
hot_update=configparser.ConfigParser()
hot_update.read(r"bin\\hot_update.ini",encoding="utf-8")
config=configparser.ConfigParser()
config.read(r"config.ini",encoding="utf-8")

#Sample Sheet
sample_form_path=r"sample//评估方法-样表.xlsx"
sample_info_path=r"sample//基础信息-样表.xlsx"
sample_report_path=r"sample//报告模板.docx"
sample_cal_path=r"sample//汇总表-样表.xlsx"

#Transaction sample
def transform_sample(filling_org,person,city,year,quarter,date,data_path,save_path):
	num_trans=0
	sheet_name=hot_update.get("Import","trans_name")
	usecols=hot_update.get("Import","trans_data")
	header=hot_update.getint("Import","trans_header")
	try:
		trans_data=pd.read_excel(data_path,sheet_name=sheet_name,usecols=usecols,header=header)
	except:
		return [False,"无交易样点表"]
	else:
		sample_data=pd.read_excel(sample_cal_path,sheet_name=sheet_name,usecols=usecols,header=header)
		trans_data=trans_data.fillna("")
		trans_lenrow=len(trans_data)
		tarns_col_names=trans_data.columns
		#Verify
		verify=verify_form_col(tarns_col_names,sample_data)
		if verify:
			return [False,u"交易样点表"+verify]

		#Add data and change data form
		trans_data["Sell_date"]=trans_data["Sell_date"].apply(lambda x:x.strftime("%Y{}%m{}%d").format("-","-"))
		trans_data.insert(loc=49,column="filling_org",value=filling_org)
		trans_data.insert(loc=50,column="person",value=person)
		trans_data.insert(loc=51,column="date",value=date)
		tarns_col_names=trans_data.columns
		
		#Import sample sheet and creat floder
		sheet=BookWriter(sample_info_path)
		sheet.jinja_env.globals.update(dir=dir,getattr=getattr)
		title=city+year+"年第"+quarter+"季度"
		dir_path=os.path.join(save_path,title+"监测成果","技术承担单位","交易点登记表")
		if not os.path.exists(dir_path):
			os.makedirs(dir_path)

		#Replace targs
		for i in range(trans_lenrow):
			context=trans_data.loc[i]
			contexts=dict(zip(tarns_col_names,context))

			contexts["tpl_name"]="trans"
			contexts["sheet_name"]=u"交易点登记表"
			sheet.render_book(payloads=[contexts])

			trans_save_path=os.path.join(dir_path,city+"交易点登记表("+str(context["NO"])+").xlsx")
			sheet.save(trans_save_path)
			num_trans+=1

	return [True,u"交易样点汇总表生成完成,共生成"+str(num_trans)+u"份交易点登记表"]

#Generate appraiser form
def transform_appraiser(filling_org,city,year,quarter,data_path,save_path):
	sheet_name=hot_update.get("Import","form_name")
	usecols=hot_update.get("Import","form_data")
	header=hot_update.getint("Import","form_header")
	#Import data from excel
	try:
		form_data=pd.read_excel(data_path,sheet_name=sheet_name,usecols=usecols,header=header,converters={"Sample_NO1":str,"Sample_NO2":str,"Sample_NO3":str})
	except:
		return [False,u"无表格信息表"]
	sample_data=pd.read_excel(sample_cal_path,sheet_name=sheet_name,usecols=usecols,header=header)
	form_data=form_data.fillna(value=" ")
	form_col_names=form_data.columns

	#Verify
	verify=verify_form_col(form_col_names,sample_data,form_data)
	if verify:
		return [False,u"表格信息表"+verify]

	#Change data form
	percents=["Newness","Depreciation_rate","Com_reduction","Real_reduction","Land_reduction","Res_income_reduction","Residual_interest","Residual_rate","Residual_profit","Cost_profit","Cost_added","Cost_reduction","Cost_other"]
	for percent in percents:
		form_data[percent]=form_data[percent].apply(lambda x:x*100)
	dates=["sell_date","Evulate_time","Complete_date","Sample1_time","Sample2_time","Sample3_time","Res_Sample1_time","Res_Sample2_time","Res_Sample3_time"]
	for date in dates:
		form_data[date]=form_data[date].apply(lambda x:x.strftime("%Y{}%m{}%d").format("/","/") if type(x)!=str else x)
	sample_ids=[]

	#Import sample sheet and creat floder
	sheet_form=BookWriter(sample_form_path)
	sheet_form.jinja_env.globals.update(dir=dir,getattr=getattr)

	#Replace targs
	num=0
	appraiser_list=form_data["Appraiser"].unique()
	title=city+year+"年第"+quarter+"季度"

	for appraiser in appraiser_list:
		form_data_appraiser=form_data[form_data["Appraiser"]==appraiser]
		form_data_appraiser.reset_index(drop=True,inplace=True)
		lenrow=len(form_data_appraiser)
		appraiser=str(appraiser)
		appraiser_ID_str=str(form_data_appraiser["Appraiser_ID"].iloc[0])
		appraiser_ID="("+appraiser_ID_str+")"

		#Appraise method
		for i in range(lenrow):
			context=form_data_appraiser.loc[i]
			contexts=dict(zip(form_col_names,context))
			contexts["filling_org"]=filling_org
			contexts["Person"]=context["Appraiser"]
			methods=list(context[["Appraise_way1","Appraise_way2"]])

			stander=str(context["Stander_ID"])+"号监测点评估【"
			dir_path_up=os.path.join(save_path,title+"监测成果","估价师成果",appraiser+appraiser_ID)
			dir_path=os.path.join(dir_path_up,title+"技术要点表"+appraiser_ID)
			if not os.path.exists(dir_path):
				os.makedirs(dir_path)

			for method in methods:
				contexts["tpl_name"]=method
				contexts["sheet_name"]=method
				sheet_form.render_book(payloads=[contexts])
				form_save_path=os.path.join(dir_path,title+stander+method+"】技术要点表"+appraiser_ID+".xlsx")
				sheet_form.save(form_save_path)
				num+=1

	return [True,u"表格信息汇总表生成完成,共生成"+str(num)+u"份技术要点表"]

#Generate rent form
def transform_rent(filling_org,city,year,quarter,data_path,save_path):
	sheet_name=hot_update.get("Import","rent_name")
	usecols=hot_update.get("Import","rent_data")
	header=hot_update.getint("Import","rent_header")
	#Import data from excel
	try:
		form_data=pd.read_excel(data_path,sheet_name=sheet_name,usecols=usecols,header=header)
	except:
		return [False,u"无表格信息表"]
	sample_data=pd.read_excel(sample_cal_path,sheet_name=sheet_name,usecols=usecols,header=header)
	form_data=form_data.fillna(value=" ")
	form_col_names=form_data.columns

	#Verify
	verify=verify_form_col(form_col_names,sample_data,form_data)
	if verify:
		return [False,u"表格信息表"+verify]

	#Import sample sheet and creat floder
	sheet_form=BookWriter(sample_form_path)
	sheet_form.jinja_env.globals.update(dir=dir,getattr=getattr)

	#Replace targs
	rent_num=0
	appraiser_list=form_data["Appraiser"].unique()
	title=city+year+"年第"+quarter+"季度"
	rent_colname=["Stander_ID","Rent_price","Real_price","Price_note"]
	rent_contexts={"city":city,"year":year,"quarter":quarter,"tpl_name":"rent","sheet_name":"房地租金、房地产价格统计表"}

	for appraiser in appraiser_list:
		form_data_appraiser=form_data[form_data["Appraiser"]==appraiser]
		form_data_appraiser.reset_index(drop=True,inplace=True)
		lenrow=len(form_data_appraiser)
		appraiser=str(appraiser)
		appraiser_ID_str=str(form_data_appraiser["Appraiser_ID"].iloc[0])
		appraiser_ID="("+appraiser_ID_str+")"
		dir_path_up=os.path.join(save_path,title+"监测成果","估价师成果",appraiser+appraiser_ID)
		if not os.path.exists(dir_path_up):
				os.makedirs(dir_path_up)

		#Rent and price
		rent_contexts["rows"]=form_data_appraiser[rent_colname].values.tolist()
		rent_contexts["Appraiser_ID"]=appraiser_ID_str
		rent_contexts["date"]=form_data_appraiser["Complete_date"][0]

		sheet_form.render_book(payloads=[rent_contexts])
		rent_save_path=os.path.join(dir_path_up,title+"房地租金、房地产价格统计表"+appraiser_ID+".xlsx")
		sheet_form.save(rent_save_path)
		rent_num+=1

	return [True,u"租金、房地产价格统计表生成完成,共生成"+str(rent_num)+u"份租金、房地产价格统计表\r"]

#generate weight
def generate_weight(city,year,quarter,data_path,save_path):
	sheet_name=hot_update.get("Import","weight_name")
	usecols=hot_update.get("Import","weight_data")
	header=hot_update.getint("Import","weight_header")
	#Import data from excel
	try:
		form_data=pd.read_excel(data_path,sheet_name=sheet_name,usecols=usecols,header=header)
	except:
		return [False,u"无报告信息表"]
	sample_data=pd.read_excel(sample_cal_path,sheet_name=sheet_name,usecols=usecols,header=header)
	form_data=form_data.fillna(value=" ")
	form_col_names=form_data.columns

	#Verify
	verify=verify_form_col(form_col_names,sample_data,form_data)
	if verify:
		return [False,u"表格信息表"+verify]

	#Import sample sheet and creat floder
	sheet_form=BookWriter(sample_form_path)
	sheet_form.jinja_env.globals.update(dir=dir,getattr=getattr)

	#Replace targs
	num=0
	appraiser_list=form_data["Appraiser"].unique()
	title=city+year+"年第"+quarter+"季度"
	rent_colname=["Stander_ID","Method1","Method2","Reason"]
	rent_contexts={"tpl_name":"weight","sheet_name":"weight"}

	for appraiser in appraiser_list:
		form_data_appraiser=form_data[form_data["Appraiser"]==appraiser]
		form_data_appraiser.reset_index(drop=True,inplace=True)
		lenrow=len(form_data_appraiser)
		appraiser=str(appraiser)
		appraiser_ID_str=str(form_data_appraiser["Appraiser_ID"].iloc[0])
		appraiser_ID="("+appraiser_ID_str+")"
		dir_path_up=os.path.join(save_path,title+"监测成果","估价师成果",appraiser+appraiser_ID)
		if not os.path.exists(dir_path_up):
				os.makedirs(dir_path_up)

		#weight
		rent_contexts["rows"]=form_data_appraiser[rent_colname].values.tolist()

		sheet_form.render_book(payloads=[rent_contexts])
		rent_save_path=os.path.join(dir_path_up,"权重统计表"+".xlsx")
		sheet_form.save(rent_save_path)
		num+=1

	return [True,u"权重统计表生成完成,共生成"+str(num)+u"份权重统计表\r"]