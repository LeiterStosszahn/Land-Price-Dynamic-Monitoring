import time,os,configparser,subprocess
import pandas as pd
from docxtpl import DocxTemplate
from xlsxtpl.writerx import BookWriter
from concurrent.futures import ThreadPoolExecutor

time_now=time.localtime()
hot_update=configparser.ConfigParser()
hot_update.read(r"bin//hot_update.ini",encoding="utf-8")

#Sample Sheet
sample_form_path=r"sample//评估方法-样表.xlsx"
sample_info_path=r"sample//基础信息-样表.xlsx"
sample_report_path=r"sample//报告模板.docx"
sample_cal_path=r"sample//汇总表-样表.xlsx"

#Form verify
def verify_form_col(input_form_col,sample_data,input_data="NA"):
	go_ahead=0
	sample_col_names=sample_data.columns
	diff=list(set(sample_col_names).difference(input_form_col))
	if diff==[] and type(input_data)==str:
		return go_ahead
	elif diff==[]:
		#Can insert another function here to verify logical relation
		#If logical relation unexcept
			#go_ahead="具体错误名称"
		return go_ahead
	else:
		return "缺少"+str(diff)+"列"

#Generate organization form
def transform_form_org(filling_org,person,city,year,quarter,date,data_path,save_path):
	#Transaction sample
	num_trans=0
	sheet_name=hot_update.get("Import","trans_name")
	usecols=hot_update.get("Import","trans_data")
	header=hot_update.getint("Import","trans_header")
	try:
		trans_data=pd.read_excel(data_path,sheet_name=sheet_name,usecols=usecols,header=header)
	except:
		if __name__=="__main__":
			tkinter.messagebox.showerror(title="无交易样点表",message="无交易样点表，无法输出：\n交易点登记表")
		return "无交易样点表"
	else:
		sample_data=pd.read_excel(sample_cal_path,sheet_name=sheet_name,usecols=usecols,header=header)
		trans_data=trans_data.fillna("")
		trans_lenrow=len(trans_data)
		tarns_col_names=trans_data.columns
		#Verify
		verify=verify_form_col(tarns_col_names,sample_data)
		if verify:
			if __name__=="__main__":
				tkinter.messagebox.showerror(title="交易样点表错误",message="交易样点表表格信息错误\n"+verify)
			return verify

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

	if __name__=="__main__":
		messages="估价机构报表生成完成，共生成：\r"
		if num_trans!=0:
			messages+=str(num_trans)+"份交易点登记表\r"		
		tkinter.messagebox.showinfo(title="生成报表",message=messages)
	return num_trans

#Generate appraiser form
def transform_form(filling_org,city,year,quarter,data_path,save_path):
	sheet_name=hot_update.get("Import","form_name")
	usecols=hot_update.get("Import","form_data")
	header=hot_update.getint("Import","form_header")
	#Import data from excel
	form_data=pd.read_excel(data_path,sheet_name=sheet_name,usecols=usecols,header=header,converters={"Sample_NO1":str,"Sample_NO2":str,"Sample_NO3":str})
	sample_data=pd.read_excel(sample_cal_path,sheet_name=sheet_name,usecols=usecols,header=header)
	form_data=form_data.fillna(value=" ")
	form_col_names=form_data.columns

	#Verify
	verify=verify_form_col(form_col_names,sample_data,form_data)
	if verify:
		if __name__=="__main__":
			tkinter.messagebox.showerror(title="汇总表错误",message="汇总表表格信息错误\n"+verify)
		return verify

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
	num=rent_num=0
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

		#Rent and price
		rent_contexts["rows"]=form_data_appraiser[rent_colname].values.tolist()
		rent_contexts["Appraiser_ID"]=appraiser_ID_str
		rent_contexts["date"]=form_data_appraiser["Complete_date"][0]

		sheet_form.render_book(payloads=[rent_contexts])
		rent_save_path=os.path.join(dir_path_up,title+"房地租金、房地产价格统计表"+appraiser_ID+".xlsx")
		sheet_form.save(rent_save_path)
		rent_num+=1

	if __name__=="__main__":
		tkinter.messagebox.showinfo(title="生成报表",message="估价师报表生成完成，共生成：\r"+str(num)+"份报估价师要点表\r"+str(rent_num)+"份租金、房地产价格统计表")
	return num,rent_num

#Generate report
def transform_report(city,year,quarter,data_path,save_path):
	sheet_name=hot_update.get("Import","report_name")
	usecols=hot_update.get("Import","report_data")
	header=hot_update.getint("Import","report_header")
	#Import data from excel
	report_data=pd.read_excel(data_path,sheet_name=sheet_name,usecols=usecols,header=header)
	sample_data=pd.read_excel(sample_cal_path,sheet_name=sheet_name,usecols=usecols,header=header)
	report_data=report_data.fillna("")
	report_lenrow=len(report_data)
	report_col_names=report_data.columns

	#Verify
	verify=verify_form_col(report_col_names,sample_data,report_data)
	if verify:
		if __name__=="__main__":
			tkinter.messagebox.showerror(title="汇总表错误",message="报告信息错误\n"+verify)
		return verify
	
	#Add data and change data form
	report_data["Evulate_time"]=report_data["Evulate_time"].apply(lambda x:x.strftime("%Y{}%m{}%d{}").format("年","月","日"))
	report_data.insert(loc=25,column="Years",value=year)
	report_col_names=report_data.columns

	#Import sample report
	report=DocxTemplate(sample_report_path)

	#Replace targs
	def save_report(report,dir_path,report_save_path):
		if not os.path.exists(dir_path):
			os.makedirs(dir_path)
		report.save(report_save_path)

	num=0
	pool_report=ThreadPoolExecutor()
	for i in range(report_lenrow):
		context=report_data.loc[i]
		
		title=city+year+"年第"+quarter+"季度"
		appraiser=str(context["Appraiser"])
		appraiser_ID="("+str(context["Appraiser_ID"])+")"
		stander=str(context["Stander_ID"])+"号标准宗地地价评估报告"
		
		report.render(context)
		
		dir_path=os.path.join(save_path,title+"监测成果","估价师成果",appraiser+appraiser_ID,title+"标准宗地地价评估报告"+appraiser_ID)
		report_save_path=os.path.join(dir_path,title+stander+appraiser_ID+".docx")
		pool_report.submit(save_report,report,dir_path,report_save_path)
		num+=1

	if __name__=="__main__":
		tkinter.messagebox.showinfo(title="生成报告",message="报告生成完成，共生成"+str(num)+"份报告")
	return num

#GUI
if __name__=="__main__":
	import tkinter,tkinter.messagebox,shutil,math
	from tkinter import filedialog,ttk,Menu

	class MY_GUI():
		def __init__(self,init_window_name):
			self.init_window_name=init_window_name
		def set_init_window(self):
			self.init_window_name.title("城市地价动态监测成果包生成程序")
			self.init_window_name.iconbitmap("bin//logo.ico")
			#self.init_window_name.geometry("1000x681")
		def quit_now(self):
			self.init_window_name.quit()
			self.init_window_name.destroy()

	def get_path(result,style="file"):
		if style=="directory":
			path=filedialog.askdirectory(title="请选择一个目录")
		elif style=="file":
			path=filedialog.askopenfilename(title="请选择一个文件", filetypes=(("excel文件",["*.xlsx","*.xls"]),))
		result.set(path)

	def download(sample_type):
		path=tkinter.StringVar()
		get_path(path,"directory")
		path=path.get()
		if sample_type=="sum" and path!="":
			shutil.copy(r"sample//汇总表-样表.xlsx",path)
			tkinter.messagebox.showinfo(title="下载",message="下载汇总表样表成功")
		elif sample_type=="cal" and path!="":
			shutil.copy(r"sample//测算表-样表.zip",path)
			tkinter.messagebox.showinfo(title="下载",message="下载测算表样表成功")
	
	def check_and_run(result_type):
		checkstatue=1
		if year_text.get()=="":
			tkinter.messagebox.showerror(title="错误",message="年份不能为空")
			checkstatue=0
		if quarter_text.get()=="":
			tkinter.messagebox.showerror(title="错误",message="季度不能为空")
			checkstatue=0
		if path_text.get()=="":
			tkinter.messagebox.showerror(title="错误",message="未选择汇总表")
			checkstatue=0
		if savepath_text.get()=="":
			tkinter.messagebox.showerror(title="错误",message="未选择保存路径")
			checkstatue=0
		if result_type!="report":
			if org_text.get()=="":
				tkinter.messagebox.showerror(title="错误",message="填表单位不能为空")
				checkstatue=0
			if person_text.get()=="":
				tkinter.messagebox.showerror(title="错误",message="填表人不能为空")
				checkstatue=0
		#All check pass
		if checkstatue:
			#Progress bar
			# progress=tkinter.Toplevel(master=main_GUI)
			# progress.title("生成中")
			# progress.geometry("150x50")
			# progressbar=tkinter.ttk.Progressbar(progress,length=200,mode="indeterminate",orient=tkinter.HORIZONTAL)
			# progressbar.pack(padx=5,pady=10)
			# progressbar.start()
			#Run function
			pool=ThreadPoolExecutor()
			if result_type in ("report","all"):
				t_report=pool.submit(transform_report,city_text.get(),year_text.get(),quarter_text.get(),path_text.get(),savepath_text.get())
			if result_type in ("form","all"):
				t_form=pool.submit(transform_form,org_text.get(),city_text.get(),year_text.get(),quarter_text.get(),path_text.get(),savepath_text.get())
				t_form_org=pool.submit(transform_form_org,org_text.get(),person_text.get(),city_text.get(),year_text.get(),quarter_text.get(),date_text.get(),path_text.get(),savepath_text.get())
			# if result_type=="all":
			# 	wait([t_report,t_form,t_form_org],return_when=ALL_COMPLETED)
			# elif result_type=="report":
			# 	wait([t_report],return_when=ALL_COMPLETED)
			# elif result_type=="form":
			# 	wait([t_form,t_form_org],return_when=ALL_COMPLETED)
			# progressbar.stop()
			# progress.destroy()

	#Initialize outside config
	config=configparser.ConfigParser()
	config.read(r"config.ini",encoding="utf-8")
	
	#Initialize GUI
	main_GUI=tkinter.Tk()
	MY_GUI(main_GUI).set_init_window()

	#Menu bar
	menubar=Menu(main_GUI)

	filemenu=Menu(menubar,tearoff=False)
	filemenu.add_command(label="下载汇总表模板",command=lambda:download("sum"))
	filemenu.add_command(label="下载测算表模板",command=lambda:download("cal"))
	filemenu.add_separator()
	filemenu.add_command(label="成果检验（未完成）",command=0)
	menubar.add_cascade(label="文件",menu=filemenu)

	online=Menu(menubar,tearoff=False)
	online.add_command(label="国家级地价监测系统",command=lambda:subprocess.run(hot_update.get("Online","CULP"),shell=True))
	online.add_command(label="成果包直接上报国家级地价监测系统（未完成）",command=0)
	online.add_separator()
	online.add_command(label="永业行地价监测系统（未完成）",command=lambda:subprocess.run(hot_update.get("Online","Realhom"),shell=True))
	menubar.add_cascade(label="在线系统",menu=online)

	# helpmenu=Menu(menubar,tearoff=False)
	# helpmenu.add_command(label="帮助")
	# helpmenu.add_separator()
	# helpmenu.add_command(label="关于")
	# menubar.add_cascade(label="帮助",menu=helpmenu)

	menubar.add_command(label="退出",command=lambda:MY_GUI(main_GUI).quit_now())
	main_GUI.config(menu=menubar)

	#Label
	label_org=tkinter.Label(main_GUI,text="填表单位:",font=(30))
	label_org.grid(row=0,column=0,padx=(50,0),pady=(50,0),sticky="w")
	label_person=tkinter.Label(main_GUI,text="填表人:",font=(30))
	label_person.grid(row=0,column=2,pady=(50,0),sticky="w")
	label_city=tkinter.Label(main_GUI,text="城    市：",font=(30))
	label_city.grid(row=0,column=4,pady=(50,0),sticky="w")
	label_year=tkinter.Label(main_GUI,text="年    份：",font=(30))
	label_year.grid(row=1,column=0,padx=(50,0),pady=(25,0),sticky="w")
	label_quarter=tkinter.Label(main_GUI,text="季  度：",font=(30))
	label_quarter.grid(row=1,column=2,pady=(25,0),sticky="w")
	label_date=tkinter.Label(main_GUI,text="填表日期：",font=(30))
	label_date.grid(row=1,column=4,pady=(25,0),sticky="w")
	label_path=tkinter.Label(main_GUI,text="汇总表路径：",font=(30))
	label_path.grid(row=2,column=0,padx=(50,0),pady=(25,0),sticky="w")
	label_savepath=tkinter.Label(main_GUI,text="保存路径：",font=(30))
	label_savepath.grid(row=3,column=0,padx=(50,0),pady=(25,0),sticky="w")

	#Input
	org_text=tkinter.StringVar()
	org_text.set(config.get("Default","organization"))
	enter_org=tkinter.Entry(main_GUI,textvariable=org_text,font=(30),width=30)
	enter_org.grid(row=0,column=1,pady=(50,0),sticky="w")

	person_text=tkinter.StringVar()
	person_text.set(config.get("Default","filling_person"))
	enter_person=tkinter.Entry(main_GUI,textvariable=person_text,font=(30),width=15)
	enter_person.grid(row=0,column=3,pady=(50,0),sticky="w")

	city_text=tkinter.StringVar()
	enter_city=ttk.Combobox(main_GUI,textvariable=city_text,font=(30),width=18)
	enter_city["value"]=(eval(config.get("GUI","city")))
	enter_city.current(0)
	enter_city.grid(row=0,column=5,padx=(0,50),pady=(50,0),sticky="w")

	year_text=tkinter.StringVar()
	year_text.set(time.strftime("%Y",time_now))
	enter_year=tkinter.Entry(main_GUI,textvariable=year_text,font=(30),width=30)
	enter_year.grid(row=1,column=1,pady=(25,0),sticky="w")

	quarter_text=tkinter.StringVar()
	quarter_text.set(math.ceil(int(time.strftime("%m",time_now))/3))
	enter_quarter=tkinter.Entry(main_GUI,textvariable=quarter_text,font=(30),width=15)
	enter_quarter.grid(row=1,column=3,pady=(25,0),sticky="w")

	#Date for the organization
	date_text=tkinter.StringVar()
	date_text.set(time.strftime("%Y/%m/%d",time_now))
	enter_date=tkinter.Entry(main_GUI,textvariable=date_text,font=(30),width=20)
	enter_date.grid(row=1,column=5,padx=(0,50),pady=(25,0),sticky="w")

	path_text=tkinter.StringVar()
	enter_path=tkinter.Entry(main_GUI,textvariable=path_text,font=(30),width=80,state="readonly")
	enter_path.grid(row=2,column=1,columnspan=4,pady=(25,0))

	savepath_text=tkinter.StringVar()
	enter_savepath=tkinter.Entry(main_GUI,textvariable=savepath_text,font=(30),width=80,state="readonly")
	enter_savepath.grid(row=3,column=1,columnspan=4,pady=(25,0))

	#Button
	button_path=tkinter.Button(main_GUI,text="选择路径",font=(30),width=10,command=lambda:get_path(path_text))
	button_path.grid(row=2,column=5,padx=(0,50),pady=(25,0),sticky="e")
	button_savepath=tkinter.Button(main_GUI,text="选择路径",font=(30),width=10,command=lambda:get_path(savepath_text,"directory"))
	button_savepath.grid(row=3,column=5,padx=(0,50),pady=(25,0),sticky="e")
	button_all=tkinter.Button(main_GUI,text="生成完整成果",font=(30),width=15,command=lambda:check_and_run("all"))
	button_all.grid(row=4,column=1,pady=(25,50),sticky="e")
	button_form=tkinter.Button(main_GUI,text="生成表格成果",font=(30),width=15,command=lambda:check_and_run("form"))
	button_form.grid(row=4,column=2,pady=(25,50))
	button_report=tkinter.Button(main_GUI,text="生成报告成果",font=(30),width=15,command=lambda:check_and_run("report"))
	button_report.grid(row=4,column=3,pady=(25,50),sticky="w")

	main_GUI.mainloop()