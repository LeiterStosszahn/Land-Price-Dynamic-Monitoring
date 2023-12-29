import time,os
import pandas as pd
from docxtpl import DocxTemplate
from xlsxtpl.writerx import BookWriter
from threading import Thread

time_now=time.localtime()

#Sample Sheet
sample_form_path=r"sample//评估方法-样表.xlsx"
sample_report_path=r"sample//报告模板.docx"

#Generate form
def transform_form(filling_org,person,city,year,quarter,complete_date,data_path,save_path):
	#Import data from excel
	form_data=pd.read_excel(data_path,sheet_name=r"表格信息",header=3)
	form_data=form_data.fillna("")
	form_lenrow=len(form_data)
	form_col_names=form_data.columns
	print(form_lenrow)
	print(form_col_names)
	#看最终上报结果是一张表还是多张表再决定读取与写入逻辑
#单张表：分方法读取，basic_col+方法col生成
#basic_col=["Stander_ID","Appraiser","Appraiser_ID","Land_use","Land_area","Setup_plot_ratio","Land_address","Complete_date"]
#多张表：整个读取替换，删除未使用方法表
	if __name__=="__main__":
		tkinter.messagebox.showinfo(title="生成报表",message="报表生成完成，共生成"+str(0)+"份报表")
	return 0

#Generate report
def transform_report(city,year,quarter,data_path,save_path):
	#Import data from excel
	report_data=pd.read_excel(data_path,sheet_name=r"报告信息",usecols="A:Z",header=1)
	report_data=report_data.fillna("")
	report_data["Evulate_time"]=report_data["Evulate_time"].apply(lambda x:x.strftime("%Y{}%m{}%d{}").format("年","月","日"))
	report_data.insert(loc=25,column="Years",value=year)
	report_lenrow=len(report_data)
	report_col_names=report_data.columns
	
	#Import sample report
	report=DocxTemplate(sample_report_path)

	#Replace targs
	num=0
	for i in range(report_lenrow):
		context=report_data.loc[i]
		contexts=dict(zip(report_col_names,context))
		
		title=city+year+"年第"+quarter+"季度"
		appraiser=context["Appraiser"]
		appraiser_ID="("+str(context["Appraiser_ID"])+")"
		stander=context["Stander_ID"]+"号标准宗地地价评估报告"
		
		report.render(context)
		
		dir_path=os.path.join(save_path,title+"监测成果","估价师成果",appraiser+appraiser_ID,title+"标准宗地地价评估报告"+appraiser_ID)
		if not os.path.exists(dir_path):
			os.makedirs(dir_path)
		report_save_path=os.path.join(dir_path,title+stander+appraiser_ID+".docx")
		
		report.save(report_save_path)
		num+=1

	if __name__=="__main__":
		tkinter.messagebox.showinfo(title="生成报告",message="报告生成完成，共生成"+str(num)+"份报告")
	return num

#GUI
if __name__=="__main__":
	import tkinter,math,tkinter.messagebox,configparser
	from tkinter import filedialog,ttk

	class MY_GUI():
		def __init__(self,init_window_name):
			self.init_window_name=init_window_name
		def set_init_window(self):
			self.init_window_name.title("城市地价动态监测成果包生成程序")
			#self.init_window_name.geometry("1000x681")

	def get_path(result,style="file"):
		if style=="directory":
			path=filedialog.askdirectory(title="请选择一个目录")
		elif style=="file":
			path=filedialog.askopenfilename(title="请选择一个文件")
		result.set(path)

	def check_and_run(result_type):
		checkstatue=1
		if year_text.get()=="":
			tkinter.messagebox.showerror(title="错误",message="年份不能为空")
			checkstatue=0
		if quarter_text.get()=="":
			tkinter.messagebox.showerror(title="错误",message="季度不能为空")
			checkstatue=0
		if path_text.get()=="":
			tkinter.messagebox.showerror(title="错误",message="未选择测算表")
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
		#All check pass
		if checkstatue:
			if result_type in ("report","all"):
				t_report=Thread(target=transform_report,args=(city_text.get(),year_text.get(),quarter_text.get(),path_text.get(),savepath_text.get()))
				t_report.start()
			if result_type in ("form","all"):
				tkinter.messagebox.showinfo(title="生成表格",message="生成表格")
				t_form=Thread(target=transform_form,args=(org_text.get(),person_text.get(),city_text.get(),year_text.get(),quarter_text.get(),date_text.get(),path_text.get(),savepath_text.get()))
				t_form.start()

	#Initialize outside config
	config=configparser.ConfigParser()
	config.read("config.ini",encoding="utf-8")
	
	#Initialize GUI
	main_GUI=tkinter.Tk()
	main_GUI.iconbitmap("resources//logo.ico")
	MY_GUI(main_GUI).set_init_window()

	#Label
	label_org=tkinter.Label(main_GUI,text="填表单位:",font=(30))
	label_org.grid(row=0,column=0,padx=(50,0),pady=(50,0),sticky="w")
	label_person=tkinter.Label(main_GUI,text="填表人:",font=(30))
	label_person.grid(row=0,column=2,pady=(50,0),sticky="w")
	label_city=tkinter.Label(main_GUI,text="城市：",font=(30))
	label_city.grid(row=0,column=4,pady=(50,0),sticky="w")
	label_year=tkinter.Label(main_GUI,text="年份：",font=(30))
	label_year.grid(row=1,column=0,padx=(50,0),pady=(25,0),sticky="w")
	label_quarter=tkinter.Label(main_GUI,text="季  度：",font=(30))
	label_quarter.grid(row=1,column=2,pady=(25,0),sticky="w")
	label_date=tkinter.Label(main_GUI,text="填表日期：",font=(30))
	label_date.grid(row=1,column=4,pady=(25,0),sticky="w")
	label_path=tkinter.Label(main_GUI,text="测算表路径：",font=(30))
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