import configparser,time
import pandas as pd
from concurrent.futures import ThreadPoolExecutor,as_completed

#Config info
time_now=time.localtime()
hot_update=configparser.ConfigParser()
hot_update.read(r"bin\\hot_update.ini",encoding="utf-8")
config=configparser.ConfigParser()
config.read(r"config.ini",encoding="utf-8")
empty=[0,"0"," ","",None]

#Sample Sheet
sample_cal_path=r"sample//汇总表-样表.xlsx"

#Form verify
def verify_form_col(input_form_col,sample_data):
	go_ahead=0
	sample_col_names=sample_data.columns
	diff=list(set(sample_col_names).difference(input_form_col))
	if diff==[]:
		return go_ahead
	else:
		return "缺少"+str(diff)+"列"

def cal_difference(a,b):
	return max(a,b)/min(a,b)-1

#show check result
def show_check(check_content,obj_list,check_percent,progressBar,result,textBrowser,text):
	#check appraiser result
	result_final=0
	#Waiting until the thread pool finished
	for future in as_completed(obj_list):
		check_percent+=0.5
		progressBar.setValue(check_percent)
	#show result
	for content in check_content:
		err_num=result[content][0]
		result_final+=err_num
		if err_num:
			sort=result[content][1:]
			sort.sort()
			for i in range(0,err_num+2):
				textBrowser.append(sort[i])
	if not result_final:
		textBrowser.append(text+u"检查无误")
	return check_percent,result_final

#Integrality
class verity_integrality(object):
	def __init__(self,data_path,ini_value):
		#transform_sample
		self.data_path=data_path
		self.sheet_name=hot_update.get("Import",ini_value+"_name")
		self.usecols=hot_update.get("Import",ini_value+"_data")
		self.header=hot_update.getint("Import",ini_value+"_header")
		self.data=[]
		self.lenrow=0
		if ini_value=="trans":
			self.gtype="交易样点"
		elif ini_value in ["form","rent"]:
			self.gtype="表格信息"
		else:
			self.gtype="报告信息"

	def verify(self):
		try:
			report_data=pd.read_excel(self.data_path,sheet_name=self.sheet_name,usecols=self.usecols,header=self.header)
		except:
			return [False,u"无"+self.gtype+u"表"]
		else:
			sample_data=pd.read_excel(sample_cal_path,sheet_name=self.sheet_name,usecols=self.usecols,header=self.header)
			report_data=report_data.fillna("")
			report_lenrow=len(report_data)
			tarns_col_names=report_data.columns
			#Verify
			verify=verify_form_col(tarns_col_names,sample_data)
			if verify:
				return [False,self.gtype+u"表"+verify]
			else:
				self.data=report_data
				self.lenrow=report_lenrow
				return [True,""]

#check report table
class verify_report(object):
	def __init__(self,path):
		self.data=pd.read_excel(
			path,
			sheet_name=hot_update.get("Import","report_name"),
			usecols=hot_update.get("Import","report_data"),
			header=int(hot_update.get("Import","report_header"))
		)
		self.lenrow=len(self.data)

	def check_weight(self):
		result=[0]
		for i in range(0,self.lenrow):
			if self.data["Method1"][i]+self.data["Method2"][i]!=1:
				result[0]+=1
				Stander_ID="<b>"+self.data["Stander_ID"][i]+"</b>"
				Appraiser="<b>"+self.data["Appraiser"][i]+"</b>"
				result.append(u"第"+str(i+3)+u"行，标准宗地编号为 "+Stander_ID+u" 估价师为 "+Appraiser+u" 的两种方法权重之和不为1")
		return result


#check form table
class verify_form(object):
	def __init__(self,path):
		data_origin=pd.read_excel(
			path,
			sheet_name=hot_update.get("Import","form_and_rent_name"),
			usecols=hot_update.get("Import","form_and_rent_data"),
			header=int(hot_update.get("Import","form_and_rent_header")),
			dtype={"Sample_NO1":str, "Sample_NO2":str, "Sample_NO3":str,
		  			"Sample1_weight":str,"Sample2_weight":str,"Sample3_weight":str}
		)
		self.data=data_origin.fillna("")
		self.lenrow=len(self.data)
		self.standers=self.data["Stander_ID"].unique()
		self.pool=ThreadPoolExecutor()
		self.result={
			#appraiser
			"difference":[0,"",u"<b>·估价方法结果差值问题</b>"],
			"Sample_NO":[0,"",u"<b>·比较法使用了相同的案例</b>"],
			"Comparision_result":[0,"",u"<b>·比较法结果问题</b>"],
			#stander
			"Stander_ID":[0,"",u"<b>·两个估价师结果差值问题</b>"],
		}

	def result_appraiser(self):
		#Check the difference of appraising result
		def result_difference(i,Stander_ID,Appraiser):
			Result_way1=float(self.data["Result_way1"][i])
			Result_way2=float(self.data["Result_way2"][i])
			difference=0
			if Result_way1!=0 and Result_way2!=0:
				difference=cal_difference(Result_way1,Result_way2)
			if Result_way1 in empty and Result_way2 in empty:
				self.result["difference"][0]+=1
				self.result["difference"].append(
					u"第"+str(i+5)+u"行，标准宗地编号为 "+Stander_ID+u" 估价师为 "+Appraiser+u" 的两种方法评估结果均为0"
				)
			elif difference>0.2:
				self.result["difference"][0]+=1
				self.result["difference"].append(
					u"第"+str(i+5)+u"行，标准宗地编号为 "+Stander_ID+u" 估价师为 "+Appraiser+u" 的两种方法评估结果差值为"+str(round(difference*100,2))+u"%，大于20%"
				)

		#Market comparision method
		def market_comparision(i,Stander_ID,Appraiser,Appraise_way):
			if u"市场比较法" in Appraise_way:
				Sample_NO1=self.data["Sample_NO1"][i]
				Sample_NO2=self.data["Sample_NO2"][i]
				Sample_NO3=self.data["Sample_NO3"][i]
				Comparison_land=self.data["Comparison_land"][i]
				Comparision_floor=self.data["Comparision_floor"][i]
				Land_use=self.data["Land_use"][i]
				land_type={"工业":"industry","住宅":"resident","商服":"commerce"}
				stander_ratio=float(config.get("stander_plot_ratio",land_type.get(Land_use,"other")))
				#wheter the same sample is existing
				if Sample_NO1==Sample_NO2 or Sample_NO1==Sample_NO3 or Sample_NO2==Sample_NO3:
					self.result["Sample_NO"][0]+=1
					self.result["Sample_NO"].append(
						u"第"+str(i+5)+u"行，标准宗地编号为 "+Stander_ID+u" 估价师为 "+Appraiser+u" 的市场比较法中存在重复使用的案例"
					)
				#correctness of the result
				if Comparison_land in empty or Comparision_floor in empty:
					self.result["Comparision_result"][0]+=1
					self.result["Comparision_result"].append(
						u"第"+str(i+5)+u"行，标准宗地编号为 "+Stander_ID+u" 估价师为 "+Appraiser+u" 的市场比较法楼面价或地面价为0"
					)
				else:
					comparision_ratio=float(Comparison_land)/float(Comparision_floor)
					diff=comparision_ratio-stander_ratio
					if diff>0.01 or diff<-0.01:
						self.result["Comparision_result"][0]+=1
						self.result["Comparision_result"].append(
							u"第"+str(i+5)+u"行，标准宗地编号为 "+Stander_ID+u" 估价师为 "+Appraiser+u" 的市场比较法地面价与楼面价比值为"+str(round(comparision_ratio,2))+u"，不等于标准容积率"+str(stander_ratio)
						)
		
		#Chreat thread pool and run
		obj_list=[]
		for i in range(0,self.lenrow):
			Stander_ID="<b>"+self.data["Stander_ID"][i]+"</b>"
			Appraiser="<b>"+self.data["Appraiser"][i]+"</b>"
			Appraise_way=[self.data["Appraise_way1"][i],self.data["Appraise_way1"][i]]
			task1=self.pool.submit(result_difference,i,Stander_ID,Appraiser)
			obj_list.append(task1)
			task2=self.pool.submit(market_comparision,i,Stander_ID,Appraiser,Appraise_way)
			obj_list.append(task2)
		return obj_list

	def result_stander(self):
		def stander_difference(stander):
			continue_tag=0
			data_fliter=self.data.loc[self.data["Stander_ID"]==stander]
			indexs=data_fliter.index.values
			Appraiser=data_fliter["Appraiser"]
			Result_lands=data_fliter["Result_land"]
			Result_floors=data_fliter["Result_floor"]
			for i in indexs:
				if Result_lands[i] in empty or Result_floors[i] in empty:
					self.result["Stander_ID"][0]+=1
					self.result["Stander_ID"].append(
						u"第"+str(i+5)+u"行，标准宗地编号为 <b>"+stander+u"</b> 估价师为 <b>"+Appraiser[i]+u"</b> 的评估结果为0"
					)
					continue_tag=1
			if continue_tag:
				return 0
			else:
				index1=indexs[0]
				index2=indexs[1]
				difference1=cal_difference(Result_lands[index1],Result_lands[index2])
				difference2=cal_difference(Result_floors[index1],Result_floors[index2])
				row1=str(index1+5)
				row2=str(index2+5)
				if difference1==0:
					self.result["Stander_ID"][0]+=1
					self.result["Stander_ID"].append(
						u"第"+row1+u"行和第"+row2+u"行，标准宗地编号为 <b>"+stander+u"</b> 估价师 <b>"+Appraiser[index1]+u"</b> 和估价师 <b>"+Appraiser[index2]+u"</b> 的地面价无差值"
					)
				elif difference2==0:
					self.result["Stander_ID"][0]+=1
					self.result["Stander_ID"].append(
						u"第"+row1+u"行和第"+row2+u"行，标准宗地编号为 <b>"+stander+u"</b> 估价师 <b>"+Appraiser[index1]+u"</b> 和估价师 <b>"+Appraiser[index2]+u"</b> 的楼面价无差值"
					)
				elif difference1>0.2:
					self.result["Stander_ID"][0]+=1
					self.result["Stander_ID"].append(
						u"第"+row1+u"行和第"+row2+u"行，标准宗地编号为 <b>"+stander+u"</b> 估价师 <b>"+Appraiser[index1]+u"</b> 和估价师 <b>"+Appraiser[index2]+u"</b> 的地面价差值为"+str(round(difference1*100,2))+u"%，大于20%"
					)
				elif difference2>0.2:
					self.result["Stander_ID"][0]+=1
					self.result["Stander_ID"].append(
						u"第"+row1+u"行和第"+row2+u"行，标准宗地编号为 <b>"+stander+u"</b> 估价师 <b>"+Appraiser[index1]+u"</b> 和估价师 <b>"+Appraiser[index2]+u"</b> 的楼面价差值为"+str(round(difference2*100,2))+u"%，大于20%"
					)
		
		#Chreat thread pool and run
		obj_list=[]
		for stander in self.standers:
			task1=self.pool.submit(stander_difference,stander)
			obj_list.append(task1)
		return obj_list

#check transaction table
class verify_trans(object):
	def __init__(self,path,year,quarter):
		data_origin=pd.read_excel(
			path,
			sheet_name=hot_update.get("Import","trans_name"),
			usecols=hot_update.get("Import","trans_data"),
			header=int(hot_update.get("Import","trans_header"))
		)
		self.data=data_origin.fillna("")
		self.lenrow=len(self.data)
		self.pool=ThreadPoolExecutor()
		self.year=year
		self.quarter=quarter
		self.plot_ratio={
			u"城镇住宅用地":config.get("stander_plot_ratio","resident"),
			u"城镇社区服务设施用地":config.get("stander_plot_ratio","public_comm"),
			u"机关团体用地":config.get("stander_plot_ratio","public_gov"),
			u"科研用地":config.get("stander_plot_ratio","public_resarch"),
			u"文化用地":config.get("stander_plot_ratio","public_cul"),
			u"教育用地":config.get("stander_plot_ratio","public_edu"),
			u"体育用地":config.get("stander_plot_ratio","public_sport"),
			u"医疗卫生用地":config.get("stander_plot_ratio","public_health"),
			u"社会福利用地":config.get("stander_plot_ratio","public_welfare"),
			u"商业用地":config.get("stander_plot_ratio","commerce"),
			u"娱乐康体用地":config.get("stander_plot_ratio","public_entertain"),
			u"商务金融用地":config.get("stander_plot_ratio","commerce_office"),
			u"其他商业服务业用地":config.get("stander_plot_ratio","commerce_other"),
			u"工业用地":config.get("stander_plot_ratio","industry"),
			u"物流仓储用地":config.get("stander_plot_ratio","storage"),
			u"其它":config.get("stander_plot_ratio","other"),
		}
		self.result={
			"diff_land_floor":[0,"",u"<b>·修正后地面价楼面价问题</b>"],
			"plot_ratio":[0,"",u"<b>·容积率或宗地面积、规划建筑面积问题</b>"],
			"sell_date":[0,"",u"<b>·成交日期不在本季度</b>"],
		}

	def result_trans(self):
		#Check the difference of appraising result
		def diff_land_floor(i,NO,Add):
			for j in range(0,3):
				land=self.data["Land_price"+str(j+1)][i]
				floor=self.data["Floor_price"+str(j+1)][i]
				Main_class=self.data["Main_class"+str(j+1)][i]
				if land in empty and floor in empty:
					continue
				elif (not land in empty) and floor in empty:
					self.result["diff_land_floor"][0]+=1
					self.result["diff_land_floor"].append(
						u"第"+str(i+3)+u"行，交易点编号为 <b>"+NO+u"</b> 宗地坐落 <b>"+Add+u"</b> ，<b>主要用途"+str(j+1)+u"</b> 楼面价为0"
					)
					continue
				elif land in empty and (not floor in empty):
					self.result["diff_land_floor"][0]+=1
					self.result["diff_land_floor"].append(
						u"第"+str(i+3)+u"行，交易点编号为 <b>"+NO+u"</b> 宗地坐落 <b>"+Add+u"</b> ，<b>主要用途"+str(j+1)+u"</b> 地面价为0"
					)
					continue
				stand_ratio=float(self.plot_ratio.get(Main_class,1))
				cal_ratio=float(land)/float(floor)
				diff=cal_ratio-stand_ratio
				if diff>0.01 or diff<-0.01:
					self.result["diff_land_floor"][0]+=1
					self.result["diff_land_floor"].append(
						u"第"+str(i+3)+u"行，交易点编号为 <b>"+NO+u"</b> 宗地坐落 <b>"+Add+u"</b> ，<b>主要用途"+str(j+1)+u"</b> 修正后的地面价与楼面价比为"+str(round(cal_ratio,2))+u"，不等于标准容积率"+str(stand_ratio)
					)
		
		#Check the plot raio
		def check_plot_ratio(i,NO,Add):
			Area=self.data["Area"][i]
			Buliding_area=self.data["Buliding_area"][i]
			Plot_ratio=self.data["Plot_ratio"][i]
			if not Area in empty and not Buliding_area in empty:
				ratio=Buliding_area/Area
				diff=Plot_ratio-ratio
				if diff>0.01 or diff<-0.01:
					self.result["plot_ratio"][0]+=1
					self.result["plot_ratio"].append(
						u"第"+str(i+3)+u"行，交易点编号为 <b>"+NO+u"</b> 宗地坐落 <b>"+Add+u"</b> 的规划建筑面积与宗地面积比为"+str(round(ratio,2))+u"，不等于容积率"+str(Plot_ratio)
					)

		#check sell date
		def check_sell_date(i,NO,Add):
			Sell_date=self.data["Sell_date"][i]
			year_trans=self.data["Sell_date"][i].year
			month_trans=self.data["Sell_date"][i].month
			range={"1":[12,1,2],"2":[3,4,5],"3":[6,7,8],"4":[9,10,11]}
			def show(i,NO,Add):
				self.result["sell_date"][0]+=1
				self.result["sell_date"].append(
					u"第"+str(i+3)+u"行，交易点编号为 <b>"+NO+u"</b> 宗地坐落 <b>"+Add+u"</b> 的成交时间不在本季度"
				)
			if not month_trans in range.get(self.quarter):
				show(i,NO,Add)
			elif self.quarter=="1" and ((month_trans in [2,3] and year_trans!=int(self.year)) or (month_trans==12 and year_trans!=int(self.year)-1)):
				show(i,NO,Add)
			elif self.quarter!="1" and year_trans!=int(self.year):
				show(i,NO,Add)

		#Chreat thread pool and run
		obj_list=[]
		for i in range(0,self.lenrow):
			NO=self.data["NO"][i]
			Add=self.data["Add"][i]
			task1=self.pool.submit(diff_land_floor,i,NO,Add)
			obj_list.append(task1)
			task2=self.pool.submit(check_plot_ratio,i,NO,Add)
			obj_list.append(task2)
			task3=self.pool.submit(check_sell_date,i,NO,Add)
			obj_list.append(task3)
		return obj_list