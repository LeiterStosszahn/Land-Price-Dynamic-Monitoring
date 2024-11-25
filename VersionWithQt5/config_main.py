import configparser
from PyQt5.QtWidgets import QWidget,QAbstractItemView
from PyQt5.QtCore import Qt,pyqtSignal
from PyQt5.QtGui import QIcon

from _ui.config_ui import Ui_config
from _ui.messagebox import show_error_message

def check_config(config_new,config_now,config_save_path):
	sections_new=config_new.sections()
	sections_now=config_now.sections()
	changed=0
	for section in sections_new:
		keys_news=config_new.options(section)
		if not section in sections_now:
			changed=1
			config_now.add_section(section)
			for key in keys_news:
				config_now.set(section,key,config_new.get(section,key))
		else:
			keys_nows=config_now.options(section)
			for key in keys_news:
				if not key in keys_nows:
					changed=1
					config_now.set(section,key,config_new.get(section,key))
	if changed:
		with open(config_save_path,"w",encoding="utf-8") as configfile:
			config_now.write(configfile)
	return changed


class config_ui(QWidget,Ui_config):
	sign_reload=pyqtSignal()

	def __init__(self,icon,config,hot_update):
		super(config_ui,self).__init__()
		self.config=config
		self.hot_update=hot_update
		self.setWindowIcon(icon)
		self.ratio_self=[]
		self.ratio_line=[]
		self.ratio_text=["resident",
			  "public_comm","public_gov","public_resarch","public_cul","public_edu",
			  "public_sport","public_health","public_welfare","public_entertain",
			  "commerce","commerce_office","commerce_other",
			  "industry","storage","other"]
		for i in range(0,16):
			self.ratio_self.append(self.config.get("stander_plot_ratio",self.ratio_text[i]))
		self.org_text=self.config.get("Default","organization")
		self.person_text=self.config.get("Default","filling_person")
		self.city_text=eval(self.config.get("GUI","city"))
		self.generate_type=self.config.get("generate_type","type")
		self.login_url=self.hot_update.get("upload_url","login_url")
		self.setupUi(self)
		self.initUI()
		

	#initialize city table
	def init_city_table(self):
		self.listWidget_city.clear()
		self.listWidget_city.addItems(self.city_text)
		self.get_all_city()

	#get all city
	def get_all_city(self):
		widgetres=[]
		count=self.listWidget_city.count()
		for i in range(count):
			item=self.listWidget_city.item(i)
			widgetres.append(item.text())
			item.setFlags(item.flags() | Qt.ItemIsEditable)
		return widgetres
	
	#add row
	def add_rows(self):
		cities_now=self.get_all_city()
		cities=self.lineEdit_cities.text().split(",")
		for i in range(0,len(cities)):
			city=cities[i]
			if city not in cities_now:
				self.listWidget_city.insertItem(0+i,city)
		self.lineEdit_cities.clear()

	#delete row
	def delete_rows(self):
		selected_rows=self.listWidget_city.currentRow()
		self.listWidget_city.takeItem(selected_rows)

	#set default city
	def set_default(self):
		selected_rows=self.listWidget_city.currentRow()
		selected_item=self.listWidget_city.currentItem().text()
		self.listWidget_city.takeItem(selected_rows)
		self.listWidget_city.insertItem(0,selected_item)
		

	def initUI(self):
		self.ratio_line=[self.lineEdit_resident,
			  self.lineEdit_public_comm,self.lineEdit_public_gov,self.lineEdit_public_resarch,self.lineEdit_public_cul,self.lineEdit_public_edu,
			  self.lineEdit_public_sport,self.lineEdit_public_health,self.lineEdit_public_welfare,self.lineEdit_public_entertain,
			  self.lineEdit_commerce,self.lineEdit_commerce_office,self.lineEdit_commerce_other,
			  self.lineEdit_industry,self.lineEdit_storage,self.lineEdit_other]
		for i in range(0,16):
			self.ratio_line[i].setText(self.ratio_self[i])
		self.lineEdit_org.setText(self.org_text)
		self.lineEdit_person.setText(self.person_text)
		self.init_city_table()
		self.listWidget_city.setDragDropMode(QAbstractItemView.InternalMove)
		self.toolButton_add.clicked.connect(self.add_rows)
		self.toolButton_delete.clicked.connect(self.delete_rows)
		self.toolButton_reset.clicked.connect(self.init_city_table)
		self.toolButton_default.clicked.connect(self.set_default)
		if self.generate_type=="正常填报":
			self.radioButton_normal.click()
		elif self.generate_type=="衔接填报":
			self.radioButton_join.click()
		self.lineEdit_culp.setText(self.login_url)

		self.pushButton_confirm.clicked.connect(self.update)

	def update(self):
		isChanged=0
		for i in range(0,16):
			if self.ratio_self[i]!=self.ratio_line[i].text():
				self.ratio_self[i]=self.ratio_line[i].text()
				self.config.set("stander_plot_ratio",self.ratio_text[i],self.ratio_self[i])
				isChanged=1

		if self.org_text!=self.lineEdit_org.text():
			self.org_text=self.lineEdit_org.text()
			self.config.set("Default","organization",self.org_text)
			isChanged=1

		if self.person_text!=self.lineEdit_person.text():
			self.person_text=self.lineEdit_person.text()
			self.config.set("Default","filling_person",self.person_text)
			isChanged=1

		city_text_now=self.get_all_city()
		if self.city_text!=city_text_now:
			self.city_text=city_text_now
			self.config.set("GUI","city",str(self.city_text))
			isChanged=1

		if self.radioButton_normal.isChecked() and self.generate_type!="正常填报":
			self.generate_type="正常填报"
			self.config.set("generate_type","type",self.generate_type)
			isChanged=1
		elif self.radioButton_join.isChecked() and self.generate_type!="衔接填报":
			self.generate_type="衔接填报"
			self.config.set("generate_type","type",self.generate_type)
			isChanged=1

		if self.login_url!=self.lineEdit_culp.text():
			self.login_url=self.lineEdit_culp.text()
			self.hot_update.set("upload_url","login_url",self.login_url)
			isChanged=2

		if isChanged:
			try:
				file=open(r"config.ini",'w',encoding="utf-8")
				self.config.write(file)
				if isChanged==2:
					file=open(r"bin\\hot_update.ini",'w',encoding="utf-8")
					self.hot_update.write(file)
			except:
				show_error_message(1,"保存失败")
			finally:
				show_error_message(4,"保存成功")
				self.sign_reload.emit()
				self.close()
		else:
			show_error_message(4,"未修改配置文件")
			self.close()