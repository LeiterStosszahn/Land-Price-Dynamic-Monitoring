import sys,configparser,math,time
from PyQt5.QtWidgets import QWidget
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

from _ui.authorize_ui import Ui_authorize
from _ui.messagebox import show_error_message
from _test import code_gen

def check_authorize(key,hot_update):
	timenow=time.time()
	pwd=hot_update.get("authorize","pwd")
	if pwd=="":
		return 1
	else:
		pwd_y,pwd_m,pwd_d,pwd_c=code_gen.decode_key(pwd)
		if timenow<time.mktime((pwd_y,pwd_m,pwd_d,0,0,0,0,0,0)) or pwd_c:
			return 1
		else:
			a=int(time.strftime('%Y',time.gmtime()))
			b=int(time.strftime('%m',time.gmtime()))
			c=int(time.strftime('%d',time.gmtime()))
			pwd=code_gen.encode_key(a,b,c)
			hot_update.set("authorize","pwd",pwd)
			try:
				file=open(r"bin\\hot_update.ini",'w',encoding="utf-8")
				hot_update.write(file)
			except:
				return 1
	year,month,day,hash_result=code_gen.decode_key(key)

	if timenow>time.mktime((year,month,day,0,0,0,0,0,0)) or hash_result:
		return 1
	else:
		return 0

class authorize_ui(QWidget,Ui_authorize):
	def __init__(self,icon,hot_update):
		self.close_key=0
		self.hot_update=hot_update
		self.key=self.hot_update.get("authorize","key")
		super(authorize_ui,self).__init__()
		self.setupUi(self)
		self.setWindowModality(Qt.ApplicationModal)
		self.setWindowFlags(Qt.Window | Qt.CustomizeWindowHint | Qt.WindowMinimizeButtonHint)
		self.setWindowIcon(icon)
		self.initUI()

	def closeEvent(self, event):
		if self.close_key:
			event.accept()
		else:
			event.ignore()  # Close event ignored
	
	def initUI(self):
		key_list=self.key.split("-")
		self.key_1.setText(key_list[0])
		self.key_2.setText(key_list[1])
		self.key_3.setText(key_list[2])
		self.key_4.setText(key_list[3])
		self.key_5.setText(key_list[4])
		self.confirm_button.clicked.connect(self.update)
		if check_authorize(self.key,self.hot_update):
			self.close_button.clicked.connect(sys.exit)
		else:
			self.close_key=1
			self.close_button.clicked.connect(self.close)
	
	def update(self):
		key_1=self.key_1.text()
		key_2=self.key_2.text()
		key_3=self.key_3.text()
		key_4=self.key_4.text()
		key_5=self.key_5.text()
		self.key=key_1+"-"+key_2+"-"+key_3+"-"+key_4+"-"+key_5
		if check_authorize(self.key,self.hot_update):
			show_error_message(1,"授权码无效")
		else:
			self.hot_update.set("authorize","key",self.key)
			try:
				file=open(r"bin\\hot_update.ini",'w',encoding="utf-8")
				self.hot_update.write(file)
			except:
				show_error_message(1,"保存失败")
			finally:
				show_error_message(4,"授权码更新成功")
				self.close_key=1
				self.close()