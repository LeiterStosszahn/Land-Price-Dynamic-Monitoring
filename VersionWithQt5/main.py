import os,sys,subprocess,time,configparser,ctypes
import pandas as pd
from PyQt5.QtWidgets import QApplication,QMainWindow,QFileDialog,QMessageBox,QStatusBar,QTableWidgetItem,QCheckBox
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtGui import QIcon
from concurrent.futures import ThreadPoolExecutor,as_completed

from config_main import check_config

from _ui.main_ui import *
from _ui.about_ui import *
from _ui.messagebox import show_error_message
from authorize_main import authorize_ui,check_authorize
from config_main import config_ui

import gevent

myappid="Land Price Dynamic Monitoring Tool"
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

#Config info
time_now=time.localtime()
hot_update=configparser.ConfigParser()
hot_update.read(r"bin\\hot_update.ini",encoding="utf-8")
config=configparser.ConfigParser()
config.read(r"config.ini",encoding="utf-8")
newest_config=configparser.ConfigParser()
newest_config.read(r"bin\\newest_config.ini",encoding="utf-8")
config_changed=check_config(newest_config,config,r"config.ini")
if config_changed:
    config.read(r"config.ini",encoding="utf-8")

#Set path
def get_path(style="file"):
    if style=="directory":
        path=QFileDialog.getExistingDirectory(None,u"选择目录","",QFileDialog.ShowDirsOnly)
    elif style=="file":
        path,_=QFileDialog.getOpenFileName(None,u"请选择一个文件","","Excel文件(*.xls *.xlsx)")
    return path

#Bound signal of main UI
class main_ui(Ui_MainWindow):
    def __init__(self,MainWindow):
        super().__init__()
        self.setupUi(MainWindow)
        icon=QIcon(r"bin\\logo.ico")
        MainWindow.setWindowIcon(icon)
        self.init_authorize=authorize_ui(icon,hot_update)
        self.init_config=config_ui(icon,config,hot_update)
        self.text_UI()
        self.initUI()
        self.status=MainWindow.statusBar()
        self.status.showMessage("233",2000)
        self.QTabWidget.setCurrentIndex(0)
        
        key=hot_update.get("authorize","key")
        if check_authorize(key,hot_update):
            show_error_message(1,"授权码已过期")
            self.init_authorize.show()
        if config_changed:
            show_error_message(4,u"配置文件已更新")

    #QLineEdit set path
    def Q_path(self,result,style="file"):
        path=get_path(style)
        result.setText(path)

    #prgressbar update
    def progressBar_update(self,start,end,bar,sleep_time):
        for i in range(start,end):
            bar.setValue(i)
            gevent.sleep(sleep_time)

    #Generate information
    def gen_info(self,result):
        if result[0]:
            self.status.showMessage(result[1],2000)
        else:
            show_error_message(1,u"汇总表有误",result[1])
        gevent.sleep(0)

    #save check result
    def save_check_result(self):
        path=get_path("directory")
        text=self.textBrowser_checkresult.toHtml()
        with open(os.path.join(path,"检查结果.html"),"w",encoding="utf-8") as f:
            f.write(text)
            show_error_message(4,u"保存成功")

    #initialize appraiser table
    def init_appraiser_table(self):
        appraiser_accounts=config.options("Appraiser_account")
        appraiser_amount=len(appraiser_accounts)
        self.appraiser_table.setRowCount(appraiser_amount)
        for i in range(0,appraiser_amount):
            content=eval(config.get("Appraiser_account",str(i+1)))
            for j in range(0,3):
                self.appraiser_table.setItem(i,j,QTableWidgetItem(content[j]))
            checkbox=QCheckBox()
            checkbox.setChecked(content[3])
            self.appraiser_table.setCellWidget(i,3,checkbox)

    #add row
    def add_rows(self):
        amount_now=self.appraiser_table.rowCount()
        self.appraiser_table.insertRow(amount_now)
        checkbox=QCheckBox()
        checkbox.setChecked(1)
        self.appraiser_table.setCellWidget(amount_now,3,checkbox)

    #delete row
    def delete_rows(self):
        selected_rows=self.appraiser_table.selectedItems()
        rows=set()
        for item in selected_rows:
            rows.add(item.row())
        rows=sorted(rows, reverse=True)
        for row in rows:
            self.appraiser_table.removeRow(row)

    #save appraiser table
    def save_appraiser_table(self):
        config.set("Account","organization",self.organization_account.text())
        config.set("Account","password",self.organization_passw.text())
        amount_now=self.appraiser_table.rowCount()
        appraiser_accounts=config.options("Appraiser_account")
        appraiser_amount=len(appraiser_accounts)
        for i in range(0,amount_now):
            contents=[]
            for j in range(0,3):
                try:
                    content=self.appraiser_table.item(i,j).text()
                except:
                    show_error_message(1,u"估价师表格不能为空")
                    return 0
                contents.append(content)
            content=self.appraiser_table.cellWidget(i,3).isChecked()
            contents.append(content)
            config.set("Appraiser_account",str(i+1),str(contents))
        if amount_now<appraiser_amount:
            for i in range(amount_now,appraiser_amount):
                config.remove_option("Appraiser_account",str(i+1))
        try:
            file=open(r"config.ini",'w',encoding="utf-8")
            config.write(file)
        except:
            show_error_message(1,u"保存失败")
        finally:
            show_error_message(4,u"保存成功")
    
    #AboutWindow
    def about_window(self):
        self.about=QtWidgets.QWidget()
        self.ui_about=Ui_about()
        self.ui_about.setupUi(self.about)
        self.about.show()

    #Trigger
    def text_UI(self):
        #transform_summary
        self.org_text.setText(self.init_config.org_text)
        self.person_text.setText(self.init_config.person_text)
        self.city_text.clear()
        self.city_text.addItems(self.init_config.city_text)
        self.year_text.setText(time.strftime("%Y",time_now))

        #upload_result
        self.city_text_upload.clear()
        self.city_text_upload.addItems(self.init_config.city_text)

    def initUI(self):
        #QMenu/File
        def save_sample(type):
            path=get_path("directory")
            try:
                os.system("xcopy "+type+" "+path.replace('/','\\'))
            except:
                show_error_message(1,u"保存失败")
            finally:
                show_error_message(4,u"保存成功")
        self.sum_template.triggered.connect(lambda:save_sample(u"sample\\汇总表-样表.xlsx"))
        self.cal_template.triggered.connect(lambda:save_sample(u"sample\\评估方法-样表.xlsx"))

        #QMenu/Online
        start="start "+self.init_config.login_url
        self.national_sys.triggered.connect(lambda:subprocess.run(start,shell=True))

        #QMenu/Config
        self.config.triggered.connect(self.init_config.show)
        self.change_config.triggered.connect(lambda:subprocess.run("config.ini",shell=True))
        self.change_advconfig.triggered.connect(lambda:subprocess.run("bin\\hot_update.ini",shell=True))

        #QMenu/About
        self.about.triggered.connect(self.about_window)
        self.authorize_key.triggered.connect(self.init_authorize.show)
        self.exit.triggered.connect(sys.exit)

        #check result
        self.button_checkpath.clicked.connect(lambda:self.Q_path(self.checkpath_text))
        self.toolButton_clearcheck.clicked.connect(self.textBrowser_checkresult.clear)
        self.toolButton_savecheck.clicked.connect(self.save_check_result)
        self.pushButton_startcheck.clicked.connect(self.check_and_check)

        #transform_summary
        self.quarter_text.setCurrentText(str((int(time.strftime("%m",time_now))-1)//3+1))
        self.date_text.setText(time.strftime("%Y/%m/%d",time_now))
        self.button_path.clicked.connect(lambda:self.Q_path(self.path_text))
        self.button_savepath.clicked.connect(lambda:self.Q_path(self.savepath_text,"directory"))
        self.button_generate.clicked.connect(self.check_and_run)

        #upload_result
        self.init_appraiser_table()
        self.organization_account.setText(config.get("Account","organization"))
        self.organization_passw.setText(config.get("Account","password"))
        self.toolButton_add.clicked.connect(self.add_rows)
        self.toolButton_delete.clicked.connect(self.delete_rows)
        self.toolButton_reset.clicked.connect(self.init_appraiser_table)
        self.toolButton_save.clicked.connect(self.save_appraiser_table)
        self.button_gen_path.clicked.connect(lambda:self.Q_path(self.gen_path_text,"directory"))
        self.button_upload.clicked.connect(self.check_and_upload)

    #check filling box and run function to check result sheet
    def check_and_check(self,direct_check=None):
        #progressBar statue
        self.progressBar_check.setValue(1)

        #text statue
        checkpath_text=self.checkpath_text.text()
        self.progressBar_check.setValue(2)

        checkstatue=1
        if checkpath_text=="" and not direct_check:
            show_error_message(1,u"请选择汇总表")
            checkstatue=0
        self.progressBar_check.setValue(3)

        #All check pass
        if checkstatue or direct_check:
            self.textBrowser_checkresult.clear()
            import check_function

            #verify the integrality
            report_check_class=check_function.verity_integrality(checkpath_text,"report")
            report_integrality=report_check_class.verify()
            form_and_rent_check_class=check_function.verity_integrality(checkpath_text,"form_and_rent")
            form_and_rent_integrality=form_and_rent_check_class.verify()
            trans_check_class=check_function.verity_integrality(checkpath_text,"trans")
            trans_integrality=trans_check_class.verify()
            check_percent=4
            self.progressBar_check.setValue(check_percent)

            if not report_integrality[0]:
                self.textBrowser_checkresult.append(report_integrality[1])
                check_percent+=30
                report_tag=1
                self.progressBar_check.setValue(check_percent)
            if not form_and_rent_integrality[0]:
                self.textBrowser_checkresult.append(form_and_rent_integrality[1])
                check_percent+=30
                form_and_rent_tag=1
                self.progressBar_check.setValue(check_percent)
            if not trans_integrality[0]:
                self.textBrowser_checkresult.append(trans_integrality[1])
                check_percent+=30.
                trans_tag=1
                self.progressBar_check.setValue(check_percent)

            #No integrality passed    
            if check_percent==94:
                show_error_message(1,u"无报告信息表、表格信息表、交易样点表或其缺少关键字段，请选择正确文件")
                self.progressBar_check.setValue(0)
                return False
            elif direct_check and check_percent!=4:
                # direct_check=[checkBox_appraiser,checkBox_rent,checkBox_transication,checkBox_upload(,checkBox_report:deleted)]
                if report_tag and (direct_check[3] or direct_check[4]):#or direct_check[5]
                    show_error_message(1,u"无报告信息表或其缺少关键字段，请选择正确文件")
                    direct_tag=1
                if form_and_rent_tag and (direct_check[0] or direct_check[1]):
                    show_error_message(1,u"无表格信息表或其缺少关键字段，请选择正确文件")
                    direct_tag=1
                if trans_tag and direct_check[2]:
                    show_error_message(1,u"无交易样点表或其缺少关键字段，请选择正确文件")
                    direct_tag=1
                if direct_tag:
                    self.progressBar_check.setValue(0)
                    self.progressBar.setValue(0)
                    return False

            #check report sheet
            report_final=0#Initialize report final True
            if report_integrality[0]:
                #Initialize check
                check_step=15#check_step=30/num of check step
                class_verify_report=check_function.verify_report(checkpath_text)
                self.textBrowser_checkresult.append(u"<b>-报告信息表检查结果：</b>")
                #check weight
                result_weight=class_verify_report.check_weight()
                if result_weight[0]:
                    for i in range(1,result_weight[0]+1):
                        self.textBrowser_checkresult.append(result_weight[i])
                check_percent+=check_step
                self.progressBar_check.setValue(check_percent)
                
                #...add other check ...#

                report_final=result_weight[0]#+other result
                if not report_final:
                    self.textBrowser_checkresult.append(u"报告信息表检查无误")

            self.textBrowser_checkresult.append("")
            
            #check form sheet
            result_final=0#Initialize result final False
            if form_and_rent_integrality[0]:
                #Initialize check
                check_step=10#check_step=30/num of check step
                class_verify_form=check_function.verify_form(checkpath_text)

                #run function and get object list
                appraiser_obj_list=class_verify_form.result_appraiser()
                stander_obj_list=class_verify_form.result_stander()
                self.textBrowser_checkresult.append(u"<b>-表格信息表检查结果：</b>")
                
                #check appraiser result
                self.textBrowser_checkresult.append("")
                self.textBrowser_checkresult.append(u"<b>--估价师成果：</b>")
                #check content
                check_content=["difference","Sample_NO","Comparision_result",]
                check_percent,result_appaiser=check_function.show_check(
                    check_content,appraiser_obj_list,check_percent,self.progressBar_check,class_verify_form.result,self.textBrowser_checkresult,u"估价师成果"
                    )
                
                #check stander result
                self.textBrowser_checkresult.append("")
                self.textBrowser_checkresult.append(u"<b>--标准宗地成果：</b>")
                #check content
                stander_content=["Stander_ID",]
                check_percent,result_stander=check_function.show_check(
                            stander_content,stander_obj_list,check_percent,self.progressBar_check,class_verify_form.result,self.textBrowser_checkresult,u"标准宗地成果"
                            )
                
                #progress bar
                result_final=result_appaiser+result_stander
                check_percent+=check_step
                self.progressBar_check.setValue(check_percent)

            result_trans=0#Initialize result trans False
            if trans_integrality[0]:
                #Initialize check
                check_step=10#check_step=30/num of check step
                class_verify_trans=check_function.verify_trans(checkpath_text,self.year_text.text(),self.quarter_text.currentText())

                #run function and get object list
                trans_obj_list=class_verify_trans.result_trans()
                self.textBrowser_checkresult.append("")
                self.textBrowser_checkresult.append(u"<b>-交易样点表检查结果：</b>")
                #check content
                trans_content=["diff_land_floor","plot_ratio","sell_date",]
                check_percent,result_trans=check_function.show_check(
                            trans_content,trans_obj_list,check_percent,self.progressBar_check,class_verify_trans.result,self.textBrowser_checkresult,u"交易样点成果"
                            )

            self.progressBar_check.setValue(100)
            show_error_message(4,u"检查完成")
            
            if direct_check and report_final+result_final+result_trans:
                return False
            elif not report_final+result_final:
                self.path_text.setText(checkpath_text)
                return True

        else:
            self.progressBar_check.setValue(0)
            return False

    #Check filling box and run function to generate result
    def check_and_run(self):
        #progressBar statue
        self.progressBar.setValue(1)

        #QCheckBox statue
        checkBox_appraiser=self.checkBox_appraiser.isChecked()
        checkBox_rent=self.checkBox_rent.isChecked()
        checkBox_transication=self.checkBox_transication.isChecked()
        checkBox_check=self.checkBox_check.isChecked()
        checkBox_upload=self.checkBox_upload.isChecked()
        # checkBox_report=self.checkBox_report.isChecked()
        self.progressBar.setValue(2)

        #Text statue
        filling_org=self.org_text.text()
        person=self.person_text.text()
        city=self.city_text.currentText()
        year=self.year_text.text()
        quarter=self.quarter_text.currentText()
        date=self.date_text.text()
        data_path=self.path_text.text()
        save_path=self.savepath_text.text()
        self.progressBar.setValue(3)

        checkstatue=1
        if city=="":
            show_error_message(1,u"城市不能为空")
            checkstatue=0
        if year=="":
            show_error_message(1,u"年份不能为空")
            checkstatue=0
        if quarter=="":
            show_error_message(1,u"季度不能为空")
            checkstatue=0
        if data_path=="":
            show_error_message(1,u"未选择汇总表")
            checkstatue=0
        if save_path=="":
            show_error_message(1,u"未选择保存路径")
            checkstatue=0
        if date=="" and checkBox_transication:
            show_error_message(1,u"填表日期不能为空")
            checkstatue=0
        if (checkBox_transication) and filling_org=="": #or checkBox_report
            show_error_message(1,u"填表单位不能为空")
            checkstatue=0
        if (checkBox_transication) and person=="": #or checkBox_report
            show_error_message(1,u"填表人不能为空")
            checkstatue=0
        if not (checkBox_appraiser or checkBox_rent or 
            checkBox_transication or 
            checkBox_upload
            #or checkBox_report
            ):
            show_error_message(1,u"请至少勾选一个成果类型")
            checkstatue=0
        self.progressBar.setValue(4)

        #All check pass
        if checkstatue:
            #Change to check index
            if checkBox_check:
                self.checkpath_text.setText(data_path)
                self.QTabWidget.setCurrentIndex(0)#page1:check result
                check_result=self.check_and_check(direct_check=[checkBox_appraiser,checkBox_rent,checkBox_transication,checkBox_upload])#,checkBox_report
                if check_result:
                    self.QTabWidget.setCurrentIndex(1)#page2:transform summary
                else:
                    self.progressBar.setValue(0)
                    return 0
            
            import generate_function
            pool=ThreadPoolExecutor()
            obj_list=[]
            obj_count=0
            self.progressBar.setValue(5)

            if checkBox_appraiser:
                task_appraise=pool.submit(generate_function.transform_appraiser,filling_org,city,year,quarter,data_path,save_path)
                obj_list.append(task_appraise)
                obj_count+=1
            if checkBox_rent:
                task_rent=pool.submit(generate_function.transform_rent,filling_org,city,year,quarter,data_path,save_path)
                obj_list.append(task_rent)
                obj_count+=1
            if checkBox_transication:
                task_transication=pool.submit(generate_function.transform_sample,filling_org,person,city,year,quarter,date,data_path,save_path)
                obj_list.append(task_transication)
                obj_count+=1
            if checkBox_upload:
                task_upload=pool.submit(generate_function.generate_weight,city,year,quarter,data_path,save_path)
                obj_list.append(task_upload)
                obj_count+=1
            # if checkBox_report:
            #     pass
            #     obj_count+=1
            self.progressBar.setValue(6)
            
            #show result
            obj_step=int(90/obj_count)
            t0=gevent.spawn(self.progressBar_update,5,5+obj_step,self.progressBar,0.03*obj_count)
            gevent.joinall([t0])
            progress=5+obj_step
            for future in as_completed(obj_list):
                result=future.result()
                t1=gevent.spawn(self.progressBar_update,progress,progress+obj_step,self.progressBar,0.03*obj_count)
                progress+=obj_step
                t2=gevent.spawn(self.gen_info,result)
                gevent.joinall([t1, t2])
                
            self.progressBar.setValue(100)
            self.gen_path_text.setText(save_path)
            self.city_text_upload.setCurrentText(city)
            show_error_message(4,u"生成完成")

        else:
            self.progressBar.setValue(0)

    #Check filling box and run function to upload result
    def check_and_upload(self):
        #progressBar_upload statue
        self.progressBar_upload.setValue(1)

        #QCheckBox statue
        checkBox_upload_org=self.checkBox_upload_org.isChecked()
        checkBox_upload_appraiser=self.checkBox_upload_appraiser.isChecked()
        checkBox_upload_weight=self.checkBox_upload_weight.isChecked()
        self.progressBar_upload.setValue(2)

        #Text statue
        organization_account=self.organization_account.text()
        organization_passw=self.organization_passw.text()
        city_upload=self.city_text_upload.currentText()
        appraiser_table_1st=self.appraiser_table.item(0,0)
        gen_path_text=self.gen_path_text.text()
        self.progressBar_upload.setValue(3)

        checkstatue=1
        if checkBox_upload_org and (organization_account=="" or organization_passw==""):
            show_error_message(1,u"估价机构账号及密码不能为空")
            checkstatue=0
        if (checkBox_upload_appraiser or checkBox_upload_weight) and appraiser_table_1st is None:
            show_error_message(1,u"估价师至少填写一位")
            checkstatue=0
        if gen_path_text=="":
            show_error_message(1,u"请选择成果目录")
            checkstatue=0
        else:
            dir_main=os.listdir(gen_path_text)
            if checkBox_upload_org and u"技术承担单位" not in dir_main:
                show_error_message(1,u"目录内无“技术承担单位成果”文件夹")
                checkstatue=0
            if (checkBox_upload_appraiser or checkBox_upload_weight) and u"估价师成果" not in dir_main:
                show_error_message(1,u"目录内无“估价师成果”文件夹")
                checkstatue=0
            elif checkBox_upload_appraiser or checkBox_upload_weight:
                path_appraiser=os.path.join(gen_path_text,u"估价师成果")
                dir_appraiser=os.listdir(path_appraiser)
                appraiser_amount=self.appraiser_table.rowCount()
                for i in range(0,appraiser_amount):
                    dir_name=self.appraiser_table.item(i,0).text()+"("+self.appraiser_table.item(i,1).text()+")"
                    if dir_name not in dir_appraiser:
                        show_error_message(1,u"估价师成果无“"+dir_name+"”文件夹")
                        checkstatue=0
                    elif "权重统计表.xlsx" not in os.listdir(os.path.join(path_appraiser,dir_name)):
                        show_error_message(1,u"“"+dir_name+"”内无“权重统计表.xlsx”")
                        checkstatue=0

        if not (checkBox_upload_org or checkBox_upload_appraiser or checkBox_upload_weight):
            show_error_message(1,u"请至少勾选一个上报内容")
            checkstatue=0
        self.progressBar_upload.setValue(4)

        if checkstatue:
            import upload_result
            generate_type=self.init_config.generate_type
            # pool_upload=ThreadPoolExecutor()
            # upload_list=[]
            self.progressBar_upload.setValue(5)
            
            if checkBox_upload_org:
                if checkBox_upload_appraiser or checkBox_upload_weight:
                    t0=gevent.spawn(self.progressBar_update,5,50,self.progressBar_upload,2)
                else:
                    t0=gevent.spawn(self.progressBar_update,5,95,self.progressBar_upload,1)
                path=os.path.join(gen_path_text,u"技术承担单位")
                #t1=gevent.spawn(upload_result.估价机构,city_upload,organization_account,organization_passw,generate_type,path)
                gevent.joinall([t0])
                #gevent.joinall([t0,t1])
                #机构上传还没做
            
            if checkBox_upload_appraiser or checkBox_upload_weight:
                #Initialize progress bar
                amount_now=self.appraiser_table.rowCount()
                if checkBox_upload_org:
                    progress=50
                    obj_step=int(45/amount_now)
                    step=1
                else:
                    progress=5
                    obj_step=int(90/amount_now)
                    step=0.5

                for i in range(0,amount_now):
                    usrname=self.appraiser_table.item(i,0).text()
                    user=self.appraiser_table.item(i,1).text()
                    pwd=self.appraiser_table.item(i,2).text()
                    is_submit=self.appraiser_table.cellWidget(i,3).isChecked()
                    path=os.path.join(gen_path_text,u"估价师成果",usrname+u"("+user+u")")

                    #Update progress bar
                    t0=gevent.spawn(self.progressBar_update,progress,progress+obj_step,self.progressBar_upload,step)
                    progress+=obj_step
                    t1=gevent.spawn(self.gen_info,u"填报估价师"+usrname+u"中")

                    if checkBox_upload_appraiser:
                        pass
                    if checkBox_upload_weight:
                        t2=gevent.spawn(upload_result.main,city_upload,user,pwd,generate_type,path,is_submit)
                        gevent.joinall([t0,t1,t2])

                # #Multi task pool
                # for i in range(0,amount_now):
                #     usrname=self.appraiser_table.item(i,0).text()
                #     user=self.appraiser_table.item(i,1).text()
                #     pwd=self.appraiser_table.item(i,2).text()
                #     is_submit=self.appraiser_table.cellWidget(i,3).isChecked()

                #     if checkBox_upload_appraiser:
                #         pass
                #     if checkBox_upload_weight:
                #         upload_weight=pool_upload.submit(upload_result.main,city_upload,user,pwd,generate_type,is_submit)
                #         upload_list.append(upload_weight)
                        
                # for future in as_completed(upload_list):
                #     #Update progress bar
                #     result=future.result()
                #     t0=gevent.spawn(self.progressBar_update,progress,progress+obj_step,self.progressBar_upload,step)
                #     progress+=obj_step
                #     t1=gevent.spawn(self.gen_info,result)
                #     gevent.joinall([t0,t1])

            self.progressBar.setValue(100)
            show_error_message(4,u"上报完成")

        else:
            self.progressBar_upload.setValue(0)

#Main
if __name__ == '__main__':
    #Initialize UI
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = main_ui(MainWindow)
    MainWindow.show()
    ui.init_config.sign_reload.connect(ui.text_UI)
    sys.exit(app.exec_())