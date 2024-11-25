from PyQt5.QtWidgets import QMessageBox

def show_error_message(statue,text,info=""):
    msg=QMessageBox()
    msg.setStyleSheet("QMessageBox{font-size:16px}")
    msg.setText(text)
    msg.setInformativeText(info)
    msg.setModal(False)
    if statue==1:
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle(u"错误")
    elif statue==2:
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle(u"警告")
    elif statue==3:
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle(u"问题")
    else:
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle(u"信息")
    msg.exec_()
    return 0