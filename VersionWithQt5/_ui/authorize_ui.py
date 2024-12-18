# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'authorize_ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.9
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_authorize(object):
    def setupUi(self, authorize):
        authorize.setObjectName("authorize")
        authorize.resize(640, 95)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        authorize.setFont(font)
        self.verticalLayout = QtWidgets.QVBoxLayout(authorize)
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.key_1 = QtWidgets.QLineEdit(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        self.key_1.setFont(font)
        self.key_1.setObjectName("key_1")
        self.horizontalLayout.addWidget(self.key_1)
        self.label = QtWidgets.QLabel(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(22)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.key_2 = QtWidgets.QLineEdit(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        self.key_2.setFont(font)
        self.key_2.setObjectName("key_2")
        self.horizontalLayout.addWidget(self.key_2)
        self.label_2 = QtWidgets.QLabel(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(22)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout.addWidget(self.label_2)
        self.key_3 = QtWidgets.QLineEdit(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        self.key_3.setFont(font)
        self.key_3.setObjectName("key_3")
        self.horizontalLayout.addWidget(self.key_3)
        self.label_3 = QtWidgets.QLabel(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(22)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout.addWidget(self.label_3)
        self.key_4 = QtWidgets.QLineEdit(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        self.key_4.setFont(font)
        self.key_4.setObjectName("key_4")
        self.horizontalLayout.addWidget(self.key_4)
        self.label_4 = QtWidgets.QLabel(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(22)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout.addWidget(self.label_4)
        self.key_5 = QtWidgets.QLineEdit(authorize)
        font = QtGui.QFont()
        font.setFamily("微软雅黑")
        font.setPointSize(12)
        self.key_5.setFont(font)
        self.key_5.setObjectName("key_5")
        self.horizontalLayout.addWidget(self.key_5)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem)
        self.confirm_button = QtWidgets.QPushButton(authorize)
        self.confirm_button.setMaximumSize(QtCore.QSize(100, 16777215))
        self.confirm_button.setObjectName("confirm_button")
        self.horizontalLayout_2.addWidget(self.confirm_button)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Maximum, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem1)
        self.close_button = QtWidgets.QPushButton(authorize)
        self.close_button.setMaximumSize(QtCore.QSize(100, 16777215))
        self.close_button.setObjectName("close_button")
        self.horizontalLayout_2.addWidget(self.close_button)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem2)
        self.verticalLayout.addLayout(self.horizontalLayout_2)

        self.retranslateUi(authorize)
        QtCore.QMetaObject.connectSlotsByName(authorize)

    def retranslateUi(self, authorize):
        _translate = QtCore.QCoreApplication.translate
        authorize.setWindowTitle(_translate("authorize", "修改授权码"))
        self.label.setText(_translate("authorize", "-"))
        self.label_2.setText(_translate("authorize", "-"))
        self.label_3.setText(_translate("authorize", "-"))
        self.label_4.setText(_translate("authorize", "-"))
        self.confirm_button.setText(_translate("authorize", "确定"))
        self.close_button.setText(_translate("authorize", "关闭"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    authorize = QtWidgets.QWidget()
    ui = Ui_authorize()
    ui.setupUi(authorize)
    authorize.show()
    sys.exit(app.exec_())
