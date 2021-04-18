import sys
import pyodbc
from PyQt5.QtSql import (QSqlTableModel)
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QPushButton,
    QSizePolicy, QLabel, QFontDialog, QApplication,QFileDialog,QDialog,QTableView)
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtSql import *
from setting import server,database_name,qr,query,query_del,query_edit

class Ui_Form(QWidget):

    db = QSqlDatabase.addDatabase('QODBC')
    db.setDatabaseName('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database_name + ';Trusted_Connection=yes')
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database_name + ';Trusted_Connection=yes')
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(620, 425)
        self.model = QSqlTableModel(Form)
        self.model.setQuery(QSqlQuery(qr))
        self.treeView = QtWidgets.QTreeView(Form)
        self.treeView.setGeometry(QtCore.QRect(10, 10, 601, 311))
        self.treeView.setObjectName("treeView")
        self.treeView.setModel(self.model)
        for i in range(self.model.columnCount()):
            self.treeView.resizeColumnToContents(i)
        self.treeView.setEditTriggers(QtWidgets.QTreeView.NoEditTriggers)
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(40, 350, 120, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(250, 350, 120, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(Form)
        self.pushButton_3.setGeometry(QtCore.QRect(460, 350, 120, 50))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.retranslateUi(Form)
        self.pushButton.clicked.connect(self.add)
        self.pushButton_2.clicked.connect(self.edit)
        self.pushButton_3.clicked.connect(self.delete)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Редактирование базы"))
        self.pushButton.setText(_translate("Form", "Добавить"))
        self.pushButton_2.setText(_translate("Form", "Изменить"))
        self.pushButton_3.setText(_translate("Form", "Удалить"))

    def add(self):
        _translate = QtCore.QCoreApplication.translate
        self.add_window=QWidget()
        self.add_window.setFixedSize(600, 300)
        self.add_window.setWindowTitle('Добавить компанию')
        self.pushButton_4 = QtWidgets.QPushButton(self.add_window)
        self.pushButton_4.setGeometry(QtCore.QRect(240, 230, 120, 50))
        self.pushButton_4.clicked.connect(self.add_commit)
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setText(_translate("Form", "Добавить"))
        self.model_add = QSqlTableModel(self.add_window)
        self.model_add.setTable('reestr')
        self.model_add.setEditStrategy(QSqlTableModel.OnFieldChange)
        self.model_add.select()
        self.tableView = QtWidgets.QTableView(self.add_window)
        self.tableView.setObjectName("tableView")
        self.tableView.setGeometry(QtCore.QRect(10, 10, 581, 201))
        self.model_add.insertRows(self.model_add.rowCount(), 1)
        self.tableView.setModel(self.model_add)
        self.add_window.show()

    def add_commit(self):
        self.cursor = self.conn.cursor()
        columns = self.model_add.columnCount()
        values = []
        for i in range(columns):
            values.append(self.model_add.index(0, i).data())
        self.cursor.execute(query, tuple(values))
        self.conn.commit()
        self.model.setQuery(QSqlQuery(self.qr))
        self.add_window.close()

    def edit(self):
        _translate = QtCore.QCoreApplication.translate
        self.edit_window = QWidget()
        self.row = self.treeView.currentIndex().row()
        self.edit_window.setFixedSize(600, 300)
        self.edit_window.setWindowTitle('Изменить компанию')
        self.pushButton_5 = QtWidgets.QPushButton(self.edit_window)
        self.pushButton_5.setGeometry(QtCore.QRect(240, 230, 120, 50))
        self.pushButton_5.clicked.connect(self.edit_commit)
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setText(_translate("Form", "Изменить"))
        self.model_edit = QSqlTableModel(self.edit_window)
        query_edit=('SELECT TOP ('+str(self.row+1)+') * FROM [dbo].[reestr]'
                                      ' EXCEPT'
                                      ' SELECT TOP ('+str(self.row)+') * FROM [dbo].[reestr]')
        self.model_edit.setQuery(QSqlQuery(query_edit))
        self.tableView = QtWidgets.QTableView(self.edit_window)
        self.tableView.setObjectName("tableView")
        self.tableView.setGeometry(QtCore.QRect(10, 10, 581, 201))
        self.tableView.setModel(self.model_edit)
        self.tableView.resizeColumnsToContents()
        self.edit_window.show()

    def edit_commit(self):
        self.cursor = self.conn.cursor()
        columns = self.model_edit.columnCount()
        values_edit = []
        for i in range(columns):
            values_edit.append(self.model_edit.index(0, i).data())
        for i in range(columns):
            values_edit.append(self.model.index(self.row, i).data())
        self.cursor.execute(query_edit, tuple(values_edit))
        self.conn.commit()
        self.model.setQuery(QSqlQuery(self.qr))
        self.edit_window.close()


    def delete(self):
        self.cursor = self.conn.cursor()
        self.data = [x.data() for x in self.treeView.selectedIndexes()]
        self.cursor.execute(query_del, tuple(self.data))
        self.conn.commit()
        self.model.setQuery(QSqlQuery(self.qr))
if __name__ in "__main__":
    class mywindow(QtWidgets.QMainWindow):
        def __init__(self):
            super(mywindow, self).__init__()
            self.ui = Ui_Form()
            self.ui.setupUi(self)


    app = QtWidgets.QApplication([])
    application = mywindow()
    application.show()
    sys.exit(app.exec())