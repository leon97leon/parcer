import sys
from PyQt5.QtSql import (QSqlTableModel)
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QPushButton,
    QSizePolicy, QLabel, QFontDialog, QApplication,QFileDialog,QDialog,QTableView)
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtSql import *
import main
import xlsxwriter
from creat_sql import writex_sql
from EditSql import Ui_Form
class Ui_MainWindow(QWidget):
    BASE_URL = "http://reestr.nostroy.ru"
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 317)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(10, 10, 191, 21))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(500, 180, 250, 70))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(50, 180, 250, 70))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(14)
        font.setBold(True)
        font.setWeight(75)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setGeometry(QtCore.QRect(180, 10, 171, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(360, 10, 75, 20))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(280, 40, 251, 61))
        font = QtGui.QFont()
        font.setFamily("Times New Roman")
        font.setPointSize(16)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setObjectName("pushButton_4")
        self.pushButton_4.clicked.connect(self.parcer)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(20, 110, 771, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        MainWindow.setCentralWidget(self.centralwidget)
        self.pushButton_3.clicked.connect(self.showDialog)
        self.pushButton_2.clicked.connect(self.sql)
        self.retranslateUi(MainWindow)
        self.pushButton.clicked.connect(self.edit)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
    def edit(self):
        class SecondWindow(QMainWindow, Ui_Form):  # +++
            def __init__(self, parent=None):
                super(SecondWindow, self).__init__(parent)

                self.setupUi(self)

        self.w2 = SecondWindow()
        self.w2.show()
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Парсер реестра Нострой"))
        self.label.setText(_translate("MainWindow", "<html><head/><body><p><span style=\" font-size:12pt;\">Парсер реестра Нострой:</span></p></body></html>"))
        self.pushButton.setText(_translate("MainWindow", "Редактировать базу"))
        self.pushButton_2.setText(_translate("MainWindow", "Загрузить в базу"))
        self.pushButton_3.setText(_translate("MainWindow", "Обзор"))
        self.pushButton_4.setText(_translate("MainWindow", "Запустить парсер"))
    def sql(self):
        try:
            writex_sql().run()
            QMessageBox.information(self, "Информация", "Данные загружены в базу данных", QMessageBox.Ok)
        except:
            QMessageBox.critical(self, "Ошибка ", "Запустите и дождитесь завершения парсинга", QMessageBox.Ok)
    def showDialog(self):
        fname = QFileDialog.getOpenFileName(self,"Выбрать файл")[0]
        self.lineEdit.setText(fname)
    def create_file(self):
        workbook = xlsxwriter.Workbook('result.xlsx')
        worksheet = workbook.add_worksheet('Sheet1')
        format = workbook.add_format(
            {'align': 'center_across', 'font_name': 'Times New Roman', 'bold': True, 'font_size': '12',
             'text_wrap': True})
        worksheet.merge_range('K1:L1', 'Сведение о приостановлении/возобновлении права', format)
        worksheet.set_column('A1:P1', 15)
        worksheet.set_column('M2:P2', 17)
        worksheet.set_column('I1:J1', 20)
        worksheet.set_column('L1:L1', 40)
        worksheet.set_row(1, 60)
        worksheet.merge_range('M1:P1', 'Архив', format)
        worksheet.merge_range('A1:A2', '№ п/п', format)
        worksheet.merge_range('B1:B2', 'Сокращенное наименование члена СРО', format)
        worksheet.merge_range('C1:C2', 'ИНН', format)
        worksheet.merge_range('D1:D2', 'Статус', format)
        worksheet.merge_range('E1:E2', 'Тип', format)
        worksheet.merge_range('F1:F2', 'Рег. Номер СРО', format)
        worksheet.merge_range('G1:G2', 'Дата регистрации в реестре', format)
        worksheet.merge_range('H1:H2', 'Дата прекращения членства в СРО', format)
        worksheet.merge_range('I1:I2', 'Стоимость работ по одному договору подряда', format)
        worksheet.merge_range('J1:J2', 'Размер обязательств по договорам подряда', format)
        worksheet.write('K2', 'Дата', format)
        worksheet.write('L2', 'статус', format)
        worksheet.write('M2', 'номер свидетельства', format)
        worksheet.write('N2', 'дата выдачи', format)
        worksheet.write('O2', 'стоимость работ по одному ГП', format)
        worksheet.write('P2', 'Статус действия свидетельства', format)
        workbook.close()
    def parcer(self):
        if self.lineEdit.text() == '':
            QMessageBox.critical(self, "Ошибка ", "Выберите файл с ИНН.", QMessageBox.Ok)
        else:
            main.run(self.lineEdit.text(),self.progressBar)
            try:
                self.create_file()
                self.progressBar.setProperty("value", 30)
            except:
                QMessageBox.critical(self, "Ошибка ", "Закройте файл result.xlsx", QMessageBox.Ok)
            self.progressBar.setProperty("value", 100)
            QMessageBox.information(self, "Информация", "Парсинг завершен, результаты записаны в result.xlsx", QMessageBox.Ok)


class mywindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)


if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    application = mywindow()
    application.show()
    sys.exit(app.exec())