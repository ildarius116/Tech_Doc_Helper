import os
import sys
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QFileDialog, QMainWindow, QInputDialog, QMessageBox
from bill_of_materials import make_bom_gost
from specification import make_spec_gost
from bom_to_buy import buy_bom
from stock_manager import find_in_storage


class Ui_MainWindow(QMainWindow, QFileDialog):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.resize(800, 375)
        MainWindow.setMinimumSize(QtCore.QSize(800, 375))
        MainWindow.setMaximumSize(QtCore.QSize(800, 375))
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(parent=self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(-4, -1, 800, 375))
        self.tabWidget.setObjectName("tabWidget")
        self.bom = QtWidgets.QWidget()
        self.bom.setObjectName("bom")
        self.btn_upload_bom = QtWidgets.QPushButton(parent=self.bom)
        self.btn_upload_bom.setGeometry(QtCore.QRect(440, 30, 75, 23))
        self.btn_upload_bom.setObjectName("btn_upload_bom")
        self.opened_file_name_bom = QtWidgets.QLabel(parent=self.bom)
        self.opened_file_name_bom.setGeometry(QtCore.QRect(10, 30, 420, 20))
        self.opened_file_name_bom.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.opened_file_name_bom.setObjectName("opened_file_name_bom")
        self.saved_file_name_bom = QtWidgets.QLabel(parent=self.bom)
        self.saved_file_name_bom.setGeometry(QtCore.QRect(10, 60, 420, 20))
        self.saved_file_name_bom.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.saved_file_name_bom.setObjectName("saved_file_name_bom")
        self.btn_save_bom = QtWidgets.QPushButton(parent=self.bom)
        self.btn_save_bom.setGeometry(QtCore.QRect(440, 60, 75, 23))
        self.btn_save_bom.setObjectName("btn_save_bom")
        self.btn_result_bom = QtWidgets.QPushButton(parent=self.bom)
        self.btn_result_bom.setGeometry(QtCore.QRect(10, 90, 500, 23))
        self.btn_result_bom.setObjectName("btn_result_bom")
        self.label_excel_bom = QtWidgets.QLabel(parent=self.bom)
        self.label_excel_bom.setGeometry(QtCore.QRect(0, 150, 800, 150))
        self.label_excel_bom.setText("")
        self.label_excel_bom.setPixmap(QtGui.QPixmap("screen_01.JPG"))
        self.label_excel_bom.setScaledContents(True)
        self.label_excel_bom.setObjectName("label_excel_bom")
        self.text_browser_bom = QtWidgets.QTextBrowser(parent=self.bom)
        self.text_browser_bom.setGeometry(QtCore.QRect(0, 0, 800, 350))
        self.text_browser_bom.setStyleSheet("background-color: rgb(191, 191, 191);")
        self.text_browser_bom.setObjectName("text_browser_bom")
        self.label_igf_bom = QtWidgets.QLabel(parent=self.bom)
        self.label_igf_bom.setGeometry(QtCore.QRect(530, 0, 260, 130))
        self.label_igf_bom.setText("")
        self.label_igf_bom.setPixmap(QtGui.QPixmap("igf_grey.jpg"))
        self.label_igf_bom.setScaledContents(True)
        self.label_igf_bom.setObjectName("label_igf_bom")
        self.text_browser_bom.raise_()
        self.btn_upload_bom.raise_()
        self.opened_file_name_bom.raise_()
        self.saved_file_name_bom.raise_()
        self.btn_save_bom.raise_()
        self.btn_result_bom.raise_()
        self.label_excel_bom.raise_()
        self.label_igf_bom.raise_()
        self.tabWidget.addTab(self.bom, "")
        self.spec = QtWidgets.QWidget()
        self.spec.setObjectName("spec")
        self.btn_upload_spec = QtWidgets.QPushButton(parent=self.spec)
        self.btn_upload_spec.setGeometry(QtCore.QRect(440, 30, 75, 23))
        self.btn_upload_spec.setObjectName("btn_upload_spec")
        self.btn_result_spec = QtWidgets.QPushButton(parent=self.spec)
        self.btn_result_spec.setGeometry(QtCore.QRect(10, 90, 500, 23))
        self.btn_result_spec.setObjectName("btn_result_spec")
        self.opened_file_name_spec = QtWidgets.QLabel(parent=self.spec)
        self.opened_file_name_spec.setGeometry(QtCore.QRect(10, 30, 420, 20))
        self.opened_file_name_spec.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.opened_file_name_spec.setObjectName("opened_file_name_spec")
        self.label_excel_spec = QtWidgets.QLabel(parent=self.spec)
        self.label_excel_spec.setGeometry(QtCore.QRect(0, 150, 800, 150))
        self.label_excel_spec.setText("")
        self.label_excel_spec.setPixmap(QtGui.QPixmap("screen_02.JPG"))
        self.label_excel_spec.setScaledContents(True)
        self.label_excel_spec.setObjectName("label_excel_spec")
        self.btn_save_spec = QtWidgets.QPushButton(parent=self.spec)
        self.btn_save_spec.setGeometry(QtCore.QRect(440, 60, 75, 23))
        self.btn_save_spec.setObjectName("btn_save_spec")
        self.saved_file_name_spec = QtWidgets.QLabel(parent=self.spec)
        self.saved_file_name_spec.setGeometry(QtCore.QRect(10, 60, 420, 20))
        self.saved_file_name_spec.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.saved_file_name_spec.setObjectName("saved_file_name_spec")
        self.label_igf_spec = QtWidgets.QLabel(parent=self.spec)
        self.label_igf_spec.setGeometry(QtCore.QRect(530, 0, 260, 130))
        self.label_igf_spec.setText("")
        self.label_igf_spec.setPixmap(QtGui.QPixmap("igf_grey.jpg"))
        self.label_igf_spec.setScaledContents(True)
        self.label_igf_spec.setObjectName("label_igf_spec")
        self.text_browser_spec = QtWidgets.QTextBrowser(parent=self.spec)
        self.text_browser_spec.setGeometry(QtCore.QRect(0, 0, 800, 350))
        self.text_browser_spec.setStyleSheet("background-color: rgb(191, 191, 191);")
        self.text_browser_spec.setObjectName("text_browser_spec")
        self.text_browser_spec.raise_()
        self.btn_upload_spec.raise_()
        self.btn_result_spec.raise_()
        self.opened_file_name_spec.raise_()
        self.label_excel_spec.raise_()
        self.btn_save_spec.raise_()
        self.saved_file_name_spec.raise_()
        self.label_igf_spec.raise_()
        self.tabWidget.addTab(self.spec, "")
        self.buy = QtWidgets.QWidget()
        self.buy.setObjectName("buy")
        self.saved_file_name_buy = QtWidgets.QLabel(parent=self.buy)
        self.saved_file_name_buy.setGeometry(QtCore.QRect(10, 60, 420, 20))
        self.saved_file_name_buy.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.saved_file_name_buy.setObjectName("saved_file_name_buy")
        self.btn_result_buy = QtWidgets.QPushButton(parent=self.buy)
        self.btn_result_buy.setGeometry(QtCore.QRect(10, 90, 500, 23))
        self.btn_result_buy.setObjectName("btn_result_buy")
        self.btn_upload_buy = QtWidgets.QPushButton(parent=self.buy)
        self.btn_upload_buy.setGeometry(QtCore.QRect(440, 30, 75, 23))
        self.btn_upload_buy.setObjectName("btn_upload_buy")
        self.btn_save_buy = QtWidgets.QPushButton(parent=self.buy)
        self.btn_save_buy.setGeometry(QtCore.QRect(440, 60, 75, 23))
        self.btn_save_buy.setObjectName("btn_save_buy")
        self.text_browser_buy = QtWidgets.QTextBrowser(parent=self.buy)
        self.text_browser_buy.setGeometry(QtCore.QRect(0, 0, 800, 350))
        self.text_browser_buy.setStyleSheet("background-color: rgb(191, 191, 191);")
        self.text_browser_buy.setObjectName("text_browser_buy")
        self.label_excel_buy = QtWidgets.QLabel(parent=self.buy)
        self.label_excel_buy.setGeometry(QtCore.QRect(0, 150, 800, 150))
        self.label_excel_buy.setText("")
        self.label_excel_buy.setPixmap(QtGui.QPixmap("screen_02.JPG"))
        self.label_excel_buy.setScaledContents(True)
        self.label_excel_buy.setObjectName("label_excel_buy")
        self.label_igf_buy = QtWidgets.QLabel(parent=self.buy)
        self.label_igf_buy.setGeometry(QtCore.QRect(530, 0, 260, 130))
        self.label_igf_buy.setText("")
        self.label_igf_buy.setPixmap(QtGui.QPixmap("igf_grey.jpg"))
        self.label_igf_buy.setScaledContents(True)
        self.label_igf_buy.setObjectName("label_igf_buy")
        self.opened_file_name_buy = QtWidgets.QLabel(parent=self.buy)
        self.opened_file_name_buy.setGeometry(QtCore.QRect(10, 30, 420, 20))
        self.opened_file_name_buy.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.opened_file_name_buy.setObjectName("opened_file_name_buy")
        self.text_browser_buy.raise_()
        self.saved_file_name_buy.raise_()
        self.btn_result_buy.raise_()
        self.btn_upload_buy.raise_()
        self.btn_save_buy.raise_()
        self.label_excel_buy.raise_()
        self.label_igf_buy.raise_()
        self.opened_file_name_buy.raise_()
        self.tabWidget.addTab(self.buy, "")
        self.store = QtWidgets.QWidget()
        self.store.setObjectName("store")
        self.saved_file_name_store = QtWidgets.QLabel(parent=self.store)
        self.saved_file_name_store.setGeometry(QtCore.QRect(10, 80, 420, 20))
        self.saved_file_name_store.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.saved_file_name_store.setObjectName("saved_file_name_store")
        self.btn_result_store = QtWidgets.QPushButton(parent=self.store)
        self.btn_result_store.setGeometry(QtCore.QRect(10, 110, 500, 23))
        self.btn_result_store.setObjectName("btn_result_store")
        self.btn_upload_bom_store = QtWidgets.QPushButton(parent=self.store)
        self.btn_upload_bom_store.setGeometry(QtCore.QRect(440, 20, 75, 23))
        self.btn_upload_bom_store.setObjectName("btn_upload_bom_store")
        self.btn_save_store = QtWidgets.QPushButton(parent=self.store)
        self.btn_save_store.setGeometry(QtCore.QRect(440, 80, 75, 23))
        self.btn_save_store.setObjectName("btn_save_store")
        self.text_browser_store = QtWidgets.QTextBrowser(parent=self.store)
        self.text_browser_store.setGeometry(QtCore.QRect(0, 0, 800, 350))
        self.text_browser_store.setStyleSheet("background-color: rgb(191, 191, 191);")
        self.text_browser_store.setObjectName("text_browser_store")
        self.label_excel_store = QtWidgets.QLabel(parent=self.store)
        self.label_excel_store.setGeometry(QtCore.QRect(0, 150, 800, 150))
        self.label_excel_store.setText("")
        self.label_excel_store.setPixmap(QtGui.QPixmap("screen_02.JPG"))
        self.label_excel_store.setScaledContents(True)
        self.label_excel_store.setObjectName("label_excel_store")
        self.label_igf_store = QtWidgets.QLabel(parent=self.store)
        self.label_igf_store.setGeometry(QtCore.QRect(530, 0, 260, 130))
        self.label_igf_store.setText("")
        self.label_igf_store.setPixmap(QtGui.QPixmap("igf_grey.jpg"))
        self.label_igf_store.setScaledContents(True)
        self.label_igf_store.setObjectName("label_igf_store")
        self.opened_file_name_bom_store = QtWidgets.QLabel(parent=self.store)
        self.opened_file_name_bom_store.setGeometry(QtCore.QRect(10, 20, 420, 20))
        self.opened_file_name_bom_store.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.opened_file_name_bom_store.setObjectName("opened_file_name_bom_store")
        self.btn_upload_store = QtWidgets.QPushButton(parent=self.store)
        self.btn_upload_store.setGeometry(QtCore.QRect(440, 50, 75, 23))
        self.btn_upload_store.setObjectName("btn_upload_store")
        self.opened_file_name_store = QtWidgets.QLabel(parent=self.store)
        self.opened_file_name_store.setGeometry(QtCore.QRect(10, 50, 420, 20))
        self.opened_file_name_store.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.opened_file_name_store.setObjectName("opened_file_name_store")
        self.text_browser_store.raise_()
        self.btn_save_store.raise_()
        self.opened_file_name_bom_store.raise_()
        self.label_excel_store.raise_()
        self.btn_upload_bom_store.raise_()
        self.saved_file_name_store.raise_()
        self.label_igf_store.raise_()
        self.btn_result_store.raise_()
        self.btn_upload_store.raise_()
        self.opened_file_name_store.raise_()
        self.tabWidget.addTab(self.store, "")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(3)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.path_dir = '/home'
        self.upload_path = ''
        self.save_path = ''

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Технический помощник"))
        self.btn_upload_bom.setText(_translate("MainWindow", "Загрузить"))
        self.opened_file_name_bom.setText(
            _translate("MainWindow", "Нажмите кнопку \"Загрузить\" и выберите файл для загрузки"))
        self.saved_file_name_bom.setText(
            _translate("MainWindow", "Нажмите кнопку \"Сохранить\" и укажите название файла для сохранения"))
        self.btn_save_bom.setText(_translate("MainWindow", "Сохранить"))
        self.btn_result_bom.setText(_translate("MainWindow", "Выполнить"))
        self.text_browser_bom.setHtml(_translate("MainWindow",
                                                 "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                                 "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                                 "p, li { white-space: pre-wrap; }\n"
                                                 "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Названия файлов</p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Стандарт загружаемого файла:</p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-style:italic;\">ПРИМЕЧАНИЕ</span>:   Элементы <span style=\" font-weight:600; color:#aa0000;\">не должны</span> быть сгруппированы</p>\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">                             Сохранение результата происходит в ту же директорию, откуда загружен BOM файл</p></body></html>"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.bom), _translate("MainWindow", "Перечень Элементов ГОСТ"))
        self.btn_upload_spec.setText(_translate("MainWindow", "Загрузить"))
        self.btn_result_spec.setText(_translate("MainWindow", "Выполнить"))
        self.opened_file_name_spec.setText(
            _translate("MainWindow", "Нажмите кнопку \"Загрузить\" и выберите файл для загрузки"))
        self.btn_save_spec.setText(_translate("MainWindow", "Сохранить"))
        self.saved_file_name_spec.setText(
            _translate("MainWindow", "Нажмите кнопку \"Сохранить\" и укажите название файла для сохранения"))
        self.text_browser_spec.setHtml(_translate("MainWindow",
                                                  "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                                  "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                                  "p, li { white-space: pre-wrap; }\n"
                                                  "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                                  "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Названия файлов</p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Стандарт загружаемого файла:</p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                  "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-style:italic;\">ПРИМЕЧАНИЕ</span>:   Элементы <span style=\" font-weight:600; color:#aa0000;\">должны</span> быть сгруппированы</p>\n"
                                                  "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">                             Сохранение результата происходит в ту же директорию, откуда загружен BOM файл</p></body></html>"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.spec), _translate("MainWindow", "Спецификация ГОСТ"))
        self.saved_file_name_buy.setText(
            _translate("MainWindow", "Нажмите кнопку \"Сохранить\" и укажите название файла для сохранения"))
        self.btn_result_buy.setText(_translate("MainWindow", "Выполнить"))
        self.btn_upload_buy.setText(_translate("MainWindow", "Загрузить"))
        self.btn_save_buy.setText(_translate("MainWindow", "Сохранить"))
        self.text_browser_buy.setHtml(_translate("MainWindow",
                                                 "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                                 "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                                 "p, li { white-space: pre-wrap; }\n"
                                                 "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Названия файлов</p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Стандарт загружаемого файла:</p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-style:italic;\">ПРИМЕЧАНИЕ</span>:   Элементы <span style=\" font-weight:600; color:#aa0000;\">должны</span> быть сгруппированы</p>\n"
                                                 "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">                             Сохранение результата происходит в ту же директорию, откуда загружен BOM файл</p></body></html>"))
        self.opened_file_name_buy.setText(
            _translate("MainWindow", "Нажмите кнопку \"Загрузить\" и выберите файл для загрузки"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.buy), _translate("MainWindow", "Перечень на покупку"))
        self.saved_file_name_store.setText(
            _translate("MainWindow", "Нажмите кнопку \"Сохранить\" и укажите название файла для сохранения"))
        self.btn_result_store.setText(_translate("MainWindow", "Выполнить"))
        self.btn_upload_bom_store.setText(_translate("MainWindow", "Загрузить"))
        self.btn_save_store.setText(_translate("MainWindow", "Сохранить"))
        self.text_browser_store.setHtml(_translate("MainWindow",
                                                   "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                                   "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                                   "p, li { white-space: pre-wrap; }\n"
                                                   "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                                   "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Названия файлов</p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">Стандарт загружаемого файла:</p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
                                                   "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-style:italic;\">ПРИМЕЧАНИЕ</span>:   Элементы <span style=\" font-weight:600; color:#aa0000;\">должны</span> быть сгруппированы</p>\n"
                                                   "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">                             Сохранение результата происходит в ту же директорию, откуда загружен BOM файл</p></body></html>"))
        self.opened_file_name_bom_store.setText(
            _translate("MainWindow", "Нажмите кнопку \"Загрузить\" и выберите файл BOM для загрузки"))
        self.btn_upload_store.setText(_translate("MainWindow", "Загрузить"))
        self.opened_file_name_store.setText(
            _translate("MainWindow", "Нажмите кнопку \"Загрузить\" и выберите файл СКЛАД для загрузки"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.store), _translate("MainWindow", "Поиск на складе"))

        self.btn_upload_bom.clicked.connect(lambda: self.event_btn_upload_clicked('bom'))
        self.btn_upload_spec.clicked.connect(lambda: self.event_btn_upload_clicked('spec'))
        self.btn_upload_buy.clicked.connect(lambda: self.event_btn_upload_clicked('buy'))
        self.btn_upload_bom_store.clicked.connect(lambda: self.event_btn_upload_clicked('store'))
        self.btn_upload_store.clicked.connect(self.event_btn_upload_storage_clicked)

        self.btn_save_bom.clicked.connect(lambda: self.event_btn_save_clicked('bom'))
        self.btn_save_spec.clicked.connect(lambda: self.event_btn_save_clicked('spec'))
        self.btn_save_buy.clicked.connect(lambda: self.event_btn_save_clicked('buy'))
        self.btn_save_store.clicked.connect(lambda: self.event_btn_save_clicked('store'))

        self.btn_result_bom.clicked.connect(lambda: self.event_btn_result_clicked('bom'))
        self.btn_result_spec.clicked.connect(lambda: self.event_btn_result_clicked('spec'))
        self.btn_result_buy.clicked.connect(lambda: self.event_btn_result_clicked('buy'))
        self.btn_result_store.clicked.connect(lambda: self.event_btn_result_clicked('store'))

    def event_btn_upload_clicked(self, page):
        fname = QFileDialog.getOpenFileName(self, 'Выберите файл для загрузки', self.path_dir, 'Excel Files (*.xls)')
        self.upload_path, *_ = fname
        self.path_dir, path_file = os.path.split(self.upload_path)
        if not self.path_dir:
            path_file = 'Нажмите кнопку "Загрузить" и выберите файл для загрузки'
        if page == 'bom':
            self.opened_file_name_bom.setText(path_file)
        elif page == 'spec':
            self.opened_file_name_spec.setText(path_file)
        elif page == 'buy':
            self.opened_file_name_buy.setText(path_file)
        elif page == 'store':
            self.opened_file_name_bom_store.setText(path_file)

    def event_btn_upload_storage_clicked(self):
        fname = QFileDialog.getOpenFileName(self, 'Выберите файл для загрузки', self.path_dir, 'Excel Files (*.xls)')
        self.upload_path_storage, *_ = fname
        self.path_storage_dir, path_file = os.path.split(self.upload_path_storage)
        if not self.path_storage_dir:
            path_file = 'Нажмите кнопку "Загрузить" и выберите файл для загрузки'
        self.opened_file_name_store.setText(path_file)

    def event_btn_save_clicked(self, page):
        pattern = "!@#$%^&()+,\/:;*?<>|'"
        self.save_fname = QInputDialog.getText(self, 'Сохранить', 'Напишите название файла:')
        file_name, btn_res = self.save_fname
        for sym in pattern:
            if sym in file_name:
                error = QMessageBox()
                error.setWindowTitle('Ошибка')
                error.setText('Не верно введено название файла')
                error.setIcon(QMessageBox.Icon.Warning)
                error.setStandardButtons(QMessageBox.StandardButton.Ok)
                error.setDetailedText(f'В названии файла не должно быть символов: "{pattern}"')
                btn_res = ''
                error.exec()
        if btn_res:
            file_name += '.xls'
            self.save_path = self.path_dir + '/' + file_name
        else:
            file_name = 'Нажмите кнопку "Сохранить" и укажите название файла для сохранения'
        if page == 'bom':
            self.saved_file_name_bom.setText(file_name)
        elif page == 'spec':
            self.saved_file_name_spec.setText(file_name)
        elif page == 'buy':
            self.saved_file_name_buy.setText(file_name)
        elif page == 'store':
            self.saved_file_name_store.setText(file_name)

    def event_btn_result_clicked(self, page):
        print("upload_path:", self.upload_path)
        print("save_path:", self.save_path)
        try:
            if page == 'bom':
                make_bom_gost(self.upload_path, self.save_path)
            elif page == 'spec':
                make_spec_gost(self.upload_path, self.save_path)
            elif page == 'buy':
                buy_bom(self.upload_path, self.save_path)
            elif page == 'store':
                find_in_storage(self.upload_path, self.upload_path_storage, self.save_path)
        except Exception as e:
            error = QMessageBox()
            error.setWindowTitle('Ошибка')
            error.setText(f'Ошибка в стандарте загруженного файла')
            error.setIcon(QMessageBox.Icon.Warning)
            error.setStandardButtons(QMessageBox.StandardButton.Ok)
            error.setDetailedText(f'{e.args}')
            error.exec()


if __name__ == "__main__":
    import sys

    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
