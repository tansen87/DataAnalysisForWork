import os
from os.path import join, dirname, abspath

from qtpy import uic, QtWidgets
from qtpy.QtCore import Qt
from qtpy.QtGui import QCursor
from qtpy.QtWidgets import QApplication, QMainWindow, QFileDialog, QAbstractItemView, QMenu

from utils import view_data, pivot_data

_pivotUI = join(dirname(abspath(__file__)), '../modulesUI/pivot_ui.ui')


class PivotWindow(QMainWindow, QtWidgets.QTableView):
    def __init__(self):
        super(PivotWindow, self).__init__()
        uic.loadUi(_pivotUI, self)

        self.btn_open.clicked.connect(self.open)
        self.btn_view.clicked.connect(self.view)
        self.btn_save.clicked.connect(self.save)
        self.btn_pivot.clicked.connect(self.pivot)

        # listWidget添加复制功能
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.showMenu)
        self.contextMenu = QMenu(self)
        self.CP = self.contextMenu.addAction('copy')
        self.CP.triggered.connect(self.copy)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

        # 实时更新lineEdit的值
        self.lineEdit_idx.textEdited[str].connect(lambda: self.real_time_updateValue())
        self.lineEdit_cols.textEdited[str].connect(lambda: self.real_time_updateValue())
        self.lineEdit_vls.textEdited[str].connect(lambda: self.real_time_updateValue())

    def real_time_updateValue(self):
        # 实时更新lineEdit的值, 如果有多个值, 输入","隔开
        self.lineEdit_index = "" .join(self.lineEdit_idx.text())
        self.lineEdit_index = [idx for idx in self.lineEdit_index.split(",")]

        self.lineEdit_columns = "".join(self.lineEdit_cols.text())
        self.lineEdit_columns = [cols for cols in self.lineEdit_columns.split(",")]

        self.lineEdit_values = "".join(self.lineEdit_vls.text())
        self.lineEdit_values = [vls for vls in self.lineEdit_values.split(",")]

    def showMenu(self, pos):
        self.contextMenu.exec(QCursor.pos())  # 在鼠标位置显示

    def selected_text(self):
        # 获取选择行的内容
        try:
            selected = self.getColumnList.selectedItems()
            texts = ''
            for item in selected:
                if texts:
                    texts = texts + '\n' + item.text()
                else:
                    texts = item.text()
        except BaseException as e:
            print(e)
            return
        return texts

    def copy(self):
        text = self.selected_text()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)

    def get_cols_to_listWidget(self, value):
        """ 将表头的value传入listWidget """
        self.getColumnList.clear()
        listWidget_cols = value.split(",")
        self.getColumnList.addItems(listWidget_cols)

    def view_display(self, text):
        self.lable_tips.setText(text)

    def update_text(self, text):
        """ 更新 view_data() 的输出 """
        self.textEdit.append(text)

    def open(self):
        """ 打开excel文件 """
        self.file = QFileDialog.getOpenFileName(
            self, "Select file", "*.csv;*.xlsx;*.xls;*.CSV;*.XLSX;*.XLS")[0]
        self.lineEdit_open.setText(self.file)

    def view(self):
        """ 查看excel数据, 开启子线程 """
        try:
            self.thread_view_data = view_data.ViewDataThread(
                self.tableView,
                self.file)
            self.thread_view_data.signal_trans.connect(self.update_text)
            self.thread_view_data.signal_trans_cols.connect(self.get_cols_to_listWidget)
            self.thread_view_data.signal_desc.connect(self.view_display)
            self.thread_view_data.moveToThread(self.thread_view_data)
            self.thread_view_data.start()
        except Exception as e:
            err = f"Error: {e}"
            self.textEdit_2.setText(err)

    def save(self):
        if not os.path.exists("examples"):
            os.mkdir("examples")
        self.save_path = QFileDialog.getSaveFileName(
            self, "Select the file save path", "examples")[0]
        self.lineEdit_save.setText(self.save_path)

    def pivot(self):
        try:
            self.thread_pt_data = pivot_data.PivotDataThread(
                self.tableView,
                self.file,
                self.lineEdit_index,
                self.lineEdit_values,
                self.save_path)
            self.thread_pt_data.signal_trans.connect(self.update_text)
            self.thread_pt_data.signal_trans_cols.connect(self.get_cols_to_listWidget)
            self.thread_pt_data.signal_desc.connect(self.view_display)
            self.thread_pt_data.moveToThread(self.thread_pt_data)
            self.thread_pt_data.start()
        except Exception as e:
            err = f"Error: {e}"
            self.textEdit_2.setText(err)
