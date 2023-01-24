import os
from os.path import join, dirname, abspath

from qtpy import uic, QtWidgets
from qtpy.QtCore import Qt, QUrl
from qtpy.QtGui import QIcon, QCursor
from qtpy.QtWebEngineWidgets import QWebEngineView
from qtpy.QtWidgets import QApplication, QMainWindow, QFileDialog, QAbstractItemView, QMenu

import qtmodern.styles
import qtmodern.windows
from utils import view_data, draw_pic

_plotUI = join(dirname(abspath(__file__)), '../modulesUI/plot_ui.ui')


class VisualWindow(QMainWindow, QtWidgets.QTableView):
    def __init__(self):
        super().__init__()
        uic.loadUi(_plotUI, self)

        self.btn_open.clicked.connect(self.open)
        self.btn_view.clicked.connect(self.view)
        self.btn_draw.clicked.connect(self.draw_pyecharts)
        self.btn_save.clicked.connect(self.save)
        self.btn_plot.clicked.connect(self.plot)

        # self.getColumnList.setDragDropMode(QtWidgets.QAbstractItemView.DragDropMode.DragDrop)
        # self.getColumnList.setDefaultDropAction(Qt.DropAction.CopyAction)

        # listWidget添加复制功能
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.showMenu)
        self.contextMenu = QMenu(self)
        self.CP = self.contextMenu.addAction('copy')
        self.CP.triggered.connect(self.copy)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

        # 实时更新lineEdit的值
        self.lineEdit_x.textEdited[str].connect(lambda: self.real_time_updateValue())
        self.lineEdit_y.textEdited[str].connect(lambda: self.real_time_updateValue())
        self.lineEdit_name.textEdited[str].connect(lambda: self.real_time_updateValue())

    def real_time_updateValue(self):
        # 实时更新lineEdit的值
        self.real_lineEdit_x = self.lineEdit_x.text()
        self.read_lineEdit_y = self.lineEdit_y.text()
        self.real_title_name = self.lineEdit_name.text()

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

    def get_cols(self, value):
        self.comboBox.clear()
        lst_cols = value.split(",")
        self.comboBox.addItems(lst_cols)
        self.comboBox.activated[str].connect(self.get_cols_value)

    def get_cols_value(self, value):
        self.column_value = value
        self.lable_tips.setText(value)

    def get_cols_to_listWidget(self, value):
        """ 将表头的value传入listWidget """
        self.getColumnList.clear()
        listWidget_cols = value.split(",")
        self.getColumnList.addItems(listWidget_cols)

    def view_display(self, text):
        self.lable_tips.setText(text)

    def view_pyecharts(self, text):
        self.lineEdit_pyecharts.setText(text)

    def view_matplotlib(self, text):
        self.lineEdit_matplotlib.setText(text)

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
            self.thread_view_data.signal_trans_cols.connect(self.get_cols)
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

    def draw_pyecharts(self):
        try:
            self.thread_pyecharts = draw_pic.DrawPicture(
                self.file,
                self.save_path,
                self.real_lineEdit_x,
                self.real_title_name)
            self.thread_pyecharts.signal_trans.connect(self.view_pyecharts)
            self.thread_pyecharts.signal_desc.connect(self.view_display)
            self.thread_pyecharts.moveToThread(self.thread_pyecharts)
            self.thread_pyecharts.start()
        except Exception as e:
            err = f"Error: {e}"
            self.textEdit_2.setText(err)

    def plot(self):
        try:
            self.plot_win = qtmodern.windows.ModernWindow(PlotWindow(self.save_path))
            self.plot_win.setWindowIcon(QIcon("modulesUI/img/visual.svg"))
            self.plot_win.setWindowTitle("plot")
            self.plot_win.show()
        except Exception as e:
            err = f"Error: {e}"
            self.textEdit_2.setText(err)

    def update_text(self, text):
        """ 更新 view_data() 的输出 """
        self.textEdit.append(text)


class PlotWindow(QMainWindow):
    def __init__(self, path):
        super().__init__()
        uic.loadUi(_plotUI, self)

        self.browser = QWebEngineView()
        url = f"{os.path.splitext(path)[0]}.html"
        self.browser.load(QUrl.fromLocalFile(url))
        self.setCentralWidget(self.browser)
