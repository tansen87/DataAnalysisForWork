import os
import sys
from os.path import join, dirname, abspath

from qtpy import uic, QtWidgets
from qtpy.QtCore import Slot, Qt
from qtpy.QtGui import QIcon
from qtpy.QtWidgets import QApplication, QMainWindow, QFileDialog

import qtmodern.styles
import qtmodern.windows
from modulesUI.ui_main import Ui_MainWindow
from utils import (record_logs, view_data, screen_data, merge_data, concat_data)


_UI = join(dirname(abspath(__file__)), 'modulesUI/main_ui.ui')


class MainWindow(QMainWindow, Ui_MainWindow, QtWidgets.QTableView):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi(_UI, self)  # load the ui into self
        self.log = record_logs.RecordLog  # logs

        self.actionLight.triggered.connect(self.lightTheme)  # load light theme
        self.actionDark.triggered.connect(self.darkTheme)  # load dark theme

        self.actionFile.triggered.connect(self.openfile)  # connect openfile function
        self.actionSave.triggered.connect(self.save_path)  # connect save_path function
        self.actionTxt.triggered.connect(self.open_target)  # connect open_target function
        self.actionCheckData.triggered.connect(self.view_data)  # connect view_data function
        self.actionScreenData.triggered.connect(self.screen_data)  # connect screen_data function

        self.actionInitTable.triggered.connect(self.open_init_table)  # connect open_init_table function
        self.actionMappingTable.triggered.connect(self.open_map_table)  # connect open_map_table function
        self.actionMerge.triggered.connect(self.merge_data)  # connect merge_data function
        self.actionCheck.triggered.connect(self.check)  # connect check function

        self.actionMergePath.triggered.connect(self.save_path)  # connect save_path function
        self.actionMergeSavePath.triggered.connect(self.merge_save_path)  # connect merge_save_path function
        self.actionMergeFile.triggered.connect(self.concat_data)  # connect concat_data function

        self.progress_bar()
        self.col_display()

    def keyPressEvent(self, event):  # 重写键盘监听事件
        # 监听 CTRL+C 组合键，实现复制数据到粘贴板
        if (event.key() == Qt.Key_C) and QApplication.keyboardModifiers() == Qt.ControlModifier:
            text = self.selected_tb_text(self.tableWidget)  # 获取当前表格选中的数据
            if text:
                try:
                    clipboard = QApplication.clipboard()
                    clipboard.setText(text)  # 复制到粘贴板
                except BaseException as e:
                    err = f"Error: {e}"
                    self.tips.setText(err)

    def selected_tb_text(self, table_view):
        """ 复制选择表格数据 """
        try:
            # 获取表格对象中被选中的数据索引列表
            indexes = table_view.selectedIndexes()
            indexes_dict = {}
            for index in indexes:  # 遍历每个单元格
                row, column = index.row(), index.column()  # 获取单元格的行号，列号
                if row in indexes_dict.keys():
                    indexes_dict[row].append(column)
                else:
                    indexes_dict[row] = [column]

            # 将数据表数据用制表符(\t)和换行符(\n)连接，使其可以复制到excel文件中
            text = ''
            for row, columns in indexes_dict.items():
                row_data = ''
                for column in columns:
                    data = table_view.model().item(row, column).text()
                    if row_data:
                        row_data = row_data + '\t' + data
                    else:
                        row_data = data

                if text:
                    text = text + '\n' + row_data
                else:
                    text = row_data
            return text
        except BaseException as e:
            err = f"Error: {e}"
            self.tips.setText(err)

    def progress_bar(self):
        """ progressBar """
        self.progressBar.setRange(0, 99)
        self.progressBar.hide()

    def callback(self, i):
        """ 实时更新 progressBar """
        self.progressBar.show()
        self.progressBar.setValue(i)

    def callback_done(self, i):
        """ 返回 progressBar 信息 """
        self.is_done = i
        if self.is_done == 1:
            self.progressBar.show()
            self.progressBar.setValue(99)
            self.progressBar.setFormat("completed")

    def open_init_table(self):
        """ 打开init表"""
        self.init_table = QFileDialog.getOpenFileName(
            self, "Select file", "*.csv;*.xlsx;*.xls;*.CSV;*.XLSX;*.XLS")[0]
        file_ = f"Init Table: {os.path.basename(self.init_table)}"
        self.openFilePath.setText(file_)
        self.log.write_log(file_)

    def open_map_table(self):
        """ 打开mapping表 """
        self.mapping_table = QFileDialog.getOpenFileName(
            self, "Select file", "*.csv;*.xlsx;*.xls;*.CSV;*.XLSX;*.XLS")[0]
        file_ = f"Mapping Table: {os.path.basename(self.mapping_table)}"
        self.openFilePath.setText(file_)
        self.log.write_log(file_)

    def openfile(self):
        """ 打开excel文件 """
        self.file = QFileDialog.getOpenFileName(
            self, "Select file", "*.csv;*.xlsx;*.xls;*.CSV;*.XLSX;*.XLS")[0]
        file_ = f"open file name: {os.path.basename(self.file)}"
        self.openFilePath.setText(file_)
        self.log.write_log(file_)

    def save_path(self):
        """ 选择文件保存路径 """
        self.file_out = QFileDialog.getExistingDirectory(
            self, "Select the file save path", "./")
        output_ = f"file path: {self.file_out}"
        self.fileSavePath.setText(output_)
        self.log.write_log(output_)

    def merge_save_path(self):
        self.file_out_mer = QFileDialog.getSaveFileName(
            self, "Select the file save path", "./")[0]
        output_ = f"file path: {self.file_out_mer}"
        self.fileSavePath.setText(output_)
        self.log.write_log(output_)

    def open_target(self):
        target = QFileDialog.getOpenFileName(
            self, "Select txt file", "*.txt")[0]
        target_ = f"screening conditions: {os.path.basename(target)}"
        self.txtPath.setText(target_)
        self.log.write_log(target_)
        try:
            with open(target, 'r', encoding='utf-8') as fp:
                lines = fp.readlines()
            self.targets = [line.strip('\n') for line in lines]
        except:
            err = f"Warning: No txt file opened."
            self.txtPath.setText(err)
            self.log.write_log(err)

    @staticmethod
    def lightTheme():
        qtmodern.styles.light(QApplication.instance())

    @staticmethod
    def darkTheme():
        qtmodern.styles.dark(QApplication.instance())

    @Slot()
    def closeEvent(self, event):
        mw.close()

    def get_cols(self, value):
        self.columnChoice.clear()
        lst_cols = value.split(",")
        self.columnChoice.addItems(lst_cols)
        # self.tips.setText(self.columnChoice.currentText())
        self.columnChoice.activated[str].connect(self.get_cols_value)

    def get_cols_value(self, value):
        self.column_value = value
        self.tips.setText(value)

    def view_display(self, text):
        self.dispalyDesc.setText(text)

    def get_cols_map(self, value):
        self.symbolChoice.clear()
        lst_cols = value.split(",")
        self.symbolChoice.addItems(lst_cols)
        # self.tips.setText(self.symbolChoice.currentText())
        self.symbolChoice.activated[str].connect(self.get_cols_value)

    def get_cols_map_value(self, value):
        global column_map_value
        column_map_value = value
        self.tips.setText(value)

    def col_display(self):
        self.leftLineEdit.setPlaceholderText("init表列名")
        self.rightLineEdit.setPlaceholderText("mapping表列名")

    def check(self):
        try:
            self.left = self.leftLineEdit.text()
            self.right = self.rightLineEdit.text()
            if self.left == "" or self.right == "":
                col_ = f"Error: init表列名和mapping表列名不能为空"
                self.recoedLog.setText(col_)
            else:
                col_ = f"init column: {self.left}, map column: {self.right}"
                self.tips.setText(col_)
            self.log.write_log(col_)
        except Exception as e:
            err = f"Error: {e}"
            self.tips.setText(err)

    def view_data(self):
        """ 查看excel数据, 开启子线程 """
        try:
            self.thread_view_data = view_data.ViewDataThread(
                self.tableWidget,
                self.file)
            self.thread_view_data.signal_trans.connect(self.update_text)
            self.thread_view_data.signal_trans_cols.connect(self.get_cols)
            self.thread_view_data.signal_desc.connect(self.view_display)
            self.thread_view_data.moveToThread(self.thread_view_data)
            self.thread_view_data.start()
        except Exception as e:
            err = f"Error: {e}"
            self.recoedLog.setText(err)

    def screen_data(self):
        """ 筛分数据, 开启子线程 """
        self.progressBar.setValue(0)
        self.is_done = 0
        try:
            self.thread_screen_data = screen_data.ScreenDataThread(
                self.file,
                self.file_out,
                self.targets,
                self.column_value)
            self.thread_screen_data.signal_trans.connect(self.update_text)
            self.thread_screen_data.progressBarValue.connect(self.callback)
            self.thread_screen_data.signal_done.connect(self.callback_done)
            self.thread_screen_data.signal_desc.connect(self.view_display)
            self.thread_screen_data.moveToThread(self.thread_screen_data)
            self.thread_screen_data.start()
        except Exception as e:
            err = f"Error: {e}"
            self.recoedLog.setText(err)

    def merge_data(self):
        """ VLookup, 开启子线程 """
        try:
            self.thread_merge_data = merge_data.MergeDataThread(
                self.tableWidget,
                self.init_table,
                self.mapping_table,
                self.left,
                self.right)
            self.thread_merge_data.signal_trans.connect(self.update_text)
            self.thread_merge_data.signal_trans_cols.connect(self.get_cols)
            self.thread_merge_data.signal_trans_cols_map.connect(self.get_cols_map)
            self.thread_merge_data.signal_desc.connect(self.view_display)
            self.thread_merge_data.moveToThread(self.thread_merge_data)
            self.thread_merge_data.start()
        except Exception as e:
            err = f"Error: {e}"
            self.recoedLog.setText(err)

    def concat_data(self):
        """ 合并数据, 开启子线程 """
        try:
            self.thread_concat_data = concat_data.ConcatDataThread(
                self.file_out,
                self.file_out_mer)
            self.thread_concat_data.signal_trans.connect(self.update_text)
            self.thread_concat_data.moveToThread(self.thread_concat_data)
            self.thread_concat_data.start()
        except Exception as e:
            err = f"Error: {e}"
            self.recoedLog.setText(err)

    def update_text(self, text):
        """ 更新 view_data() screen_data() 的输出 """
        self.recoedLog.append(text)


if __name__ == '__main__':
    app = QApplication(sys.argv)

    qtmodern.styles.dark(app)
    mw = qtmodern.windows.ModernWindow(MainWindow())
    mw.setWindowIcon(QIcon("modulesUI/excel.svg"))
    mw.setWindowTitle("ExcelTools")
    mw.show()

    sys.exit(app.exec())
