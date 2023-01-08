import datetime
import sys
import os
from os.path import join, dirname, abspath

from qtpy import uic, QtWidgets
from qtpy.QtGui import QStandardItemModel, QStandardItem
from qtpy.QtCore import Slot, QThread, Signal, Qt
from qtpy.QtWidgets import QApplication, QMainWindow, QFileDialog

import qtmodern.styles
import qtmodern.windows
from module.ui_main import Ui_MainWindow

global file_name, file_out, column_value, targets, column_map_value, init_table, mapping_table, col_left, col_right, \
        file_out_mer

_UI = join(dirname(abspath(__file__)), './module/main_ui.ui')


class MainWindow(QMainWindow, Ui_MainWindow, QtWidgets.QTableView):
    def __init__(self):
        super(MainWindow, self).__init__()
        uic.loadUi(_UI, self)  # load the ui into self
        self.log = RecordLog  # logs

        self.actionLight.triggered.connect(self.lightTheme)  # load light theme
        self.actionDark.triggered.connect(self.darkTheme)  # load dark theme

        self.actionFile.triggered.connect(self.openfile)  # connect openfile function
        self.actionSave.triggered.connect(self.save_path)  # connect save_path function
        self.actionTxt.triggered.connect(self.open_target)  # connect open_target function
        self.actionCheckData.triggered.connect(self.view_data)  # connect view_data function
        self.actionScreenData.triggered.connect(self.screen_data)  # connect screen_data function

        self.actionInitTable.triggered.connect(self.open_init_table)  # connect open_init_table function
        self.actionMappingTable.triggered.connect(self.open_map_table)  # connect open_map_table function
        self.actionMerge.triggered.connect(self.vlookup_data)  # connect vlookup_data function
        self.actionCheck.triggered.connect(self.check)  # connect check function

        self.actionMergePath.triggered.connect(self.save_path)  # connect save_path function
        self.actionMergeSavePath.triggered.connect(self.merge_save_path)  # connect merge_save_path function
        self.actionMergeFile.triggered.connect(self.merge_data)  # connect merge_data function

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
            self.progressBar.setFormat("screening completed")

    def open_init_table(self):
        """ 打开init表"""
        global init_table
        file = QFileDialog.getOpenFileName(self, "Select file", "*.csv;*.xlsx;*.xls;*.CSV;*.XLSX;*.XLS")[0]
        file_ = f"Init Table: {os.path.basename(file)}"
        self.openFilePath.setText(file_)
        self.log.write_log(file_)
        init_table = file

    def open_map_table(self):
        """ 打开mapping表 """
        global mapping_table
        file = QFileDialog.getOpenFileName(self, "Select file", "*.csv;*.xlsx;*.xls;*.CSV;*.XLSX;*.XLS")[0]
        file_ = f"Mapping Table: {os.path.basename(file)}"
        self.openFilePath.setText(file_)
        self.log.write_log(file_)
        mapping_table = file

    def openfile(self):
        """ 打开excel文件 """
        global file_name
        file = QFileDialog.getOpenFileName(self, "Select file", "*.csv;*.xlsx;*.xls;*.CSV;*.XLSX;*.XLS")[0]
        file_ = f"open file name: {os.path.basename(file)}"
        self.openFilePath.setText(file_)
        self.log.write_log(file_)
        file_name = file

    def save_path(self):
        """ 选择文件保存路径 """
        global file_out
        output = QFileDialog.getExistingDirectory(self, "Select the file save path", "./")
        output_ = f"file path: {output}"
        self.fileSavePath.setText(output_)
        self.log.write_log(output_)
        file_out = output

    def merge_save_path(self):
        global file_out_mer
        output = QFileDialog.getExistingDirectory(self, "Select the file save path", "./")
        output_ = f"file path: {output}"
        self.fileSavePath.setText(output_)
        self.log.write_log(output_)
        file_out_mer = output

    def open_target(self):
        global targets
        target = QFileDialog.getOpenFileName(self, "Select txt file", "*.txt")[0]
        target_ = f"screening conditions: {os.path.basename(target)}"
        self.txtPath.setText(target_)
        self.log.write_log(target_)
        file_target = target
        with open(file_target, 'r', encoding='utf-8') as fp:
            lines = fp.readlines()
        targets = [line.strip('\n') for line in lines]

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
        self.tips.setText(self.columnChoice.currentText())
        self.columnChoice.activated[str].connect(self.get_cols_value)

    def get_cols_value(self, value):
        global column_value
        column_value = value
        self.tips.setText(value)

    def view_display(self, text):
        self.dispalyDesc.setText(text)

    def get_cols_map(self, value):
        self.symbolChoice.clear()
        lst_cols = value.split(",")
        self.symbolChoice.addItems(lst_cols)
        self.tips.setText(self.symbolChoice.currentText())
        self.symbolChoice.activated[str].connect(self.get_cols_value)

    def get_cols_map_value(self, value):
        global column_map_value
        column_map_value = value
        self.tips.setText(value)

    def col_display(self):
        self.leftLineEdit.setPlaceholderText("init表列名")
        self.rightLineEdit.setPlaceholderText("mapping表列名")

    def check(self):
        global col_left, col_right
        try:
            left = self.leftLineEdit.text()
            right = self.rightLineEdit.text()
            col_left = left
            col_right = right
            if left == "" or right == "":
                col_ = f"Error: init表列名和mapping表列名不能为空"
                self.dispalyDesc.setText(col_)
            else:
                col_ = f"init column: {left}, map column: {right}"
                self.tips.setText(col_)
            self.log.write_log(col_)
        except Exception as e:
            err = f"Error: {e}"
            self.tips.setText(err)

    def view_data(self):
        """ 查看excel数据, 开启子线程 """
        self.thread_view_data = ViewDataThread(self.tableWidget)
        self.thread_view_data.signal_trans.connect(self.update_text)
        self.thread_view_data.signal_trans_cols.connect(self.get_cols)
        self.thread_view_data.signal_desc.connect(self.view_display)
        self.thread_view_data.moveToThread(self.thread_view_data)
        self.thread_view_data.start()

    def screen_data(self):
        """ 筛分数据, 开启子线程 """
        self.progressBar.setValue(0)
        self.is_done = 0
        self.thread_screen_data = ScreenDataThread()
        self.thread_screen_data.signal_trans.connect(self.update_text)
        self.thread_screen_data.progressBarValue.connect(self.callback)
        self.thread_screen_data.signal_done.connect(self.callback_done)
        self.thread_screen_data.signal_desc.connect(self.view_display)
        self.thread_screen_data.moveToThread(self.thread_screen_data)
        self.thread_screen_data.start()

    def vlookup_data(self):
        self.thread_vlookup_data = VLookupThread(self.tableWidget)
        self.thread_vlookup_data.signal_trans.connect(self.update_text)
        self.thread_vlookup_data.signal_trans_cols.connect(self.get_cols)
        self.thread_vlookup_data.signal_trans_cols_map.connect(self.get_cols_map)
        self.thread_vlookup_data.signal_desc.connect(self.view_display)
        self.thread_vlookup_data.moveToThread(self.thread_vlookup_data)
        self.thread_vlookup_data.start()

    def merge_data(self):
        self.thread_merge = MergeDataThread()
        self.thread_merge.signal_trans.connect(self.update_text)
        self.thread_merge.moveToThread(self.thread_merge)
        self.thread_merge.start()

    def update_text(self, text):
        """ 更新 view_data() screen_data() 的输出 """
        self.recoedLog.append(text)


class ViewDataThread(QThread):
    """ 查看excel数据 """
    signal_trans = Signal(str)
    signal_trans_cols = Signal(str)
    signal_desc = Signal(str)

    def __init__(self, tableWidget):
        super(ViewDataThread, self).__init__()
        self.tableWidget = tableWidget
        self.log = RecordLog

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        import pandas as pd
        import numpy as np

        try:
            load_tips = f"Reading {os.path.basename(file_name)}..."
            self.signal_desc.emit(load_tips)
            file = os.path.splitext(file_name)
            if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                df = pd.read_excel(file_name, dtype=object)
            if file[1].lower() == '.csv':
                df = pd.read_csv(file_name, dtype=object)
            read_tips = f"File name: {os.path.basename(file_name)}"
            self.write(read_tips)

            rows = df.shape[0]
            cols = df.shape[1]

            excel_detail = f"rows: {rows}, cols: {cols}."
            self.write(excel_detail)
            self.log.write_log(excel_detail)

            col_names = [df.columns[col] for col in range(len(df.columns))]

            str_col = ",".join(col_names)
            self.signal_trans_cols.emit(str_col)

            # 对excel行数进行判断, if行数大于5w行, 行数设置为5w行; 否则以实际为准
            if rows > 30_000:
                row = 30_000
                # 设置数据层次结构，row行cols列
                model = QStandardItemModel(row, cols)
            else:
                row = rows
                # 设置数据层次结构，row行cols列
                model = QStandardItemModel(row, cols)
            # 设置列名
            model.setHorizontalHeaderLabels([str(i) for i in col_names])

            # 将excel数据显示在tabelWidget中
            for i in range(row):
                row_values = df.iloc[[i]]
                row_values_array = np.array(row_values)
                row_values_list = row_values_array.tolist()[0]
                for j in range(cols):
                    newItem = QStandardItem(str(row_values_list[j]))
                    model.setItem(i, j, newItem)
            # 实例化表格视图，设置模型为自定义的模型
            self.tableWidget.setModel(model)

        except Exception as e:
            err = f"Error: {e}"
            self.signal_desc.emit(err)
        else:
            success_tips = f"{os.path.basename(file_name)} read successfully"
            self.signal_desc.emit(success_tips)


class ScreenDataThread(QThread):
    signal_trans = Signal(str)
    progressBarValue = Signal(int)
    signal_done = Signal(int)
    signal_desc = Signal(str)

    def __init__(self):
        super(ScreenDataThread, self).__init__()
        self.log = RecordLog

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        self.signal_desc.emit('开始运行...')
        import pandas as pd
        try:
            file = os.path.splitext(file_name)
            if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                df = pd.read_excel(file_name, dtype=object)
            if file[1].lower() == '.csv':
                df = pd.read_csv(file_name, dtype=object)
            ls_col = [df.columns[col] for col in range(len(df.columns))]
            read_tips = f"{os.path.basename(file_name)} read successfully."
            self.write(read_tips)
            self.log.write_log(read_tips)
            data_fr = pd.DataFrame(df, columns=ls_col)
            for i in range(len(targets)):
                cd = {
                    column_value: targets[i]
                }
                new_df = self.df_screen(data_fr, cd)
                try:
                    new_df.to_excel(f"{file_out}/{str(targets[i])}.xlsx", encoding='gbk', index=False)
                    out = f"{str(i + 1)} {str(targets[i])}.xlsx is saved."
                    self.write(out)
                    self.log.write_log(out)
                    self.progressBarValue.emit(int(i / len(targets) * 100))
                except:
                    new_df.to_csv(f"{file_out}/{str(targets[i])}.csv", encoding="utf-8", index=False)
                    out = f"{str(i + 1)} {str(targets[i])}.csv is saved."
                    self.write(out)
                    self.log.write_log(out)
                    self.progressBarValue.emit(int(i / len(targets) * 100))
                else:
                    pass
            self.signal_done.emit(1)

        except Exception as e:
            err = f"Error: {e}"
            self.write(err)
            self.log.write_log(err)
        else:
            com_tips = f"{len(targets)} files split."
            self.write(com_tips)
            self.log.write_log(com_tips)
        self.signal_desc.emit('运行结束')

    @staticmethod
    def df_screen(data, cd_data):
        df_cy = data.copy()
        idx_z = [True for i in df_cy.index]
        orcom = lambda a, b: [any([a[i], b[i]]) for i in range(len(a))]  # 列表a与列表b 或 比较
        addcom = lambda a, b: [all([a[i], b[i]]) for i in range(len(a))]  # 列表a与列表b 与 比较
        for z in cd_data:
            if isinstance(cd_data[z], list):
                for index, c in enumerate(cd_data[z]):
                    if index != 0:
                        idx_c = orcom(idx_c, list(df_cy[z] == c))
                    else:
                        idx_c = list(df_cy[z] == c)
            else:
                idx_c = list(df_cy[z] == cd_data[z])
            idx_z = addcom(idx_z, idx_c)
        return df_cy.loc[idx_z, :]


class VLookupThread(QThread):
    signal_trans = Signal(str)
    signal_trans_cols = Signal(str)
    signal_trans_cols_map = Signal(str)
    signal_desc = Signal(str)

    def __init__(self, table):
        super(VLookupThread, self).__init__()
        self.log = RecordLog
        self.tableWidget = table

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        import pandas as pd
        import numpy as np

        try:
            file = os.path.splitext(init_table)
            if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                df_init = pd.read_excel(init_table, dtype=object)
            if file[1].lower() == '.csv':
                df_init = pd.read_csv(init_table, dtype=object)
            read_tips = f"File name: {os.path.basename(init_table)}"
            self.write(read_tips)

            col_init = [df_init.columns[col] for col in range(len(df_init.columns))]
            str_col_init = ",".join(col_init)
            self.signal_trans_cols.emit(str_col_init)
            self.write(f"{os.path.basename(init_table)} read successfully.")
        except Exception as e:
            err = f"Error: {e}"
            self.write(err)

        else:
            try:
                file = os.path.splitext(mapping_table)
                if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                    df_map = pd.read_excel(mapping_table, dtype=object)
                if file[1].lower() == '.csv':
                    df_map = pd.read_csv(mapping_table, dtype=object)
                read_tips = f"File name: {os.path.basename(mapping_table)}"
                self.write(read_tips)

                col_map = [df_map.columns[col] for col in range(len(df_map.columns))]
                str_col_map = ",".join(col_map)
                self.signal_trans_cols_map.emit(str_col_map)
                self.write(f"{os.path.basename(mapping_table)} read successfully.")

                df_merge = pd.merge(df_init, df_map, left_on=col_left, right_on=col_right, how='left')

                rows = df_merge.shape[0]
                cols = df_merge.shape[1]

                excel_detail = f"rows: {rows}, cols: {cols}."
                self.write(excel_detail)
                self.log.write_log(excel_detail)

                col_names = [df_merge.columns[col] for col in range(len(df_merge.columns))]
                # 对excel行数进行判断, if行数大于5w行, 行数设置为5w行; 否则以实际为准
                if rows > 30_000:
                    row = 30_000
                    # 设置数据层次结构，row行cols列
                    model = QStandardItemModel(row, cols)
                else:
                    row = rows
                    # 设置数据层次结构，row行cols列
                    model = QStandardItemModel(row, cols)
                # 设置列名
                model.setHorizontalHeaderLabels([str(i) for i in col_names])

                # 将excel数据显示在tabelWidget中
                for i in range(row):
                    row_values = df_merge.iloc[[i]]
                    row_values_array = np.array(row_values)
                    row_values_list = row_values_array.tolist()[0]
                    for j in range(cols):
                        newItem = QStandardItem(str(row_values_list[j]))
                        model.setItem(i, j, newItem)
                # 实例化表格视图，设置模型为自定义的模型
                self.tableWidget.setModel(model)
                now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                df_merge.to_csv(f"{now}_vlookup.csv", index=False, encoding="utf-8")
                succ_tips = f"{now}_vlookup.csv is saved."
                self.write(succ_tips)
            except Exception as e:
                err = f"Error: {e}"
                self.write(err)


class MergeDataThread(QThread):
    signal_trans = Signal(str)

    def __init__(self):
        super(MergeDataThread, self).__init__()
        self.log = RecordLog

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        import pandas as pd
        try:
            lst_file = list()
            init_path = f"{file_out}/"
            files = os.listdir(init_path)
            for file in range(len(files)):
                read_file = f"{init_path}{files[file]}"
                try:
                    df = pd.read_excel(read_file, dtype=object)
                except:
                    df = pd.read_csv(read_file, dtype=object)
                lst_file.append(df)
                read_tips = f"{file + 1} {os.path.basename(read_file)} read successfully"
                self.write(read_tips)
                self.log.write_log(read_tips)
            all_file = pd.concat(lst_file, axis=0, ignore_index=True)
            now = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            try:
                all_file.to_excel(f"{file_out_mer}/{now}_merged.xlsx", index=False, encoding="gbk")
                com_tips = f"file save path: {file_out_mer}, Merge File Name: '{now}_merged.xlsx'."
                self.write(com_tips)
                self.log.write_log(com_tips)
            except:
                all_file.to_csv(f"{file_out_mer}/{now}_merged.csv", index=False, encoding="utf-8")
                com_tips = f"file save path: {file_out_mer}, Merge File Name: '{now}_merged.csv'."
                self.write(com_tips)
                self.log.write_log(com_tips)
            else:
                pass
        except Exception as e:
            err = f"Error: {e}"
            self.write(err)
            self.log.write_log(err)


class RecordLog:
    def __init__(self):
        super(RecordLog, self).__init__()

    @staticmethod
    def write_log(logs) -> None:
        with open("log.txt", "a+", encoding="utf-8") as fp:
            now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            re_logs = f"{now} -> {str(logs)}\n"
            fp.write(re_logs)
        return


if __name__ == '__main__':
    app = QApplication(sys.argv)

    qtmodern.styles.dark(app)
    mw = qtmodern.windows.ModernWindow(MainWindow())
    mw.setWindowTitle("DataAnalysis")
    mw.show()

    sys.exit(app.exec())
