import os

from qtpy.QtCore import QThread, Signal, QModelIndex
from qtpy.QtGui import QStandardItemModel, QStandardItem

from utils.record_logs import RecordLog


class ViewDataThread(QThread):
    """ 查看excel数据 """
    signal_trans = Signal(str)
    signal_trans_cols = Signal(str)
    signal_desc = Signal(str)

    def __init__(self, tableWidget, file_name):
        super(ViewDataThread, self).__init__()
        self.tableWidget = tableWidget
        self.log = RecordLog
        self.file_name = file_name

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        import pandas as pd
        import numpy as np

        try:
            load_tips = f"Reading {os.path.basename(self.file_name)}..."
            self.signal_desc.emit(load_tips)
            file = os.path.splitext(self.file_name)
            if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                df = pd.read_excel(self.file_name, dtype=object)
            if file[1].lower() == '.csv':
                df = pd.read_csv(self.file_name, dtype=object)
            read_tips = f"File name: {os.path.basename(self.file_name)}"
            self.write(read_tips)

            rows = df.shape[0]
            cols = df.shape[1]

            col_names = [col for col in df.columns]

            # 将list转换为str
            str_col = ",".join(col_names)
            self.signal_trans_cols.emit(str_col)

            # 对excel行数进行判断, if行数大于3w行, 行数设置为3w行; 否则以实际为准
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
            excel_detail = f"rows: {rows}, cols: {cols}."
            self.write(excel_detail)
            self.log.write_log(excel_detail)
            success_tips = f"{os.path.basename(self.file_name)} read successfully"
            self.signal_desc.emit(success_tips)
