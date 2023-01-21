import os
import datetime

from qtpy.QtCore import QThread, Signal
from qtpy.QtGui import QStandardItemModel, QStandardItem

from utils.record_logs import RecordLog


class MergeDataThread(QThread):
    """ VLookup """
    signal_trans = Signal(str)
    signal_trans_cols = Signal(str)
    signal_trans_cols_map = Signal(str)
    signal_desc = Signal(str)

    def __init__(self, table, init_table, mapping_table, left, right):
        super(MergeDataThread, self).__init__()
        self.log = RecordLog
        self.tableWidget = table
        self.init_table = init_table
        self.mapping_table = mapping_table
        self.left = left
        self.right = right

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        import pandas as pd
        import numpy as np

        try:
            try:
                df_init = pd.read_excel(self.init_table, dtype=object)
            except:
                df_init = pd.read_csv(self.init_table, dtype=object)
            read_tips = f"File name: {os.path.basename(self.init_table)}"
            self.write(read_tips)

            col_init = [df_init.columns[col] for col in range(len(df_init.columns))]
            str_col_init = ",".join(col_init)
            self.signal_trans_cols.emit(str_col_init)
            self.write(f"{os.path.basename(self.init_table)} read successfully.")
        except Exception as e:
            err = f"Error: {e}"
            self.write(err)
        else:
            try:
                file = os.path.splitext(self.mapping_table)
                if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                    df_map = pd.read_excel(self.mapping_table, dtype=object)
                if file[1].lower() == '.csv':
                    df_map = pd.read_csv(self.mapping_table, dtype=object)
                read_tips = f"File name: {os.path.basename(self.mapping_table)}"
                self.write(read_tips)

                col_map = [df_map.columns[col] for col in range(len(df_map.columns))]
                str_col_map = ",".join(col_map)
                self.signal_trans_cols_map.emit(str_col_map)
                self.write(f"{os.path.basename(self.mapping_table)} read successfully.")

                df_merge = pd.merge(df_init, df_map, left_on=self.left, right_on=self.right, how='left')

                rows = df_merge.shape[0]
                cols = df_merge.shape[1]

                excel_detail = f"rows: {rows}, cols: {cols}."
                self.write(excel_detail)
                self.log.write_log(excel_detail)

                col_names = [df_merge.columns[col] for col in range(len(df_merge.columns))]
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
