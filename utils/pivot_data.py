import os

from qtpy.QtCore import QThread, Signal

from utils.record_logs import RecordLog


class PivotDataThread(QThread):
    signal_trans = Signal(str)
    signal_trans_cols = Signal(str)
    signal_desc = Signal(str)

    def __init__(self, tableWidget, file_name,  idx, values, file_out):
        super(PivotDataThread, self).__init__()
        self.tableWidget = tableWidget
        self.file_name = file_name
        self.index = idx
        self.values = values
        self.file_out = file_out

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

            df[self.values] = df[self.values].astype(float)

            pt = pd.pivot_table(df, index=self.index, values=self.values, aggfunc=np.sum)
            pt.to_excel(f"{self.file_out}.xlsx", encoding="gbk")

        except Exception as e:
            err = f"Error: {e}"
            self.write(err)
        else:
            success_tips = f"{os.path.basename(self.file_name)} pivot succeeded"
            self.signal_desc.emit(success_tips)
