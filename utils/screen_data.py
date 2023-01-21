import os

from qtpy.QtCore import QThread, Signal

from utils.record_logs import RecordLog


class ScreenDataThread(QThread):
    signal_trans = Signal(str)
    progressBarValue = Signal(int)
    signal_done = Signal(int)
    signal_desc = Signal(str)

    def __init__(self, file_name, file_out, targets, column_value):
        super(ScreenDataThread, self).__init__()
        self.log = RecordLog
        self.file_name = file_name
        self.file_out = file_out
        self.targets = targets
        self.column_value = column_value

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        self.signal_desc.emit('开始运行...')
        import pandas as pd
        try:
            file = os.path.splitext(self.file_name)
            if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                df = pd.read_excel(self.file_name, dtype=object)
            if file[1].lower() == '.csv':
                df = pd.read_csv(self.file_name, dtype=object)
            if file[1].lower() == '.tsv':
                df = pd.read_csv(self.file_name, dtype=object, sep="\t")
            ls_col = [df.columns[col] for col in range(len(df.columns))]
            read_tips = f"{os.path.basename(self.file_name)} read successfully."
            self.write(read_tips)
            self.log.write_log(read_tips)
            data_fr = pd.DataFrame(df, columns=ls_col)
            for tar in range(len(self.targets)):
                cd = {
                    self.column_value: self.targets[tar]
                }
                new_df = self.df_screen(data_fr, cd)
                try:
                    new_df.to_excel(f"{self.file_out}/{str(self.targets[tar])}.xlsx", encoding='gbk', index=False)
                    out = f"{str(tar + 1)} {str(self.targets[tar])}.xlsx is saved."
                    self.write(out)
                    self.log.write_log(out)
                    self.progressBarValue.emit(int(tar / len(self.targets) * 100))
                except:
                    new_df.to_csv(f"{self.file_out}/{str(self.targets[tar])}.csv", encoding="utf-8", index=False)
                    out = f"{str(tar + 1)} {str(self.targets[tar])}.csv is saved."
                    print(out)
                    self.write(out)
                    self.log.write_log(out)
                    self.progressBarValue.emit(int(tar / len(self.targets) * 100))
            self.signal_done.emit(1)
        except Exception as e:
            err = f"Error: {e}"
            self.write(err)
            self.log.write_log(err)
        else:
            com_tips = f"{len(self.targets)} files split."
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