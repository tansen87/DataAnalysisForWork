import os

from qtpy.QtCore import QThread, Signal

from utils.record_logs import RecordLog


class ConcatDataThread(QThread):
    """ 合并一个文件夹内的所有csv和excel文件 """
    signal_trans = Signal(str)

    def __init__(self, file_out, file_out_mer):
        super(ConcatDataThread, self).__init__()
        self.log = RecordLog
        self.file_out = file_out
        self.file_out_mer = file_out_mer

    def write(self, text):
        self.signal_trans.emit(str(text))

    def run(self):
        import pandas as pd
        try:
            lst_file = list()
            init_path = f"{self.file_out}/"
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
            all_file = pd.concat(lst_file, axis=0, ignore_index=True, sort=False)
            try:
                all_file.to_excel(f"{self.file_out_mer}.xlsx", index=False, encoding="gbk")
                com_tips = f"file save path: {self.file_out_mer}."
                self.write(com_tips)
                self.log.write_log(com_tips)
            except:
                all_file.to_csv(f"{self.file_out_mer}.csv", index=False, encoding="utf-8")
                com_tips = f"file save path: {self.file_out_mer}."
                self.write(com_tips)
                self.log.write_log(com_tips)
        except Exception as e:
            err = f"Error: {e}"
            self.write(err)
            self.log.write_log(err)
