import datetime


class RecordLog:
    def __init__(self):
        super(RecordLog, self).__init__()

    @staticmethod
    def write_log(logs) -> None:
        with open("examples/logs.txt", "a+", encoding="utf-8") as fp:
            now_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            re_logs = f"{now_time} -> {str(logs)}\n"
            fp.write(re_logs)
        return
