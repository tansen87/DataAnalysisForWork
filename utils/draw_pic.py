import os

from qtpy.QtCore import QThread, Signal


# from utils.record_logs import RecordLog


class DrawPicture(QThread):
    signal_trans = Signal(str)
    signal_desc = Signal(str)

    def __init__(self, file, save_path, x, title, y=None):
        super(DrawPicture, self).__init__()
        self.file_name = file
        self.save_path = save_path
        self.x = x
        self.title = title
        self.y = y

    def write(self, text):
        self.signal_trans.emit(str(text))

    def bubble_sort(self, lst):
        cnt = len(lst)
        for i in range(cnt):
            for j in range(i + 1, cnt):
                if lst[i] > lst[j]:
                    lst[i], lst[j] = lst[j], lst[i]
        return lst

    def run(self):
        try:
            import pandas as pd
            import pyecharts.options as opts
            from pyecharts.charts import Bar

            draw_tips = f"Start drawing..."
            self.signal_desc.emit(draw_tips)

            file = os.path.splitext(self.file_name)
            if file[1].lower() == ".xlsx" or file[1].lower() == ".xls":
                df = pd.read_excel(self.file_name, dtype=object)
            if file[1].lower() == '.csv':
                df = pd.read_csv(self.file_name, dtype=object)

            df["len_col"] = df[self.x].str.len()
            df["len_col"] = df["len_col"].fillna(0)

            tmp = df["len_col"].value_counts().index.tolist()
            tmp_sort = self.bubble_sort(tmp)

            lst_x = [x for x in tmp_sort]
            lst_y = [df["len_col"].value_counts()[y] for y in tmp_sort]
            lst_y = list(map(int, lst_y))

            c = (
                Bar(
                    # 初始化配置项
                    init_opts=opts.InitOpts(
                        # 设置动画
                        # animation_opts=opts.AnimationOpts(
                        #     animation_delay=1000, animation_easing="elasticOut"),
                        # 设置宽度、高度
                        width='900px',
                        height='500px',
                        bg_color="white",)
                )
                .add_xaxis(lst_x)
                .add_yaxis("count", lst_y)
                .set_global_opts(title_opts=opts.TitleOpts(title=self.title),
                                 toolbox_opts=opts.ToolboxOpts(is_show=True))
                .render(f"{self.save_path}.html")
            )
            save_tips = f"{self.save_path}.html"
            self.write(save_tips)
            complete = f"Drawing completed"
            self.signal_desc.emit(complete)
        except Exception as e:
            err = f"Error: {e}"
            self.write(err)
