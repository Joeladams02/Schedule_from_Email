import pandas as pd
from datetime import datetime
import warnings
22
warnings.filterwarnings("ignore", message="Unknown extension is not supported and will be removed")
warnings.filterwarnings("ignore", message="Print area cannot be set")


class ReadExcel:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.months_array = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

        self.df_full = pd.read_excel(excel_file, sheet_name=self.months_array)
        self.dates = {}

    def check_if_scheduled(self, item, month_df, row_index, column_index, month_name):
        try:
            if type(item) == datetime and item.month == (self.months_array.index(month_name)+1):
                scheduled_lesson = month_df[row_index + 1][column_index] 
                if type(scheduled_lesson) != float:
                    return scheduled_lesson
        except IndexError:
            return None

    def process_Excel(self):
        for month_name in self.months_array:
            month_df = self.df_full[month_name].to_numpy()

            for row_index, row in enumerate(month_df):
                for column_index, item in enumerate(row):
                    scheduled_lesson = self.check_if_scheduled(item, month_df, row_index, column_index, month_name)
                    if scheduled_lesson is not None:
                        self.dates[item] = scheduled_lesson

    def get_scheduled_dates(self):
        return {key: value for key, value in self.dates.items() if len(value) >= 5}
