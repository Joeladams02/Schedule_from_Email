import pandas as pd
import warnings
from datetime import datetime
warnings.filterwarnings("ignore", message="Unknown extension is not supported and will be removed")
warnings.filterwarnings("ignore", message="Print area cannot be set")

def check_if_scheduled(item, months):
    try:
        if type(item) == datetime and item.month == (months_array.index(months)+1):
            associated = month_df[row_index+1][column_index]
            if type(associated) != float:
                return month_df[row_index+1][column_index]
    except IndexError:
        return None


months_array = ['January', 'February', 'March','April','May','June','July','August','September','October','November','December']
df_full = pd.read_excel('Schedule_2025.xlsx', sheet_name = months_array)


dates = {}


for months in months_array:
    print(months)

    month_df = df_full[months].to_numpy()
    for row_index, row in enumerate(month_df):
        for column_index,item in enumerate(row):
            check = check_if_scheduled(item, months)
            if check != None:
                dates[item] = check
