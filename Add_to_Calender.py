import pandas as pd
import numpy as np
import warnings
from datetime import datetime
warnings.filterwarnings("ignore", message="Unknown extension is not supported and will be removed")
warnings.filterwarnings("ignore", message="Print area cannot be set")

months = ['January', 'February', 'March','April','June','July','August','September','October','November','December']
df = pd.read_excel('Schedule_2025.xlsx', sheet_name = months)
Jan = df['January'].to_numpy()

dates = {}
for row_index, row in enumerate(Jan):
    for column_index,item in enumerate(row):
        if type(item) == datetime and item.month == 1:
            associated = Jan[row_index+1][column_index]
            if type(associated) != float:
                dates[item] = Jan[row_index+1][column_index]

print(dates)