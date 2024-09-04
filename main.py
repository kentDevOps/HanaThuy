import pandas as pd
import numpy as np
import glob
from openpyxl import load_workbook

file_paths = glob.glob('./*BOM*.xlsx')
#df = [pd.read_excel(file) for file in file_paths]
df = pd.read_excel(file_paths[0])
result = df.groupby(['Mã sản phẩm', 'Mã NPL'], as_index=False).agg({'Lượng NL, VT thực tế sử dụng để sản xuất một sản phẩm ': 'sum'})
print(result)
colNpl = result.iloc[:,0]
colSoLuong = result.iloc[:,2]
fileReport = 'vd.xlsx'
dfReport = pd.read_excel(fileReport,sheet_name='vd')
lastRowRp = len(dfReport) + 1
print(dfReport)
dfReport.iloc[:,:] = np.nan
print(dfReport)
# Ghi vào một trang tính cụ thể 
with pd.ExcelWriter(r'D:\02_Study\08_Python\16_Hana_Report\HanaThuy\vd.xlsx' , engine='openpyxl', mode='a',if_sheet_exists='overlay') as writer:
    dfReport.to_excel(writer, sheet_name='vd', index=False, startcol= 0,startrow=22, header=False) 
    colNpl.to_excel(writer, sheet_name='vd', index=False, startcol= 1,startrow=22, header=False)  
    colSoLuong.to_excel(writer, sheet_name='vd', index=False, startcol= 5,startrow=22, header=False)  