import os
from glob import glob
import pandas as pd
from pandas.core.frame import DataFrame



def rename(path, pattern, newname):
  files=glob(path+pattern)
  os.rename(files[0], path+newname)
#  files=glob('C:\\Users\\A.NAVARRO\\Downloads\\IMI_*')
#  os.rename(files[0], 'C:\\Users\\A.NAVARRO\\Downloads\\IMI.xlsx')

def read_Excel(path: str, fileName: str, sheetName: str)-> DataFrame:
    df=pd.read_excel(path+fileName, sheet_name=sheetName)

    return df
"""
def create_df(path, fileName, sheetName):
    Entt=pd.read_excel(path+fileName, sheet_name=sheetName)
    Dump=pd.read_excel(path+fileDump)
    main=pd.read_excel(path+fileIMI, sheet_name='Main')
     
    main[['ForeCast12M']]=main[['ForeCast12M']].astype(str)  
    main['ForeCast12M']=main['ForeCast12M'].str.replace(',', '.').replace('nan', '0')
    main=main.drop_duplicates() # Elimina todas las filas duplicadas osea las que estan en blanci al final conceros den Forecast12M
    main=main[:-1] # Como me queda una en la ultima linea que esta vacia elimino, cojo todas las lineas menos la ultima 
    ReportStatus=Dump.merge(Entt, how='left', left_on='Mánager', right_on='NameInApp')
    ReportStatus['Entity']=ReportStatus[['entidad']]
    ReportStatus=ReportStatus.drop(columns=['ManagerFullName', 'NameInApp', 'ManagerEmail', 'entidad'])
    return main, ReportStatus
 """   
def write():
    
    main, ReportStatus=create_df()
    writer = pd.ExcelWriter(fileout, engine='xlsxwriter')
    main.to_excel(writer, sheet_name='Sheet1', index=False)
    ReportStatus.to_excel(writer, sheet_name='Sheet2', index=False)
 #   writer.save() ---> Hago que la funcion devuelva el objero writer despues apico el metodo save()
    return writer

def remove():
  files=glob(path+'Reports_D*')
  os.remove(files[0])
  files=glob(path+'IMI_G*')
  os.remove(files[0])