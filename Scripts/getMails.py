import pandas as pd
from funcion import rename, read_Excel

path='C:\\Users\\A.NAVARRO\\Downloads\\'
imi="IMI_General_21.xlsx"
newname='Reports_Dump.xlsx'
pattern='Reports_D*'

def main():
    rename(path, pattern, newname)
    df=pd.read_excel(path+imi, sheetName='Main')[['Email', 'ManagerEmail']]
    rep=pd.read_excel(path+newname)
    rep=rep[(rep.Estado=='Pendiente') | (rep.Estado=='Rechazado')]
    ptes=rep[rep['Fecha del informe']=='1/2021']
    mrg=ptes.merge(df, how='left', left_on='Correo del consultor', right_on='Email')
    mrg=mrg[['Correo del consultor', 'ManagerEmail', 'Fecha del informe', 'Estado']]
    data=mrg.rename(columns={'Correo del consultor':'consultor', 'ManagerEmail':'manager', 'Fecha del informe':'mes', 'Estado':'estado'})
    data.to_csv('data/mails.csv', index=False, header=True)


if __name__=="__main__":

    main()