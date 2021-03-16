import pandas as pd
from funcion import rename

path='C:\\Users\\A.NAVARRO\\Downloads\\'
imi="IMI_General_21.xlsx"
newname='Reports_Dump.xlsx'
pattern='Reports_D*'
wdir='C:/Users/A.NAVARRO/OneDrive - o2f-it/UiPath/RoboMails/Scripts'

def main(month, year):

    mesMasUno=str(int(month)+1)

    month=str(month)
    year=str(year) 
    
    monthYear=month+'/'+year
    
    rename(path, pattern, newname)
    df=pd.read_excel(path+imi, sheetName='Main')[['Email', 'ManagerEmail', 'Bonificados']]
    
    rep=pd.read_excel(path+newname)

    rep[['month','year']] = rep['Fecha del informe'].apply(lambda x: pd.Series(str(x).split("/")))
    
    rep=rep[((rep.Estado=='Pendiente') | (rep.Estado=='Rechazado'))&(rep.year==year)]

    countPtes=rep[rep.month!=mesMasUno].groupby('Correo del consultor').agg({'Id':'count'})

    ptes=rep[rep['Fecha del informe']==monthYear]
    
    new = ptes["Consultor"].str.split(" ", n = 1, expand = True) 
    new=new[[0]].rename(columns={0:'FName'})
    ptes=pd.concat([new,ptes], axis=1)

    mrg=ptes.merge(df, how='left', left_on='Correo del consultor', right_on='Email')\
        [['FName','Correo del consultor', 'ManagerEmail', 'Fecha del informe', 'Estado', 'Bonificados']]

    mrg=mrg[mrg['Bonificados'].isin(['Si', 'si'])]

    mrg=mrg.merge(countPtes, how='left', on='Correo del consultor')
    
    mrg=mrg.rename(columns={'Correo del consultor':'consultor', 'ManagerEmail':'manager', 'Fecha del informe':'mes/año', 'Estado':'estado', 'Id': 'backlog'})
    
    mrg_soloUno=mrg[mrg.backlog==1]
    mrg_masdeUno=mrg[mrg.backlog!=1]

    mrg_soloUno.to_csv(wdir+'/data/mails.csv', index=False, header=True)
    mrg_masdeUno.to_csv(wdir+'/data/mails_nanual.csv', index=False, header=True)


if __name__=="__main__":

    main(2, 2021)
