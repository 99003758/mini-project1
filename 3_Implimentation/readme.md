
import pandas as pd
from openpyxl import load_workbook
import sys
import matplotlib.pyplot as plt
df = pd.DataFrame()
for i in range(int(input())):
    h= int(input("give ps"))
    a= input("name")
    v=input("mail")
    z= pd.read_excel('pythonexcel1.xlsx',sheet_name=['Sheet1','Sheet2','Sheet3','Sheet4','Sheet5'])
    df2= pd.DataFrame()
    for j in z.keys():
        x= z[j]
        t= x[x['PS number']==h]
        t= x[x['Display Name']==a]
        t= x[x['Official Email Address']==v]
        col = x.columns
        for k in col:
            df2[k]=t[k]
            if i==0:
                df[k]=t[k]


        if i !=0:
            df=pd.concat([df,df2])
            print(df2)


df.to_excel('pythonexcel.xlsx',sheet_name= 'master',index = False)
print(df)
df.plot.bar(x='Phone no',y='BUS NUM')
plt.show()





