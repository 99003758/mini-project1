import pandas as pd

import sys
import matplotlib.pyplot as plt

df2 = pd.DataFrame()
for a in range(int(input(" Enter the no of inputs you required:"))):
    first = (input("Enter the p.s. number:"))
    second = (input("Enter the name:"))
    third = input("Enter the email:")
    class1 = pd.read_excel('final sheet211.xlsx', sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'])
    g = class1['Sheet1']
    g = g[g['p.s.no'] == first]
    g = g[g['NAME'] == second]
    g = g[g['email'] == third]
    df = pd.DataFrame(g)

    if df.empty:
        sys.exit("ENTERED DETAILS ARE INCORRECT")
    for i in class1.keys():

        sheet = class1[i]
        g = sheet[sheet['p.s.no'] == first]

        g = g[g['NAME'] == second]
        g = g[g['email'] == third]
        col = sheet.columns

        for j in col:

            df[j] = g[j]

            if a == 0:
                df2[j] = g[j]
    if a != 0:
        df2 = pd.concat([df, df2])
        print(df2)
k = ('final sheet.xlsx')
df2.to_excel('final sheet.xlsx', sheet_name='Sheet6')
print(df2)
df2.plot.bar(x='NAME', y='laptops')
plt.show()
