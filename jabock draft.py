# -*- coding: utf-8 -*-


import win32com.client
import scipy

xl=win32com.client.gencache.EnsureDispatch("Excel.Application")
wb=xl.Workbooks('jabock data actual.xlsx')
sheet=wb.Sheets('datasheet')
def getdata(sheet, Range):
    data=sheet.Range(Range).Value
    data=scipy.array(data)
    #data=data.reshape((1,len(data)))[0]
    return data
    
grc=getdata(sheet, "C2:C42")  #this will form an array for all the group contributions
a=getdata(sheet, "D2:D42")  #('a'the values of constants of specific heat variation with temp)
b=getdata(sheet, "E2:E42")  #('b'the values of constants of specific heat variation with temp)
c=getdata(sheet, "F2:F42")  #('c'the values of constants of specific heat variation with temp)
d=getdata(sheet, "G2:G42")  #('d'the values of constants of specific heat variation with temp)            
text_file=open("GROUP SYMBOLS.txt","r")
lines=text_file.readlines()

n=int(raw_input("Please enter the number of groups in the compound>> "))
T=float(raw_input("Please enter the temperature >> "))
i=0
list_of_symbols=[]
list_of_frequencies=[]
while(i<n):

    symbol,freq=input("Enter the group symbol and the group frequency ").split()
    symbol,freq=[symbol,int(freq)]
    list_of_symbols.append(symbol)
    list_of_frequencies.append(freq)
    i=i+1
s=0
j=0 
aa=0
bb=0
cc=0
dd=0   
for j in list_of_symbols:
    if(list_of_symbols[j]==lines[j]):
        H=freq[j]*grc[j]
        s=s+H
        aa = aa + freq[j] * a[j]
        bb = bb + freq[j] * b[j]
        cc = cc + freq[j] * c[j]
        dd = dd + freq[j] * d[j]
    else:
        j=j+1
s1=s+68.29  
aaa = aa - 37.93
bbb = bb + 0.21
ccc = cc - 3.91 * 10 ^ -4
ddd = dd + 2.06 * 10 ^ -7  
hh = s1 + (aaa * (T - 298) + bbb * (T ^ 2 - 298 ^ 2) / 2 + ccc * (T ^ 3 - 298 ^ 3) / 3 + ddd * (T ^ 4 - 298 ^ 4) / 4) / 1000    
        
        
    
    


    
        
    
    