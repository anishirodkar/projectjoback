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
text_file=open("GROUP SYMBOLS.txt","r") #will read the text file data
lines=text_file.readlines() # this is the array that has been formd after reading the file

n=int(raw_input("Please enter the number of groups in the compound>> "))
T=float(raw_input("Please enter the temperature >> "))
i=0
list_of_symbols=[] #empty list of the symbols, which will eventually be updated by the user
list_of_frequencies=[] #same as above, here for the group frequencies.
while(i<n):

    symbol,freq=input("Enter the group symbol and the group frequency ").split()
    symbol,freq=[symbol,int(freq)]
    list_of_symbols.append(symbol)
    list_of_frequencies.append(freq)
    i=i+1 #here, user keeps getting the above message 'n' times. he enters the details
s=0
j=0 
aa=0
bb=0
cc=0
dd=0 
#above are some parameters, which will be calculated below. 
#calculation of enthalpy  
for j in list_of_symbols:
    if(list_of_symbols[j]==lines[j]):
        H=freq[j]*grc[j]
        s=s+H
        aa = aa + freq[j] * a[j]
        bb = bb + freq[j] * b[j]
        cc = cc + freq[j] * c[j]
        dd = dd + freq[j] * d[j] 
        # above are the formulae of Jabock. 
    else:
        j=j+1
s1=s+68.29  # some correction factor
aaa = aa - 37.93
bbb = bb + 0.21
ccc = cc - 3.91 * 0.0001
ddd = dd + 2.06 * 0.0000001  

#below, we calculate the enthalpy of formation of the compound at another temperature.
hh = s1 + (aaa * (T - 298) + bbb * (T ^ 2 - 298 * 298) / 2 + ccc * (T ^ 3 - 298 ^ 3) / 3 + ddd * (T ^ 4 - 298 ^ 4) / 4) / 1000 

