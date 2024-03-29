"""
Program to Mark attendence from the list in Microsoft Teams
@author: Tejasri
"""

import pandas as pd
import numpy as np
 
mfile=input("Enter the file name where the attendence should be stored :  ")
mfile=mfile+".xlsx"
writer = pd.ExcelWriter(mfile, engine='xlsxwriter',mode='w+')

pd.set_option('mode.chained_assignment','raise')

a=pd.read_excel(mfile,sheet_name='IX-A',index_col=0)
b=pd.read_excel(mfile,sheet_name='IX-B',index_col=0)
c=pd.read_excel(mfile,sheet_name='IX-C',index_col=0)
d=pd.read_excel(mfile,sheet_name='IX-D',index_col=0)
x=input("Enter the attendance file name:  ")
x=x+'.xlsx'
date=input("Enter the date in DD/MM/YY format:  ")

file=pd.read_excel(x)

#function to insert the day in ascending order
def pos(p,date,k):
    temp=list(p.columns.values.tolist())
    date_chk=list(date.split("/"))
    for i in range(1,len(temp)):
        u=list(temp[i].split("/"))
        if(date_chk[0]<u[0] and date_chk[1]<=u[1] and date_chk[2]<=u[2]):
            p.insert(i,date,k)
            return
        elif(date_chk[0]==u[0] and date_chk[1]==u[1] and date_chk[2]==u[2]):
            print("Attendence for %s already exisits.\nIf you wish to overwrite please enter 'Yes', to terminate enter 'No'",date)
            t=input()
            if(t=="Yes"):
                p.replace(to_replace = p[temp[i]], value =k)
            elif(t=='No'):
                print("Terminated")
            else:
                print("Unrecognized Error")
            return
    p.insert(len(p.columns),date,k)
            
            
    #function to mark attendence based on the file uploaded
def mark(p,q,date):
    k = np.empty(p.shape[0], dtype = str)
    for i in range(1,p.shape[0]+1):
        #temp=0
        for j in range(0,q.shape[0]):
            #if(k[i-1]!=""):
                #temp+=1
            if((p['NAME'][i].upper()).replace(" ","")==q['NAME'][j].upper() and (q['ATTENDENCE'][j]=='Joined' or q['ATTENDENCE'][j]=='Joined before')):
                k[i-1]='P'
                #k[i-1]='P'+" "+str(temp)
            
            elif((p['NAME'][i].upper()).replace(" ","")==q['NAME'][j].upper() and q['ATTENDENCE'][j]=='Left'):
                k[i-1]='A'
                #k[i-1]='A'+" "+str(temp)
    pos(p,date,k)
             
    

mark(a,file,date)
mark(b,file,date)
mark(c,file,date)
mark(d,file,date)


a.to_excel(writer,sheet_name='IX-A')
b.to_excel(writer,sheet_name='IX-B')
c.to_excel(writer,sheet_name='IX-C')
d.to_excel(writer,sheet_name='IX-D')
writer.save()

