import pandas as pd 
from tkinter import *
from tkinter import filedialog 
from tkinter import messagebox as msg 
from pandastable import Table 
from tkintertable import TableCanvas
import pandas as pd
import numpy as np
import xlsxwriter
   
  
class attendance_generator: 
   
    def __init__(self, root): 
   
        self.root = root 
        self.file_name = ''
        self.writer=''
        self.a=''
        self.b=''
        self.c=''
        self.d=''
        self.file=''
        self.file_name2=''
        self.date=''
        self.f = Frame(self.root,height = 50,width = 70) 
          
        # Place the frame on root window 
        self.f.pack() 
           
        # Creating label widgets 
        self.message_label = Label(self.f, 
                                   text = '\nHindu Senior Secondary School\n', 
                                   font = ('Arial', 24), 
                                   fg = 'Blue') 
        self.message_label2 = Label(self.f, 
                                    text = 'Online Class Attendance Register Generator', 
                                    font = ('Arial', 18,'underline'), 
                                    fg = 'Red')
        self.message_label3 = Label(self.f, 
                                    text = 'Utility for Teachers\n', 
                                    font = ('Arial', 15,'underline'), 
                                    fg = 'Red')
        self.message_label4 = Label(self.f, 
                                    text = 'Step 1 - Select Student Master File (Output):', 
                                    font = ('Arial', 12,'underline'), 
                                    fg = 'Black')
        self.message_label5 = Label(self.f, 
                                    text = 'Step 2 - Enter the Date (DD/MM/YY):', 
                                    font = ('Arial', 12,'underline'), 
                                    fg = 'Black')
        self.message_label6 = Label(self.f, 
                                    text = 'Step 3 - Select the Input File (Attendance List):', 
                                    font = ('Arial', 12,'underline'), 
                                    fg = 'Black')
        self.message_label7 = Label(self.f, 
                                    text = 'Step 4 - Process & Generate Attendance Register:', 
                                    font = ('Arial', 12,'underline'), 
                                    fg = 'Black')
        self.message_label8 = Label(self.f, 
                                    text = '\n\n© Copyright 2020, Tejasri Nagavati Bogu', 
                                    font = ('Arial', 9), 
                                    fg = 'Black')
        
        
        self.txt = Entry(self.f,textvariable = self.date,font=('calibre',10,'normal'))
   
        # Buttons 
        self.mainfile_button = Button(self.f, 
                                     text = 'Master File', 
                                     font = ('Arial', 12), 
                                     bg = 'White', 
                                     fg = 'Black', 
                                     command = self.open_mainfile) 
        self.list_button = Button(self.f, 
                                    text = 'Attendance File', 
                                    font = ('Arial', 12),  
                                    bg = 'White', 
                                    fg = 'Black', 
                                    command = self.open_list) 
        self.confirm_button = Button(self.f, 
                                    text = 'Process', 
                                    font = ('Arial', 12), 
                                    bg = 'Yellow', 
                                    fg = 'Black',  
                                    command = self.calculate_att)
        self.date_button = Button(self.f, 
                                  text = 'Confirm', 
                                  font = ('Arial', 12), 
                                  bg = 'white',
                                  fg = 'Black',  
                                  command = self.date_read)
        self.exit_button = Button(self.f, 
                                  text = 'Exit', 
                                  font = ('Arial', 12), 
                                  bg = 'Red', 
                                  fg = 'Black',  
                                  command = root.destroy)
        self.clear_button = Button(self.f, 
                                  text = 'Clear', 
                                  font = ('Arial', 12), 
                                  bg = 'Red', 
                                  fg = 'Black',  
                                  command = root.destroy)
    # Placing the widgets using grid manager 
        self.message_label.grid(row=0,column=0, columnspan=3) 
        self.message_label2.grid(row = 1,column=0, columnspan=3)
        self.message_label3.grid(row =2,column=0, columnspan=3)
        self.message_label4.grid(row = 3, column = 1,sticky=W)
        self.mainfile_button.grid(row = 3, column = 2,sticky=W) 
        self.message_label5.grid(row = 4, column = 1,sticky=W)
        #self.txt.grid(row=5,column=1,columnspan=3)    
        #self.date_button.grid(row = 4, column = 2,sticky=W) 
        
        self.txt.grid(row=4,column=2,sticky=W)    
        self.date_button.grid(row = 4, column = 2,sticky=E)
        
        self.message_label6.grid(row = 6, column = 1, sticky = W)
        self.list_button.grid(row = 6, column = 2,sticky=W)
        self.message_label7.grid(row = 7, column = 1,sticky=W)
        self.confirm_button.grid(row = 7, column = 2,sticky=W)
        self.exit_button.grid(row = 8, column = 1, pady = 40,columnspan=3)
        #self.clear_button.grid(row = 8, column = 1,pady=5,sticky=E)
        self.message_label8.grid(row = 9, column = 1,columnspan=3) 
        
   
    def open_mainfile(self): 
        try: 
            self.file_name = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select an Excel file', 
                                                        filetypes = (('excel file','*.xlsx'), 
                                                                     ('excel file','*.xlsx'))) 
               
            #df = pd.read_csv(self.file_name)
            self.writer = pd.ExcelWriter(self.file_name, engine='xlsxwriter',mode='w+')
            pd.set_option('mode.chained_assignment','raise')

            self.a=pd.read_excel(self.file_name,sheet_name='IX-A',index_col=0)
            
            self.b=pd.read_excel(self.file_name,sheet_name='IX-B',index_col=0)
            
            self.c=pd.read_excel(self.file_name,sheet_name='IX-C',index_col=0)
            
            self.d=pd.read_excel(self.file_name,sheet_name='IX-D',index_col=0)
            
            msg.showinfo('Success', 'File selected Successfully')           
              
            # Next - Pandas DF to Excel file on disk 
            #if(len(self.writer) == 0):       
             #   msg.showinfo('No Rows Selected', 'Excel has no rows') 
           # else: 
                #mfile=input("Enter the file name where the attendance should be stored :  ")
                #mfile=mfile+".xlsx"
                #self.writer = pd.ExcelWriter(mfile, engine='xlsxwriter',mode='w+')
                    
               
        except FileNotFoundError as e: 
                msg.showerror('Error in opening file', e) 
        
   
    def open_list(self):
        try:
            self.file_name2 = filedialog.askopenfilename(initialdir = '/Desktop', 
                                                        title = 'Select an Excel file', 
                                                        filetypes = (('excel file','*.xlsx'), 
                                                                     ('excel file','*.xlsx'))) 
               
            #df = pd.read_csv(self.file_name)
            self.file=pd.read_excel(self.file_name2)
            msg.showinfo('Success', 'File selected Successfully')
            
            #k=self.file.empty()
            #if(k=='True'):
               # print("ERROR in reading File")
               
            #else:
                 #print("Read File")  
            # Next - Pandas DF to Excel file on disk 
            #if(len(self.writer) == 0):       
             #   msg.showinfo('No Rows Selected', 'Excel has no rows') 
           # else: 
                #mfile=input("Enter the file name where the attendance should be stored :  ")
                #mfile=mfile+".xlsx"
                #self.writer = pd.ExcelWriter(mfile, engine='xlsxwriter',mode='w+')
        except FileNotFoundError as e: 
                msg.showerror('Error in opening file', e) 
    def display_xls_file(self): 
        try: 
            self.file_name = filedialog.askopenfilename(initialdir = '/', 
                                                        title = 'Select a excel file', 
                                                        filetypes = (('excel file','*.xlsx'), 
                                                                     ('excel file','*.xlsx'))) 
            df = pd.read_excel(self.file_name) 
            msg.showinfo('Success', 'File selected Successfully')
                            
            # Now display the DF in 'Table' object 
            # under'pandastable' module 
            self.f2 = Frame(self.root, height=200, width=300)  
            self.f2.pack(fill=BOTH,expand=1) 
            self.table = Table(self.f2, dataframe=df,read_only=True) 
            self.table.show()
        except FileNotFoundError as e: 
            print(e) 
            msg.showerror('Error in opening file',e) 
      
    def date_read(self):
        self.date=self.txt.get()
        if(len(self.date)==8):
            msg.showinfo('Success', 'Proper Date has been entered')
        else:
            msg.showinfo('ERROR', 'Empty or incorrect date.\n Please enter Date in the format DD/MM/YY')
            
    def calculate_att(self):   
        #function to insert the day in ascending order
        
        def pos(p,date,k):
            temp=list(p.columns.values.tolist())
            date_chk=list(date.split("/"))
            for i in range(1,len(temp)):
                u=list(temp[i].split("/"))
                #print(date_chk)
                if(date_chk[0]<u[0] and date_chk[1]<=u[1] and date_chk[2]<=u[2]):
                    p.insert(i,date,k)
                    return 0
                elif(date_chk[0]==u[0] and date_chk[1]==u[1] and date_chk[2]==u[2]):
                    MsgBox = messagebox.askyesno ('Error','Attendance exists for the given date.\nAre you sure you want to over write.',icon = 'warning')
                    #print(MsgBox)
                    if (MsgBox == True):
                        p.drop([date], axis = 1, inplace = True) 
                        p.insert(i,date,k)
                        return 0
                    elif(MsgBox== False):
                        msg.showinfo('ERROR', 'Empty or incorrect date.\n Please enter Date in the format DD/MM/YY')
                    return 0
            p.insert(len(p.columns),date,k)
            
            
        #function to mark attendance based on the file uploaded
        def mark(p,q,date):
            k = np.empty(p.shape[0], dtype = str)
            for i in range(1,p.shape[0]+1):
                #temp=0
                for j in range(0,q.shape[0]):
                    #if(k[i-1]!=""):
                        #temp+=1
                    if((p['NAME'][i].upper()).replace(" ","")==q['NAME'][j].upper() and (q['ATTENDANCE'][j]=='Joined' or q['ATTENDANCE'][j]=='Joined before')):
                        k[i-1]='P'
                        #k[i-1]='P'+" "+str(temp)
            
                    elif((p['NAME'][i].upper()).replace(" ","")==q['NAME'][j].upper() and q['ATTENDANCE'][j]=='Left'):
                        k[i-1]='A'
                        #k[i-1]='A'+" "+str(temp)
            pos(p,date,k)
             
    

        mark(self.a,self.file,self.date)
        mark(self.b,self.file,self.date)
        mark(self.c,self.file,self.date)
        mark(self.d,self.file,self.date)


        self.a.to_excel(self.writer,sheet_name='IX-A')
        
        self.b.to_excel(self.writer,sheet_name='IX-B')
        
        self.c.to_excel(self.writer,sheet_name='IX-C')
        
        self.d.to_excel(self.writer,sheet_name='IX-D')
        
        self.writer.save()
        
        msg.showinfo('Success', 'Student Attendance Register Updated Successfully!')
  
# Driver Code  
root = Tk() 
root.title('Attendance Compiler') 

obj = attendance_generator(root) 
root.geometry('700x500')
root.resizable(width=False, height=False)
root.mainloop() 
