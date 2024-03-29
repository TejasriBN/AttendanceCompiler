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
        self.f = Frame(self.root,height = 200,width = 300) 
          
        # Place the frame on root window 
        self.f.pack() 
           
        # Creating label widgets 
        self.message_label = Label(self.f, 
                                   text = 'Hindu Senior Secondary School', 
                                   font = ('Arial', 20,'underline'), 
                                   fg = 'Blue') 
        self.message_label2 = Label(self.f, 
                                    text = 'Online Attendance Compiler', 
                                    font = ('Arial', 15,'underline'), 
                                    fg = 'Red')
        self.message_label5 = Label(self.f, 
                                    text = 'Step 3: Enter the Date (DD/MM/YY):', 
                                    font = ('Arial', 15,'underline'), 
                                    fg = 'Black')
        self.message_label3 = Label(self.f, 
                                    text = 'Step 1: Select the Input File (Attendance List)', 
                                    font = ('Arial', 15,'underline'), 
                                    fg = 'Black')
        self.message_label4 = Label(self.f, 
                                    text = 'Step 2: Select the Output File (Master File)', 
                                    font = ('Arial', 15,'underline'), 
                                    fg = 'Black')
        self.message_label6 = Label(self.f, 
                                    text = 'Step 4: Proceed to update', 
                                    font = ('Arial', 15,'underline'), 
                                    fg = 'Black')
        self.message_label7 = Label(self.f, 
                                    text = 'Press Exit to terminate', 
                                    font = ('Arial', 15,'underline'), 
                                    fg = 'Black')
        
        self.txt = Entry(self.f,textvariable = self.date,font=('calibre',10,'normal'))
   
        # Buttons 
        self.mainfile_button = Button(self.f, 
                                     text = 'File', 
                                     font = ('Arial', 14), 
                                     bg = 'White', 
                                     fg = 'Black', 
                                     command = self.open_mainfile) 
        self.list_button = Button(self.f, 
                                     text = 'File', 
                                     font = ('Arial', 14),  
                                     bg = 'White', 
                                     fg = 'Black', 
                                     command = self.open_list) 
        self.confirm_button = Button(self.f, 
                                  text = 'Process', 
                                  font = ('Arial', 14), 
                                  bg = 'Yellow', 
                                  fg = 'Black',  
                                  command = self.calculate_att)
        self.date_button = Button(self.f, 
                                  text = 'Confirm', 
                                  font = ('Arial', 10), 
                                  bg = 'white', 
                                  fg = 'Black',  
                                  command = self.date_read)
        self.exit_button = Button(self.f, 
                                  text = 'Exit', 
                                  font = ('Arial', 14), 
                                  bg = 'Red', 
                                  fg = 'Black',  
                                  command = root.destroy) 
   
        # Placing the widgets using grid manager 
        self.message_label.grid(row = 1, column = 1) 
        self.message_label2.grid(row = 2, column = 1)
        self.message_label3.grid(row = 4, column = 1)
        self.list_button.grid(row = 5, column = 1,  
                                 padx = 10, pady = 15)
        self.message_label4.grid(row = 8, column = 1)
        self.mainfile_button.grid(row = 9, column = 1, 
                                 padx = 0, pady = 15) 
        self.txt.grid(row=11,column=1)
        self.date_button.grid(row = 11, column = 2,  
                                 padx = 2, pady = 2)
        self.message_label5.grid(row = 10, column = 1)
        self.message_label6.grid(row = 12, column = 1)
        self.confirm_button.grid(row = 13, column = 1,  
                                 padx = 10, pady = 15)
        self.message_label7.grid(row = 14, column = 1)
        self.exit_button.grid(row = 15, column = 1, 
                              padx = 10, pady = 15) 
   
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
            
            msg.showinfo('Success', 'File Successfully Selected')           
              
            # Next - Pandas DF to Excel file on disk 
            #if(len(self.writer) == 0):       
             #   msg.showinfo('No Rows Selected', 'Excel has no rows') 
           # else: 
                #mfile=input("Enter the file name where the attendence should be stored :  ")
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
            msg.showinfo('Success', 'File Successfully Selected')
            
            #k=self.file.empty()
            #if(k=='True'):
               # print("ERROR in reading File")
               
            #else:
                 #print("Read File")  
            # Next - Pandas DF to Excel file on disk 
            #if(len(self.writer) == 0):       
             #   msg.showinfo('No Rows Selected', 'Excel has no rows') 
           # else: 
                #mfile=input("Enter the file name where the attendence should be stored :  ")
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
            msg.showinfo('Success', 'File Successfully Selected')
                            
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
            msg.showinfo('Success', 'Date Read')
        else:
            msg.showinfo('ERROR', 'Could not read the date.\n Enter Date in the form DD/MM/YY')
            
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
                    return
                elif(date_chk[0]==u[0] and date_chk[1]==u[1] and date_chk[2]==u[2]):
                    print("Attendence for %s already exisits.\nIf you wish to overwrite please enter 'Yes', to terminate enter 'No'",self.date)
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
             
    

        mark(self.a,self.file,self.date)
        mark(self.b,self.file,self.date)
        mark(self.c,self.file,self.date)
        mark(self.d,self.file,self.date)


        self.a.to_excel(self.writer,sheet_name='IX-A')
        
        self.b.to_excel(self.writer,sheet_name='IX-B')
        
        self.c.to_excel(self.writer,sheet_name='IX-C')
        
        self.d.to_excel(self.writer,sheet_name='IX-D')
        
        self.writer.save()
        
        msg.showinfo('Success', 'Successfully Updated')
        
            
            
            
  
# Driver Code  
root = Tk() 
root.title('Attendance Compiler') 

obj = attendance_generator(root) 
root.geometry('800x600') 
root.mainloop() 
