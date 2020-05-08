from tkinter import Tk,Label,PhotoImage,Entry,Button,Frame

root=Tk()
root.geometry('500x550+100+100')
root.overrideredirect(True)
root.resizable(0,0)


img=PhotoImage(file='UI/Loading.gif')
label=Label(root,image=img)
label.pack(fill='both')

root.update()

from tkinter.filedialog import askopenfilename,askdirectory
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import numpy as np
import os
from threading import Thread

from reportlab.lib.pagesizes import A4,landscape
from reportlab.platypus import SimpleDocTemplate, LongTable
from reportlab.pdfgen import canvas

import shutil
from time import sleep


import warnings

warnings.filterwarnings('ignore')


OUTPUT_PATH="C:\\Users\\{0}\\Desktop\\Output\\".format(os.getlogin())



try:
    os.mkdir(OUTPUT_PATH)
except:
    pass



class Search:
    def f():
        pass

        

    def back():
        Search.Search_Frame.pack_forget()
        Upload.File_Upload_Frame.pack(fill='both')
        root.title('')


    Search_Frame=Frame(root,width=1200,height=650)
    img3=PhotoImage(file='UI/back.png')
    back_btn=Button(Search_Frame,image=img3,border=0,command=back)
    back_btn.place(x=20,y=20)


    Search_Frame=Frame(root,width=1200,height=650)
#    Label(View_Frame,image=img2).pack(fill='both')

    img3=PhotoImage(file='UI/back.png')
    back_btn=Button(Search_Frame,image=img3,border=0,command=back)
    back_btn.pack()
    
    tree = ttk.Treeview(Search_Frame, columns = (), height=25,show = "headings")
    tree.pack()

    scroll2=ttk.Scrollbar(Search_Frame,orient="horizontal",command=tree.xview)
    scroll2.pack(side='bottom',fill='x')
    tree.configure(xscrollcommand=scroll2.set)


    
class View:
    def f():
        pass
    
    def back():
        View.View_Frame.pack_forget()
        Upload.File_Upload_Frame.pack(fill='both')
        root.title('')
    
    View_Frame=Frame(root,width=1200,height=650)

    img3=PhotoImage(file='UI/back.png')
    back_btn=Button(View_Frame,image=img3,border=0,command=back)
    back_btn.pack()
    
    tree = ttk.Treeview(View_Frame, columns = (), height=25,show = "headings")
    tree.pack()

    scroll2=ttk.Scrollbar(View_Frame,orient="horizontal",command=tree.xview)
    scroll2.pack(side='bottom',fill='x')
    tree.configure(xscrollcommand=scroll2.set)



class Upload:
    def f():
        pass
        
    def dividend_report():
        try:
            Upload.dividend_status.config(text='Running')
            percentage=float(Upload.dividend_entry.get())
            data=pd.read_excel('Database/Master.xlsx')
            data2=pd.read_excel('Database/Report.xlsx')
            location=np.unique(data['Loc_Name'])
            df_dict={}
            for loc in location:
                df_dict[loc]=pd.DataFrame({'Emp_No':[],'Member_No':[],'Name':[],'Loc_Name':[],'Sh_Fund':[],'Dividend_Fund':[]})

            for i in range(data.shape[0]):
                emp_i=data['Emp_No'][i]
                for j in range(data2.shape[0]):
                    emp_j=data['Emp_No'][j]
                    if emp_i==emp_j:
                        tmp_data=df_dict[data['Loc_Name'][i]]
                        index=tmp_data.shape[0]
                        tmp_data.loc[index]=[emp_i,data['Member_No'][i],data['Name'][i],data['Loc_Name'][i],data2['Sh_Fund'][j],round(data2['Sh_Fund'][j]*percentage,2)]

            def temp_f(data,y):
                cols=data.columns
                data=data.values.tolist()
                data.insert(0,cols)


                output=OUTPUT_PATH+y+'.pdf'
                
                elements = []
                doc = SimpleDocTemplate(output, pagesize=landscape(A4))
                t = LongTable(data,repeatRows=1)
                elements.append(t)
                doc.build(elements)

            for i in df_dict:
                data=df_dict[i]
                name=i
                t1=Thread(target=temp_f,args=(data,name,))
                t1.start()
        except:
            Upload.dividend_status.config(text='Error')
            sleep(1)
        Upload.dividend_status.config(text='K')


    def New_Batch_mini():
        file=askopenfilename(filetypes=[('Excel',".xlsx")])
        Upload.new_batch_status.config(text='Running')

        try:
            new=pd.read_excel(file).values
            data_1=pd.read_excel('Database/Master.xlsx')
            shutil.copy('Database/Master.xlsx','Database/tmp_Master.xlsx')
            index_1=data_1.shape[0]
            data_2=pd.read_excel('Database/Report.xlsx')
            shutil.copy('Database/Report.xlsx','Database/tmp_Report.xlsx')
            index_2=data_2.shape[0]
            for i in new:
                data_1.loc[index_1]=i
                data_2.loc[index_2]=[i[0],i[1],i[2],0,0,0,0,0,0]
                index_1+=1
                index_2+=1
            data_1.sort_values('Emp_No')
    ##        print(data_1)
    ##        print(data_2)
            data_1.to_excel('Database/Master.xlsx',index=False)
            data_2.sort_values('Emp_No')
            data_2.to_excel('Database/Report.xlsx',index=False)
        except:
            Upload.new_batch_status.config(text='Failed')
            sleep(1)
            
        Upload.new_batch_status.config(text='K')

    def New_Batch():
        t1=Thread(target=Upload.New_Batch_mini)
        t1.start()
        
        
        
    
    def Retire_Batch_mini():
        file=askopenfilename(filetypes=[('Excel',".xlsx")])
        Upload.retire_batch_status.config(text='Running')
#        print(file)
        try:
            new=pd.read_excel(file).values
            data_1=pd.read_excel('Database/Master.xlsx')
            data_2=pd.read_excel('Database/Report.xlsx')
            shutil.copy('Database/Master.xlsx','Database/tmp_Master.xlsx')
            shutil.copy('Database/Report.xlsx','Database/tmp_Report.xlsx')
            for i in new:
                emp_no=i[0]
                for j in range(len(data_1['Emp_No'])):
                    if emp_no==data_1['Emp_No'][j]:
                        data_1['Status'][j]=2
                for k in range(len(data_2['Emp_No'])):
                    if emp_no==data_2['Emp_No'][k]:
                        data_2['T_Fund'][k]=round(data_2['T_Fund'][k]+data_2['Sh_Fund'][k],2)
                        data_2['Sh_Fund'][k]=0
            data_1.to_excel('Database/Master.xlsx',index=False)

            data_2.to_excel('Database/Report.xlsx',index=False)
        except:
            Upload.retire_batch_status.config(text='Failed')
            sleep(1)
        Upload.retire_batch_status.config(text="K")
        

    def Retire_Batch():
        t1=Thread(target=Upload.Retire_Batch_mini)
        t1.start()
        

    def Location_Batch_mini():
        file=askopenfilename(filetypes=[('Excel',".xlsx")])
#        print(file)
        Upload.location_batch_status.config(text='Running')
        try:
            new=pd.read_excel(file).values
            data_1=pd.read_excel('Database/Master.xlsx')

            shutil.copy('Database/Master.xlsx','Database/tmp_Master.xlsx')
            shutil.copy('Database/Report.xlsx','Database/tmp_Report.xlsx')
            
            for i in new:
                emp_no=i[0]
                for j in range(len(data_1['Emp_No'])):
                    if emp_no==data_1['Emp_No'][j]:
                        data_1['Loc'][j]=i[1]
                        data_1['Loc_Name'][j]=i[2]
            
            data_1.to_excel('Database/Master.xlsx',index=False)
        except:
            Upload.location_batch_status.config(text='Failed')
            sleep(1)
        Upload.location_batch_status.config(text='K')

    def Location_Batch():
        t1=Thread(target=Upload.Location_Batch_mini)
        t1.start()

    def Folder_Batch_mini():
        path=askdirectory()
        #print(path)
        #sleep(4)
        Upload.folder_upload_status.config(text='Running')
        try:
            file_list=os.listdir(path)
            time_chart={1:12,2:11,3:10,4:9,5:8,6:7,7:6,8:5,9:4,10:3,11:2,12:1}
            data=pd.read_excel('Database/Report.xlsx')

            shutil.copy('Database/Master.xlsx','Database/tmp_Master.xlsx')
            shutil.copy('Database/Report.xlsx','Database/tmp_Report.xlsx')


            data['Int_T_Fund']=round(data['Int_T_Fund']+data['T_Fund']*0.07,2)
            data['Int_G_Fund']=round(data['Int_G_Fund']+data['G_Fund']*0.06,2)
            #print(data)
            df_list=[]
            for file in file_list:
                time_period=time_chart[int(file.split('.')[0])]
                tmp_data=pd.read_excel(path+'/'+file)
                tmp_data['Int_T_Fund']=round(tmp_data['T_Fund']*0.07*time_period/12,2)
                tmp_data['Int_G_Fund']=round(tmp_data['G_Fund']*0.06*time_period/12,2)
                df_list.append(tmp_data)
                #print(1)
            for df in df_list:
                #print(2)
                for i in range(len(df['Emp_No'])):
                    emp_no_i=df['Emp_No'][i]
                    for j in range(len(data['Emp_No'])):
                        emp_no_j=data['Emp_No'][j]
                        if emp_no_i==emp_no_j:
                            data['Sh_Fund'][j]=round(data['Sh_Fund'][j]+df['Sh_Fund'][i],2)
                            data['T_Fund'][j]=round(data['T_Fund'][j]+df['T_Fund'][i],2)
                            data['Int_T_Fund'][j]=round(data['Int_T_Fund'][j]+df['Int_T_Fund'][i],2)
                            data['G_Fund'][j]=round(data['G_Fund'][j]+df['G_Fund'][i],2)
                            data['Int_G_Fund'][j]=round(data['Int_G_Fund'][j]+df['Int_G_Fund'][i],2)
            #print(5)
            #print(data)
            #os.remove('Database/Report.xlsx')
            data.to_excel('Database/Report.xlsx',index=False)
            #print('done')
                       

        except:
            Upload.folder_upload_status.config(text='Failed')
            sleep(1)
        Upload.folder_upload_status.config(text='K')
                        
            
    def Folder_Batch():
        t1=Thread(target=Upload.Folder_Batch_mini)
        t1.start()


    def logout():
        Upload.File_Upload_Frame.pack_forget()
        root.geometry('500x550+100+100')
        root.overrideredirect(True)
        Login.Login_Frame.pack(fill='both')


    def View():
        Upload.File_Upload_Frame.pack_forget()
        View.View_Frame.pack(fill='both')
        root.title('View')

        Data=pd.read_excel('Database/Report.xlsx')
        Columns=Data.columns
        tuple_count = tuple(Columns)

        View.tree.config(columns=tuple_count, height=25,show = "headings")
        View.tree.delete(*View.tree.get_children())
        
        List_comp=[]
        for i in range(len(Columns)):
            List_comp.append((Columns[i],i))

        for data, num in List_comp:
                if num==0 or num == 1 or num == 2:
                        width_tree = 200
                if num == 8:
                        width_tree = 300
                if num in(3,4,5,6,7):
                        width_tree = 130
                if num == 9:
                        width_tree = 140
                View.tree.heading(num, text = data)
                View.tree.column(num, width = width_tree)
        
        list_db = Data.values.tolist()

        for item in list_db:
                values_t = [] 
                k = 0
                for items in item:
                        values_t.insert(k,str(items))
                        k=k+1
                tuple_A = tuple(values_t)
                View.tree.insert('','end', values = tuple_A)



    def Search():
        Upload.File_Upload_Frame.pack_forget()
        Search.Search_Frame.pack(fill='both')
        root.title('Search')

        Data=pd.read_excel('Database/Master.xlsx')
        Columns=Data.columns
        tuple_count = tuple(Columns)

        Search.tree.config(columns=tuple_count, height=25,show = "headings")
        Search.tree.delete(*Search.tree.get_children())
        
        List_comp=[]
        for i in range(len(Columns)):
            List_comp.append((Columns[i],i))

        for data, num in List_comp:
                if num==0 or num == 1 or num == 2:
                        width_tree = 200
                if num == 8:
                        width_tree = 300
                if num in(3,4,5,6,7):
                        width_tree = 130
                if num == 9:
                        width_tree = 140
                Search.tree.heading(num, text = data)
                Search.tree.column(num, width = width_tree)
        
        list_db = Data.values.tolist()

        for item in list_db:
                values_t = [] 
                k = 0
                for items in item:
                        values_t.insert(k,str(items))
                        k=k+1
                tuple_A = tuple(values_t)
                Search.tree.insert('','end', values = tuple_A)


    def restore_mini():
        Upload.restore_status.config(text='Running')
        try:
            file_list=os.listdir('Database')
            if 'tmp_Report.xlsx' in file_list:
                os.remove('Database/Report.xlsx')
                os.rename('Database/tmp_Report.xlsx','Database/Report.xlsx')
            if 'tmp_Master.xlsx' in file_list:
                os.remove('Database/Master.xlsx')
                os.rename('Database/tmp_Master.xlsx','Database/Master.xlsx')
            try:
                os.remove('Report.pdf')
            except:
                pass
        except:
            Upload.restore_status.config(text='Failed')
            sleep(1)
        Upload.restore_status.config(text='K')

    def restore():
        t1=Thread(target=Upload.restore_mini)
        t1.start()

    def full_report_mini(x,y):
        try:
            data=pd.read_excel(x)
            cols=data.columns
            data=data.values.tolist()
            data.insert(0,cols)


            output=OUTPUT_PATH+y
            elements = []
            doc = SimpleDocTemplate(output, pagesize=landscape(A4))
            t = LongTable(data,repeatRows=1)
            elements.append(t)
            doc.build(elements)
        except:
            Upload.full_report_status.config(text='Error')
            sleep(1)
        Upload.full_report_status.config(text='K')
        
        
    def full_report():
        t1=Thread(target=Upload.full_report_mini,args=('Database/Report.xlsx','Report.pdf',))
        t2=Thread(target=Upload.full_report_mini,args=('Database/tmp_Report.xlsx','tmp_Report.pdf',))
        t1.start()
        t2.start()
        Upload.full_report_status.config(text='Running')
        
    def emp_report_mini():
        Upload.emp_report_status.config(text='Running')
        try:
            data1=pd.read_excel('Database/tmp_Report.xlsx')
            data2=pd.read_excel('Database/Report.xlsx')

            cols=data1.columns
            data1=data1.values.tolist()

            data2=data2.values.tolist()

            data=[]
            for i in range(len(data1)):
                data.append(cols)
                data.append(data1[i])
                data.append(data2[i])
                data.append([])

            output=OUTPUT_PATH+'Emp_Report.pdf'
            elements = []
            doc = SimpleDocTemplate(output, pagesize=landscape(A4))
            t = LongTable(data,repeatRows=1)
            elements.append(t)
            doc.build(elements)
        except:
            Upload.emp_report_status.config(text='Error')
            sleep(1)
        Upload.emp_report_status.config(text='K')



    def emp_report():
        t1=Thread(target=Upload.emp_report_mini)
        t1.start()


    def show_developer():
        if Upload.shown==0:
            Upload.developer_label.place(x=200,y=20)
            Upload.shown=1
        else:
            Upload.developer_label.place_forget()
            Upload.shown=0
        

    img2=PhotoImage(file='UI/2.gif')
    File_Upload_Frame=Frame(root,width=1200,height=650)
    Label(File_Upload_Frame,image=img2).pack(fill='both')


    

    search_btn=Button(File_Upload_Frame,text='Master',font=('',12),border=0,command=Search,width=10)
    view_btn=Button(File_Upload_Frame,text='View All',font=('',12),border=0,command=View,width=10)
    dividend_btn=Button(File_Upload_Frame,text='Dividend Report',font=('',12),border=0,command=dividend_report,width=14)
    dividend_status=Label(File_Upload_Frame,text='K')
    dividend_entry=Entry(File_Upload_Frame)
    
    img3=PhotoImage(file='UI/Logout.png')
    logout_btn=Button(File_Upload_Frame,image=img3,border=0,command=logout)

    file_1=Button(File_Upload_Frame,text='New Batch Emp',font=('',12),border=0,command=New_Batch,width=14)
    new_batch_status=Label(File_Upload_Frame,text='K')

    file_2=Button(File_Upload_Frame,text='Retire Batch Emp',font=('',12),border=0,command=Retire_Batch,width=14)
    retire_batch_status=Label(File_Upload_Frame,text='K')

    file_3=Button(File_Upload_Frame,text='Location',font=('',12),border=0,command=Location_Batch,width=14)
    location_batch_status=Label(File_Upload_Frame,text='K')

    file_4=Button(File_Upload_Frame,text='Folder Upload',font=('',12),border=0,command=Folder_Batch,width=14)
    folder_upload_status=Label(File_Upload_Frame,text='K')

    search_btn.place(x=50,y=120)
    view_btn.place(x=170,y=120)
    dividend_btn.place(x=290,y=120)
    logout_btn.place(x=1095,y=0)
    file_1.place(x=550,y=120)
    file_2.place(x=700,y=120)
    file_3.place(x=850,y=120)
    file_4.place(x=1000,y=120)

    dividend_status.place(x=290,y=160)
    dividend_entry.place(x=290,y=200)
    new_batch_status.place(x=550,y=160) 
    retire_batch_status.place(x=700,y=160)
    location_batch_status.place(x=850,y=160)
    folder_upload_status.place(x=1000,y=160)

    Button(File_Upload_Frame,text='Restore',font=('',12),border=0,command=restore,width=14).place(x=50,y=580)
    restore_status=Label(File_Upload_Frame,text='K')
    restore_status.place(x=50,y=620)

    Button(File_Upload_Frame,text='Full_Report',font=('',12),border=0,command=full_report,width=14).place(x=1050,y=480)
    full_report_status=Label(File_Upload_Frame,text='K')
    full_report_status.place(x=1050,y=520)
    
    Button(File_Upload_Frame,text='Emp_Report',font=('',12),border=0,command=emp_report,width=14).place(x=1050,y=580)
    emp_report_status=Label(File_Upload_Frame,text='K')
    emp_report_status.place(x=1050,y=620)


    Button(File_Upload_Frame,text='Developer',font=('',12),border=0,command=show_developer,width=14).place(x=20,y=20)
    developer_label=Label(File_Upload_Frame,text='TAMOJIT DAS\nIEM CSE 2 YEAR\n\nSHUBHOMOY CHAKRABARTI\nIEM BCA 3 YEAR',fg='blue')
    shown=0

    
class Login:

    def defId(event=None):
        Login.Id.delete(0,'end')
        Login.Password.config(show='*')
    def defPassword(event=None):
        Login.Password.delete(0,'end')
        Login.Password.config(show='*')
    def f():
        pass
    def Quit():
        root.destroy()
    def Login():
        if Login.Id.get()=='root' and Login.Password.get()=='root':
            Login.Login_Frame.pack_forget()
            root.geometry('1200x650+10+10')
            root.overrideredirect('False')
            root.iconbitmap('UI/download.ico')
            root.title('')
            Upload.File_Upload_Frame.pack(fill='both')


       
        Login.Id.delete(0,'end')            
        Login.Password.delete(0,'end')            
        Login.Id.insert(0,'Login ID')
        Login.Password.config(show='')
        Login.Password.insert(0,'Password')
        
    
    img=PhotoImage(file='UI/5.gif')
    Login_Frame=Frame(root,width=500,height=550,border=5)
    Label(Login_Frame,image=img).pack(fill='both')


    Id=Entry(Login_Frame,width=30,relief='groove',border=1,font=('',12,'italic'))
    Id.insert(0,'Login ID')
    Id.bind('<Button>',defId)

    Password=Entry(Login_Frame,width=30,relief='groove',border=1,font=('',12,'italic'))
    Password.insert(0,'Password')
    Password.bind('<Button>',defPassword)

    img2=PhotoImage(file='UI/Login.png')
    login_btn=Button(Login_Frame,image=img2,border=0,command=Login)
    img4=PhotoImage(file='UI/Rest.png')
    reset_btn=Button(Login_Frame,image=img4,border=0,command=f)
    img3=PhotoImage(file='UI/Exit.png')
    quit_btn=Button(Login_Frame,image=img3,border=0,command=Quit)

    Id.place(x=110,y=300)
    Password.place(x=110,y=340)


    login_btn.place(x=365,y=460)
    reset_btn.place(x=25,y=470)
    quit_btn.place(x=440,y=1)


label.pack_forget()


Login.Login_Frame.pack(fill='both')




root.mainloop()
