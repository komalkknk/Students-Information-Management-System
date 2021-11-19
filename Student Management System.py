from tkinter import *
from tkinter import ttk
import pypyodbc
from tkinter import messagebox


class Student:
    def __init__(self,root):
        self.root=root
        self.root.title("Student Management System")
        self.root.geometry("1350x700+0+0")

        title=Label(self.root,text="Student Information Management System",bd=10,relief=GROOVE,font=("times new roman",40,"bold"),bg="grey",fg="#000080")
        title.pack(side=TOP,fill=X)

        #==========================All Variable======================================

        
        self.Roll_No_var=StringVar()
        self.Name_var=StringVar()
        self.Email_var=StringVar()
        self.Gender_var=StringVar()
        self.Contact_var=StringVar()
        self.DOB_var=StringVar()

        self.Search_by=StringVar()
        self.Search_txt=StringVar()
        
        

        #==========================Manage Frame======================================

        Manage_Frame=Frame(self.root,bd=4,relief=RIDGE,bg="grey")
        Manage_Frame.place(x=20,y=100,width=450,height=600)

        m_title=Label(Manage_Frame,text="Manage Students",bg="grey",fg="black",font=("times new roman",30,"bold"))
        m_title.grid(row=0,columnspan=2,pady=20)

        lbl_Roll=Label(Manage_Frame,text=" Roll No.",bg="grey",fg="black",font=("times new roman",18,"bold"))
        lbl_Roll.grid(row=1,column=0,pady=10,padx=20,sticky="w")

        txt_Roll=Entry(Manage_Frame,textvariable=self.Roll_No_var,font=("times new roman",13,"bold"),bd=5,relief=GROOVE)
        txt_Roll.grid(row=1,column=1,pady=10,padx=20,sticky="w")

        lbl_Name=Label(Manage_Frame,text="Name",bg="grey",fg="black",font=("times new roman",18,"bold"))
        lbl_Name.grid(row=2,column=0,pady=10,padx=20,sticky="w")

        txt_Name=Entry(Manage_Frame,textvariable=self.Name_var,font=("times new roman",13,"bold"),bd=5,relief=GROOVE)
        txt_Name.grid(row=2,column=1,pady=10,padx=20,sticky="w")

        lbl_Email=Label(Manage_Frame,text="Email",bg="grey",fg="black",font=("times new roman",18,"bold"))
        lbl_Email.grid(row=3,column=0,pady=10,padx=20,sticky="w")

        txt_Email=Entry(Manage_Frame,textvariable=self.Email_var,font=("times new roman",13,"bold"),bd=5,relief=GROOVE)
        txt_Email.grid(row=3,column=1,pady=10,padx=20,sticky="w")

        lbl_Gender=Label(Manage_Frame,text="Gender",bg="grey",fg="black",font=("times new roman",18,"bold"))
        lbl_Gender.grid(row=4,column=0,pady=10,padx=20,sticky="w")

        txt_Gender=Entry(Manage_Frame,textvariable=self.Gender_var,font=("times new roman",13,"bold"),bd=5,relief=GROOVE)
        txt_Gender.grid(row=4,column=1,pady=10,padx=20,sticky="w")

        combo_Gender=ttk.Combobox(Manage_Frame,font=("times new roman",12,"bold"),state='readonly')
        combo_Gender["values"]=("Select","Male","Female","Other")
        combo_Gender.grid(row=4,column=1,pady=10,padx=20)
        combo_Gender.current(0)

        lbl_Contact=Label(Manage_Frame,text="Contact",bg="grey",fg="black",font=("times new roman",18,"bold"))
        lbl_Contact.grid(row=5,column=0,pady=10,padx=20,sticky="w")

        txt_Contact=Entry(Manage_Frame,textvariable=self.Contact_var,font=("times new roman",13,"bold"),bd=5,relief=GROOVE)
        txt_Contact.grid(row=5,column=1,pady=10,padx=20,sticky="w")

        lbl_DOB=Label(Manage_Frame,text="D.O.B",bg="grey",fg="black",font=("times new roman",18,"bold"))
        lbl_DOB.grid(row=6,column=0,pady=10,padx=20,sticky="w")

        txt_DOB=Entry(Manage_Frame,textvariable=self.DOB_var,font=("times new roman",13,"bold"),bd=5,relief=GROOVE)
        txt_DOB.grid(row=6,column=1,pady=10,padx=20,sticky="w")

        lbl_Address=Label(Manage_Frame,text="Address",bg="grey",fg="black",font=("times new roman",18,"bold"))
        lbl_Address.grid(row=7,column=0,pady=10,padx=20,sticky="w") 

        self.txt_Address=Text(Manage_Frame,width=27,height=4,font=("",10))
        self.txt_Address.grid(row=7,column=1,pady=10,padx=20,sticky="w")

         #==========================Buttom Frame======================================
        
        btn_Frame=Frame(Manage_Frame,bd=4,relief=RIDGE,bg="grey")
        btn_Frame.place(x=15,y=510,width=380)

        addbtn=Button(btn_Frame,text="ADD",width=7,font="bold",command=self.add_students).grid(row=0,column=0,padx=10,pady=10)
        updatebtn=Button(btn_Frame,text="UPDATE",width=7,font="bold",command=self.update_data).grid(row=0,column=1,padx=10,pady=10)
        deletebtn=Button(btn_Frame,text="DELETE",width=7,font="bold",command=self.delete_data).grid(row=0,column=2,padx=10,pady=10)
        clearbtn=Button(btn_Frame,text="CLEAR",width=7,font="bold",command=self.clear).grid(row=0,column=3,padx=10,pady=10)
        
       
         #==========================Detail Frame======================================


        Detail_Frame=Frame(self.root,bd=4,relief=RIDGE,bg="grey")
        Detail_Frame.place(x=500,y=100,width=850,height=600)

        lbl_search=Label(Detail_Frame,text="Search By",bg="grey",fg="black",font=("times new roman",25,"bold"))
        lbl_search.grid(row=0,column=0,pady=25,padx=10,sticky="w")

        combo_search=ttk.Combobox(Detail_Frame,textvariable=self.Search_by,width=12,font=("times new roman",20,"bold"),state='readonly')
        combo_search["values"]=("Select"," Roll_No","Name","Contact")
        combo_search.grid(row=0,column=1,pady=7,padx=10)
        combo_search.current(0)

        txt_Search=Entry(Detail_Frame,textvariable=self.Search_txt,width=17,font=("times new roman",16,"bold"),bd=5,relief=GROOVE,)
        txt_Search.grid(row=0,column=2,pady=10,padx=10,sticky="w")

        searchbtn=Button(Detail_Frame,text="Search",width=10,pady=5,font="bold",command=self.search_data).grid(row=0,column=3,padx=10,pady=10)
        showallbtn=Button(Detail_Frame,text="Show All",width=10,pady=5,font="bold",command=self.fetch_data).grid(row=0,column=4,padx=10,pady=10)

        #==========================Table Frame======================================

        Table_Frame=Frame(Detail_Frame,bd=4,relief=RIDGE,bg="grey")
        Table_Frame.place(x=10,y=80,width=820,height=500)

        scroll_x = Scrollbar(Table_Frame, orient=HORIZONTAL)
        scroll_y = Scrollbar(Table_Frame, orient=VERTICAL)
        self.Student_table = ttk.Treeview(Table_Frame, columns=("Roll No", "Name", "Email", "Gender", "Contact", "DOB", "Address"),
                             xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)

        scroll_x.pack(side=BOTTOM,fill=X)
        scroll_y.pack(side=RIGHT,fill=Y)
        
        scroll_x.config(command=self.Student_table.xview)
        scroll_y.config(command=self.Student_table.yview)

        self.Student_table.heading("Roll No",text="Roll No")
        self.Student_table.heading("Name",text="Name")
        self.Student_table.heading("Email",text="Email")
        self.Student_table.heading("Gender",text="Gender")
        self.Student_table.heading("Contact",text="Contact")
        self.Student_table.heading("DOB",text="D.O.B")
        self.Student_table.heading("Address",text="Address")

        self.Student_table['show']='headings'

        self.Student_table.column("Roll No",width=70)
        self.Student_table.column("Name",width=120)
        self.Student_table.column("Email",width=120)
        self.Student_table.column("Gender",width=120)
        self.Student_table.column("Contact",width=120)
        self.Student_table.column("DOB",width=120)
        self.Student_table.column("Address",width=120)
        self.Student_table.bind("<ButtonRelease-1>",self.get_cursor)
        self.Student_table.pack(fill=BOTH,expand=1)
        self.fetch_data()
        

    def add_students(self):
        if self.Roll_No_var.get()=='' or self.Name_var.get()=='':
            messagebox.showerror('error','All fields are required')
        else:
            DRIVER='SQL Server'
            SERVER_NAME='DESKTOP-BR2ONM1'
            DATABASE_NAME='demo'

            conn_string=f"""
            Driver={{{DRIVER}}};
            Server={SERVER_NAME};
            Database={DATABASE_NAME};
            Trust_connection=yes;
            """     

            INSERT_STATEMENT="""insert into tblStudents values(?,?,?,?,?,?,?)"""
            RECORD=[
                    self.Roll_No_var.get(),
                    self.Name_var.get(),
                    self.Email_var.get(),
                    self.Gender_var.get(),
                    self.Contact_var.get(),
                    self.DOB_var.get(),
                    self.txt_Address.get("1.0",END)]
            conn=pypyodbc.connect(conn_string)
            cursor=conn.cursor()
            cursor.execute(INSERT_STATEMENT,RECORD)                                                                     
        
            cursor.commit()
            self.fetch_data()
            self.clear()
            cursor.close()
            conn.close()
            messagebox.showinfo('Success','Record has been inserted successfully')


    def fetch_data(self):
        DRIVER='SQL Server'
        SERVER_NAME='DESKTOP-BR2ONM1'
        DATABASE_NAME='demo'

        conn_string=f"""
        Driver={{{DRIVER}}};
        Server={SERVER_NAME};
        Database={DATABASE_NAME};
        Trust_connection=yes;
        """
        conn=pypyodbc.connect(conn_string)
        cursor=conn.cursor()
        cursor.execute("select * from tblStudents ")
        rows=cursor.fetchall()
        if len(rows)!=0:
            self.Student_table.delete(*self.Student_table.get_children())
            for row in rows:
                self.Student_table.insert('',END,values=row)
            cursor.commit()
        conn.close() 

    def clear(self):
        self.Roll_No_var.set(""),
        self.Name_var.set(""),
        self.Email_var.set(""),
        self.Gender_var.set(""),
        self.Contact_var.set(""),
        self.DOB_var.set(""),
        self.txt_Address.delete("1.0",END)

    def get_cursor(self,ev):
        cursor_row=self.Student_table.focus()
        contents=self.Student_table.item(cursor_row)
        row=contents['values']
        self.Roll_No_var.set(row[0]),
        self.Name_var.set(row[1]),
        self.Email_var.set(row[2]),
        self.Gender_var.set(row[3]),
        self.Contact_var.set(row[4]),
        self.DOB_var.set(row[5]),
        self.txt_Address.delete("1.0",END)
        self.txt_Address.insert(END,row[6])

    def update_data(self):
        DRIVER='SQL Server'
        SERVER_NAME='DESKTOP-BR2ONM1'
        DATABASE_NAME='demo'

        conn_string=f"""
        Driver={{{DRIVER}}};
        Server={SERVER_NAME};
        Database={DATABASE_NAME};
        Trust_connection=yes;
        """     

        INSERT_STATEMENT="""update tblStudents set Name=?,Email=?,Gender=?,Contact=?,DOB=?,Address=? where Roll_No=?"""
        RECORD=[                
                self.Name_var.get(),
                self.Email_var.get(),
                self.Gender_var.get(),
                self.Contact_var.get(),
                self.DOB_var.get(),
                self.txt_Address.get("1.0",END),
                self.Roll_No_var.get()]
        conn=pypyodbc.connect(conn_string)
        cursor=conn.cursor()
        cursor.execute(INSERT_STATEMENT,RECORD)                                                                     
        
        cursor.commit()
        self.fetch_data()
        self.clear()
        cursor.close()
        conn.close()

    def delete_data(self):
        DRIVER='SQL Server'
        SERVER_NAME='DESKTOP-BR2ONM1'
        DATABASE_NAME='demo'

        conn_string=f"""
        Driver={{{DRIVER}}};
        Server={SERVER_NAME};
        Database={DATABASE_NAME};
        Trust_connection=yes;
        """
        conn=pypyodbc.connect(conn_string)
        cursor=conn.cursor()
        INSERT_STATEMENT="""DELETE from tblStudents where Roll_No=?"""
        RECORD=[self.Roll_No_var.get()]
        cursor.execute(INSERT_STATEMENT,RECORD)
        cursor.commit()
        self.fetch_data()
        self.clear()
        cursor.close()
        conn.close()
            
     
    def search_data(self):
        DRIVER='SQL Server'
        SERVER_NAME='DESKTOP-BR2ONM1'
        DATABASE_NAME='demo'

        conn_string=f"""
        Driver={{{DRIVER}}};
        Server={SERVER_NAME};
        Database={DATABASE_NAME};
        Trust_connection=yes;
        """
        conn=pypyodbc.connect(conn_string)
        cursor=conn.cursor()
        cursor.execute("select * from tblStudents where "+str(self.Search_by.get())+" LIKE '%"+str(self.Search_txt.get())+"%'")
        rows=cursor.fetchall()
        if len(rows)!=0:
            self.Student_table.delete(*self.Student_table.get_children())
            for row in rows:
                self.Student_table.insert('',END,values=row)
            cursor.commit()
        conn.close() 
        

    #==========================Login Page Frame======================================

        


             
root=Tk()
ob=Student(root)
root.mainloop()       
    
