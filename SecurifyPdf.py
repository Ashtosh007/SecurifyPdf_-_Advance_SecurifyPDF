from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfWriter
from PyPDF2 import  PdfReader 
import os

root = Tk()
root.title("Pdf Protector")
root.geometry("600x430")
root.resizable(False,False)


# Button , place where from we have to find the pdf .
def browse():
    global filename
    filename=filedialog.askopenfilename(initialdir=os.getcwd(),title="Select Image File",filetypes=(('PDF File', '.pdf'),('all files','.*')))
    entry1.insert(END,filename)



# button for Protected.
def protect():
    mainfile=source.get()
    protectfile=target.get()
    code=password.get()




    if mainfile=="" and protectfile=="" and code=="":
        messagebox.showerror("invalid","enter all the detail") 

    elif mainfile=="":
        messagebox.showerror("invalid" ," enter source file ")

    
    elif protectfile=="":
        messagebox.showerror("invalid ", "enter target file name ")

        
    elif password.get()=="":
        messagebox.showerror("invalid" , "enter password ")
    

    else:
        try:
            out = PdfWriter()
            file = PdfReader(filename)
            num = len(file.pages)
            
            

            for i in range(num):
                page=file.pages[i]
                out.add_page(page)

            #password
            out.encrypt(code)

            with open(protectfile, "wb" ) as f:
                out.write(f)
            
            
            messagebox.showinfo("info","sucessfully done")
        
        except:
            messagebox.showerror("Invalid","Invalid entry")


#main


Label(root,text="SecurifyPdf",font="ariel 30 bold").place(x=75,y=35)


frame=Frame(root,width=580, height=290,bd=5,relief=GROOVE,bg="LIGHT BLUE")
frame.place(x=10,y=130)


########
source=StringVar()
Label(frame,text="SOURCE PDF file:", font ="arial 10 bold", fg="black").place(x=30,y=50)
entry1=Entry(frame,width=25,textvariable=source,font="arial 15",bd=1)
entry1.place(x=150,y=48)


########
select_button=Button(frame,text="BROWSE FILE",bg="Light Pink",font="ariel 12 bold",command=browse).place(x=440,y=45)

#####
target=StringVar()
Label(frame,text="TARGET PDF file:", font ="arial 10 bold", fg="black").place(x=30,y=100)
entry2=Entry(frame,width=25,textvariable=target,font="arial 15",bd=1)
entry2.place(x=150,y=100)


##########
password=StringVar()
Label(frame,text="SET PASSWORD:", font ="arial 10 bold", fg="BLACK").place(x=15,y=150)
entry3=Entry(frame,width=25,textvariable=password,font="arial 15",bd=1)
entry3.place(x=150,y=150)



#######
Protect=Button(root,text="PROTECT PDF ",compound=LEFT,width=20, height=0,bg="Light Pink",font="arial 14 bold" , command=protect)
Protect.pack(side=BOTTOM,pady=40)


root.mainloop()
