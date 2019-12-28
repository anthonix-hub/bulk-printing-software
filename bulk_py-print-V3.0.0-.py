import os
import time
import tkinter as tk
from datetime import date
from tkinter import *
from tkinter import filedialog, ttk
from tkinter.messagebox import *
from tkinter.ttk import Frame, LabelFrame, OptionMenu

import win32com
from PIL import Image, ImageTk
from win32com import client
import win32print 


root = Tk()


        #**************************************** splash screen ********************************

class SplashScreen(Frame):
    def __init__(self, master=None, width=0.6, height=0.4, useFactor=True):
        Frame.__init__(self, master)
        self.pack(side=TOP, fill=BOTH, expand=YES)

        # get screen width and height
        ws = self.master.winfo_screenwidth()
        hs = self.master.winfo_screenheight()
        w = (useFactor and ws*width) or width
        h = (useFactor and ws*height) or height
        # calculate position x, y 
        x = (ws/2) - (w/2) 
        y = (hs/2) - (h/2)
        self.master.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        self.master.overrideredirect(True)
        self.lift() 

def splash():
    if __name__ == '__main__':
        origin = Tk()

        sp = SplashScreen(origin)
        # sp.config(bg="red")

        m = Label(sp, text="NABTEB")
        m.pack(side=TOP, expand=YES)
        m.config(bg="#3366ff", justify=CENTER, font=("calibri", 95))

        # p = ImageTk.PhotoImage(Image.open('about.png'))
        # ph =  Button(sp,image=p,compound='top',text='about',height=45,width=50,bg='#fff').pack(side=TOP, expand=YES)
        
        # Button(sp, text="Press this button to kill the program", bg='red', command=origin.destroy).pack(side=BOTTOM, fill=X)

        sp.after(150,origin.destroy)
        
        # MSG = Label(origin,text='message after splash screen').pack()
        

        origin.mainloop()
# splash()


root.iconbitmap(r'about - Copy.ico')
# root.iconphoto(default='True')
root.title('Bulk printer V3.0.0-py')
root.configure(background='#dcebf1')
root.minsize(width=1040,height=700)
root.maxsize(width=1020,height=600)

# root.configure(bg='#dfd')
        #*************** Variables used **********************
  
dir_path = os.getcwd()
print(dir_path)
nam = 'Bulk-printed_files'
curr_date = date.today()
dty = str(curr_date)
folder_path = os.path.join(str(dir_path),str(nam))


        #*************** functions **************
def check_folder_exist():

    if os.path.exists(folder_path):
        print('folder exists')
        chg_dir = os.chdir(folder_path)
        sub_dir = os.path.join(str(folder_path),dty)

        if os.path.exists(str(sub_dir)):
            print('sub folder exists')
            root.nw_path = os.path.join(str(os.getcwd()),str(sub_dir))
            print(os.getcwd())
        else:
            sub_dir = os.mkdir(dty)
        

    else:
        dir_d = os.mkdir(nam)
        print('does not exit')
        
        chg_dir = os.chdir(folder_path)
        sub_dir = os.mkdir(dty)

        root.nw_path = os.path.join(str(os.getcwd()),str(sub_dir))

    os.chdir(str(dir_path))        

check_folder_exist()                     #***calls function to create a folder for recording names of the printed files 

def folder_opener():
        
    root.filename = filedialog.askdirectory(title='folder Finder',initialdir=os.path.dirname('desktop'))

    info_label = Label(printer_frame,text=root.filename,bg='#dda')
    info_label.place(x=2,y=30)
    info_label.configure(text=root.filename)

#     Label.configure(JOB_frame,text=root.filename)

    return root.filename

def file_create():
        file_path = os.path.join(str(root.filename)+'.txt') 
        txt_var = os.path.basename(str(file_path))
        try:
                nw_file_path = os.path.join(str(root.nw_path),str(txt_var))
                tk.f =  open(nw_file_path,'a')
        except:
                showinfo(detail='Sorry something went wrong :(\n\npleace exit and re-enter program !!!')
                root.title('bulk printer')
        # return tk.f
 
def Dir_scan():
    f = file_create()
    var_choice_mnu = choice_mnu.get()

    dir_name = root.filename[root.filename.find('/Users'):]
        
    with os.scandir(dir_name) as it:
        for  entry in it:
          if entry.is_file():
              files = entry.name
              file_arr = [files]
               
              for x in file_arr:   
                   print(x)
                   df = 'printed -- '+ x +'\n'
                   txt_area.insert(0.0,df)
                   print(x,file=tk.f)
                   
                   #********************* doc & printer control********************************
                   for copies in range(int(var_choice_mnu)):
                        
                        word = win32com.client.Dispatch('Word.Application')
                        time.sleep(0.1)
                        dir_name = root.filename[root.filename.find('/Users'):]
                        word.Documents.Open(os.path.join( str(dir_name),x))
                        word.ActiveDocument.PrintOut()
                        time.sleep(0.1)
                        word.visible = 0
                        word.ActiveDocument.Close()
                   
                   #********************************************************************************

                   root.update() #returns control to the program while doing job

                   disp = 'Bulk printing :'+ x
                   root.title(disp)
                  
    showinfo(detail='  Printing Completed!!!')
    root.title('bulk printer')

print(os.getcwd())
def file_select():
        var_choice_mnu = choice_mnu.get()

        root.select = filedialog.askopenfilenames()
        root.in_select = list(root.select)
        print(root.in_select)
        
        for x in root.in_select:
                print(str(x))
                txt_area.insert(0.0,x +'\n')
                filr_name = x[x.find('/Users'):]
                
                #******************* files & printer control ************************************
                for copies in range(int(var_choice_mnu)):
                        word = win32com.client.Dispatch('Word.Application')
                        time.sleep(0.1)
                        word.Documents.Open(filr_name)
                        word.ActiveDocument.PrintOut()
                        time.sleep(0.3)
                        word.visible = 0
                        word.ActiveDocument.Close()
                
                #*********************************************************************************

                # txt_area.insert(0.0,filr_name +'\n')
                root.update()
        showinfo(detail='  Printing Completed!!!')
        return filr_name

def Dir_scan2(e):
        Dir_scan()

def clear():
    txt_area.delete(1.0,END)

def exit():
   ask = askquestion(title='Quit',message='Do you want to close program?')
   if ask == 'yes':
        root.destroy()
        print('application exited')
        
def exit2(e):
    exit()

        #********************** menus ************************
def help_menu():
   showinfo(title='help for bulk printer ',type='ok',message='HELP',detail="*** OPEN FOLDER *** Use the Open folder Button to navigate and select a folder,"
             "which is container for the the job you which to print.\n\n"
             "*** PRINT *** After using the open folder,click on the print button.The print button will take the desired job to the printer for printing.\n\n "
             "*** PRINTER INFO *** this program works with the system's default printer,when a printer is changed, do well to set at as the default printer in the system's settings.\n\n"
             """*** clear field *** this can be used to clear the printed files displayed on the text area. \n\n"""
             """*** copies *** you can select the desired number of copies you wish to print, this can be done by clicking on the up or down arrow on the spinbox on the copies option menu or the up or down key on the keyboard,you can also enter the number from the keyboard. \n\n""" 
             "**** QUIT *** Contrl + q can be used to terminate the program, or go to the the exit menu above the program, and click 'YES' on the dialog section.\n\n"
             "*** PROGRAM error *** If the program encounter any error or malfunctions,please quit and restart the program. \n\n"
             """When the program is running, any miscrosoft word document opened will be forced to close. So when using this program do not use miscrosoft word, because the word doc will close""")
def about():
    
    showinfo(title='About bulk printer \n version 3.0.0 ',type='ok',message='About',detail="This program is designed for the purpose of helping in the printing of parking Lists,center statistics and other "
            "form of operations required in the smooth running of the examinations. It's major function is to speed up the process of printing any required documents of very "
             "large quantities of files of '.doc','.pdf','.xmls' exetensions in leser time and less efforts and even reduse the number of staffs it may take to do such work. \n\n"
             "This is a ground breaking version  and hope to improve on it's fuctionalities, as staffs get to use the program and suggestions may arise as to what might be desired to be included."
             " That will lead to factory recall for improvement and adjustments of feastures thank you.\n\n "
            "********************************************************************************************************************************************************"
             "\t\t  created by --ANTHONY EKOH-- \n\nfor the National Business and Technical Examinations Board (NABTEB)\n"
             "******************************************************************************************************************************************************\n\n"
             "Thanks to the staffs and management of ICT department(NABTEB)\n\n my fellow IT guys :\n JUDE ONOHWOSAFA, AUSTINE OGBEIDE ,OSAKUE GODSWILL, WISDOM ADAMS and OSADEBAWMEN OKOYO. \n\n"
             "I dedicate my success to my lovely mum ***** Mrs JUSTINA EKOH ***** \n\n"
            "\t\t\t\t\tAlrights reserved " )

# menu = Menu(root)
# root.config(menu = menu)

# subMenu = Menu(menu,tearoff=False)
# menu.add_cascade(label="File",menu=subMenu )
# subMenu.add_command(label="new Ctr+N")
# subMenu.add_separator()
# subMenu.add_command(label="Exit Ctr+Q",command=exit)
 
# editMenu = Menu(menu,tearoff=False)
# menu.add_cascade(label="Edit",menu=editMenu)
# editMenu.add_command(label="redo Ctr+Z")

# optionsMenu = Menu(menu,tearoff=False)
# menu.add_cascade(label='options',menu=optionsMenu)
# optionsMenu.add_command(label='print',command='crt+p',underline=0)

# helpmenu = Menu(menu,tearoff=False)
# menu.add_cascade(label='Help',menu=helpmenu)
# helpmenu.add_command(label="About",underline=0,command=about)
# helpmenu.add_separator()
# helpmenu.add_command(label='Help',command=help_menu)

    
        #********************** Frames ***************************

s = ttk.Style()
s.configure('blue.TLabelframe.Label',text='erer',background='#5cd9e2')

main_frame = Frame(root,relief=RIDGE,borderwidth=5)
main_frame.grid(padx=30,pady=66,ipady=0)

display_frame = Frame(main_frame,borderwidth=30,relief=RAISED,width=250)
display_frame.grid(pady=25,padx=33)

printer_frame = ttk.LabelFrame(main_frame, borderwidth=20,style='blue.TLabelframe.Label')
printer_frame.place(x=43,y=0)

JOB_frame = LabelFrame(display_frame,text="",borderwidth=30,style="blue.TLabelframe.Label")
JOB_frame.grid(column=0,padx=2,pady=15)

file_JOB_frame = LabelFrame(display_frame,text="",borderwidth=25,style="blue.TLabelframe.Label")
file_JOB_frame.grid(row=0,column=2,padx=1,pady=0)

buttom_frame = Frame(main_frame)
buttom_frame.grid(sticky='e',pady=1)

        #****************** options menu section *******************************
help_photo =ImageTk.PhotoImage(Image.open('help.png'))

option_menu = Button(root,image= help_photo,cursor='hand2',overrelief=GROOVE,text='help',width=50,height=45,compound='top',command=help_menu,bg='#fff')
option_menu.place(x=45,y=10)

about_photo =ImageTk.PhotoImage(Image.open('about.png'))

about_menu = Button(root,image=about_photo,cursor='hand2',overrelief=GROOVE,compound='top',text='about',height=45,width=50,command=about,bg='#fff')
about_menu.place(x=110,y=10)

close_butt = Button(root,height=2,width=7,cursor='hand2',overrelief=GROOVE,font='bold 12',bg='#ec365d',text='exit',command=exit)
close_butt.place(x=340,y=10)

clear_butt = Button(root,height=2,cursor='hand2',overrelief=GROOVE,width=9,bd=2,text='clear field',command=clear,relief=RAISED)
clear_butt.place(x=176,y=10)

choice_mnu = Spinbox(root,from_=1,to=50,width=7,insertbackground="red",wrap='0')
choice_mnu.place(x=270,y=40)

choice_label = Label(root,text='copies',height=1,width=6,font='italic 10',fg='purple',compound='top')
choice_label.place(x=270,y=10)

       #********************* printer section *************************

vok = win32print.GetDefaultPrinterW()
vok2 = win32print.OpenPrinter(vok)
default_ptr = "Detected printer: " + vok

prter_dlog = Label(printer_frame,fg='#afd',bg='grey',width=53,text=default_ptr,borderwidth=7,font='varient 10')
prter_dlog.grid(row=0,column=0,padx=1)

                        
        #****************** job printer section *********************

folder_photo =ImageTk.PhotoImage(Image.open('folder.png'))

dir_butt = Button(JOB_frame,image=folder_photo,cursor='hand2',overrelief=GROOVE,height=85,width=90,fg='#000',font='10',compound='top',text='open folder',command=folder_opener,relief=RAISED)
dir_butt.grid(row=2,column=0,rowspan=2,pady=13,padx=0)

file_photo =ImageTk.PhotoImage(Image.open('file.png'))

file_butt = Button(file_JOB_frame,image=file_photo,cursor='hand2',overrelief=GROOVE,height=90,width=80,fg='#000',font='7',compound='top',text='select file(s)',command=file_select,relief=RAISED)
file_butt.grid(row=4,column=0,pady=10)

photo = ImageTk.PhotoImage(file="printer.png")

print_butt = Button(JOB_frame,image=photo,height=75,cursor='hand2',overrelief=GROOVE,width=80,fg='red',bd=4,compound='top',text='print',command=Dir_scan,relief=RAISED)
print_butt.grid(row=3,column=1,columnspan=2,padx=12,pady=15,sticky='e')


txt_area = Text(main_frame,undo=True,bg='#fff',fg="#000",height="30",width='43',bd=1,wrap=WORD)
txt_area.grid(sticky='e',column=2,row=0,columnspan=2,padx=30,ipady=25)


        #********************* event Binbings **************
root.bind("<Control-q>",exit2)
root.bind("<Control-p>",Dir_scan2)




info = Label(root,text='(C) by Anthony Ekoh 2018',font='variant 10')
info.grid(sticky='e')

root.mainloop()