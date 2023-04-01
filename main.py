import sys
from tkinter import Tk
from tkinter import ttk
from tkinter import *
import functions


def close_stop_process():
    sys.exit()


def add_kwor_remove(event):
    var_new_keyword=var_entry.get()
    print(var_new_keyword,functions.search_var,sep=';')
    if var_new_keyword in functions.search_var:
        functions.search_var.remove(var_new_keyword)
        var_txt.set(functions.search_var)
        lbl.update()
    else:
        functions.search_var.append(var_new_keyword)
        var_txt.set(functions.search_var)
        lbl.update()
    print(var_new_keyword,functions.search_var,sep=';')


root=Tk()
root.title('scraper')

frm_txt=ttk.Frame(master=root)
frm_txt.grid(row=0,column=0,rowspan=2)
frm_btn=ttk.Frame(master=root)
frm_btn.grid(row=2,column=0,rowspan=2)

var_txt=StringVar()
var_txt.set(functions.search_var)
var_entry=StringVar()

stile=ttk.Style()
stile.configure('W.TButton',foreground='red')

lbl=ttk.Label(master=frm_txt,textvariable=var_txt)
ent=ttk.Entry(master=frm_txt,textvariable=var_entry)
ent.grid(row=0,column=0)
ent.bind('<Return>',add_kwor_remove)
lbl.grid(row=1,column=0)

btn=ttk.Button(master=frm_btn,text='start scraping',command=functions.start_app)
btn.grid(row=0,column=0,sticky=(W,S))
btn_close=ttk.Button(master=root,command=close_stop_process,text='close',style='W.TButton')
btn_close.grid(row=2,column=2,columnspan=2,sticky=(E,S))

root.mainloop()
#pyinstaller -F  main.py --onefile --noconsole