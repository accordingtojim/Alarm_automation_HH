from tkinter import *
from tkinter.ttk import *
from tkinter.messagebox import showinfo
import array
import config
from openpyxl import load_workbook
from openpyxl import Workbook
import shutil
import os
import define_hybrid_house as dhh
import gen_HH.gen_AUX_HH,gen_HH.gen_BB_HH, gen_HH.gen_BR_HH,gen_HH.gen_HVAC_HH,gen_HH.gen_INV_HH,gen_HH.gen_PEMS_HH,gen_HH.gen_SKID_HH
# root window
root = Tk()
root.geometry("700x300")
root.resizable(True, True)
root.title('Number of Power Island')

number_of_PI = StringVar()
number_of_HH = StringVar()

# def PI_clicked():
#     # callback when the login button clicked
#     msg = f'You entered email: {number_of_PI.get()}' 
#     showinfo(title='Information',message=msg)

# empty label 1
empty_label = Label(root, text="")
empty_label.grid(row=0,column=0,columnspan=3,sticky=N,pady=20)

# PI label 
PI_label = Label(root, text="Insert the number of total Power Island")
PI_label.grid(row=1,column=2,padx=150)

# empty label 2
empty_label = Label(root, text="")
empty_label.grid(row=2,column=0,columnspan=3,sticky=N,pady=0)

# PI entry
PI_entry = Entry(root, textvariable=number_of_PI)
PI_entry.grid(row=3,column=2)
PI_entry.focus()

# empty label 3
empty_label = Label(root, text="")
empty_label.grid(row=4,column=0,columnspan=3,sticky=N,pady=20)

# HH label 
HH_label = Label(root, text="Insert the number of Hybrid House")
HH_label.grid(row=5,column=2,padx=150)

# empty label 4
empty_label = Label(root, text="")
empty_label.grid(row=6,column=0,columnspan=3,sticky=N,pady=0)

# HH entry
HH_entry = Entry(root, textvariable=number_of_HH)
HH_entry.grid(row=7,column=2)

# empty label 5
empty_label = Label(root, text="")
empty_label.grid(row=8,column=0,columnspan=3,sticky=N,pady=20)

# Define HH button
OK_button = Button(root,text='Define HH', command=dhh.openNewWindow)
OK_button.grid(row=7,column=2,padx=10,sticky=E)

# next button
# next_button = Button(root,text='Define PI', command=nw.openNewWindow)
# next_button.grid(row=3,column=2,padx=10,sticky=E)

# empty label
# empty_label = Label(root, text="")
# empty_label.grid(row=10,column=0,columnspan=3,sticky=N,pady=0)

# exit button
exit_button = Button(root,text='Exit',command=lambda:config.destroy_ww(root))
exit_button.grid(row=7,column=1,padx=20)

# quit button
# exit_button = Button(root,text='Quit',command=root.quit)
# exit_button.grid(row=11,column=3,padx=20)

root.mainloop()

#FINE PARTE GRAFICA
###############################################################################
#INIZIO GESTIONE FILE EXCEL

config.num_HH_GUI.append(number_of_HH.get())
config.n_PI = number_of_PI.get()


# hybrid house
if config.num_HH_GUI != '':
    # gen_HH.gen_INV_HH.file_creation_0('./template_excel/template_CL_CBESSHD.xlsx')
    # gen_HH.gen_BB_HH.file_creation_1('./template_excel/template_BB_samsung.xlsx')
    # gen_HH.gen_AUX_HH.file_creation_2('./template_excel/template_AUX_CBESSHD.xlsx')
    # gen_HH.gen_HVAC_HH.file_creation_4('./template_excel/template_HVAC_HH.xlsx')
    # gen_HH.gen_PEMS_HH.file_creation_3('./template_excel/template_PEMS_HH.xlsx')
    # gen_HH.gen_SKID_HH.file_creation_5('./template_excel/template_SKID_HH.xlsx')
    gen_HH.gen_BR_HH.file_creation_6('./template_excel/template_BR_samsung.xlsx')


config.file_aggregation(config.global_list)
config.file_removal(config.global_list)
config.file_reorder('list_alarm.xlsx','new_list_alarm.xlsx')
config.file_header('new_list_alarm.xlsx')
config.file_numbering('new_list_alarm.xlsx')
os.remove('list_alarm.xlsx')
print ('Program ended!')



    
    
    

