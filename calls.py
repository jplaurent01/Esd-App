import csv
import sys
import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter
import datetime
from collections import Counter
from collections import defaultdict
import re
from tkinter import *
import random
import datetime
from tkinter import ttk
from tkinter import scrolledtext
#from tkinter.ttk import *
import tkinter as tk
from tkcalendar import Calendar
from tkinter import messagebox
import win32com.client
import pymsteams
from tkinter import filedialog as fd
import tkcalendar
from datetime import timedelta
from collections import OrderedDict
from datetime import datetime

def date_today():
    date_time = datetime.datetime.now()
    #print("Date time object:", date_time)
    fm_dt = date_time.strftime("%d_%m_%Y_%H_%M_%S")
    return fm_dt

def graph(quant, defects_quant, string, name, date1,date2):
    df = pd.DataFrame({'country': quant})
    df.index = defects_quant
    df = df.sort_values(by='country',ascending=False)
    df["cumpercentage"] = df["country"].cumsum()/df["country"].sum()*100


    fig, ax = plt.subplots()
    ax.bar(df.index, df["country"], color="C0")
    ax2 = ax.twinx()
    ax2.plot(df.index, df["cumpercentage"], color="C1", marker="D", ms=7)
    ax2.yaxis.set_major_formatter(PercentFormatter())

    ax.tick_params(axis="y", colors="C0")
    ax2.tick_params(axis="y", colors="C1")

    ax.set_xticklabels(df.index, rotation = 90)
    save_name = name + " Dudas " +"(" + date1.strftime("%d/%m/%Y") + " - "+ date2.strftime("%d/%m/%Y") +")"
    #plt.savefig(save_name)
    plt.title(save_name)
    #plt.xlabel("Defectos")
    #plt.ylabel("Cantidad")
    plt.show()

def count_break(line):
    count = 0
    for element in line:
        if element == "\n":
            count +=1
    return count

def isdate(line):
    flag_1 = False
    flag_2 = False
    flag_3 = False
    #print(line)
    #print(len(line))
    if len(line)>2:
        for i in range(len(line)):
            if line[0].isdigit()==True:
                #print(line[0])
                flag_1 =True
            if line[1].isdigit()==True or line[1] == "/":
                #print(line[1])
                flag_2 =True
            if line[2].isdigit()==True or line[2] == "/":
                #print(line[2])
                flag_3 = True
    else:
        pass
        #print(flag_1, flag_2, flag_3)
    return flag_1, flag_2, flag_3


def pareto_print(date1,date2, path, y2, name_what):
    if len(name_what) != 0:
        lst_date = date_range(date1,date2)
        #print(lst_date[0])
        name_what = name_what.strip()
        if lst_date[1] == True:
            with open(path, mode='r', encoding='utf8') as f:

                data = f.readlines()

            final_data_set = data[1:]

            chat_lines = []
            for i in range(len(final_data_set)):
                #message = line.split("-",1)
                text = isdate(final_data_set[i])
                if text[0]==True and text[1]==True and text[2]==True:
                    chat_lines.append(final_data_set[i])
                else:
                    pos = len(chat_lines)-1
                    str= chat_lines[pos] + " " + final_data_set[i]
                    chat_lines.pop()
                    chat_lines.append(str)

            date = []
            name_info = []
            split_lines = [i.split('-', 1) for i in chat_lines]
            for i in range(len(chat_lines)):
                date.append(chat_lines[i].split("-", 1)[0])
                name_info.append(chat_lines[i].split("-", 1)[1])


            date = [x.split()[0] for x in date]
            date = [datetime.strptime(x, '%d/%m/%Y') for x in date]
            date = [x.strftime('%d/%m/%Y') for x in date]
            #print(date)
            name = []
            info = []
            date_aux = []
            for i in range(len(name_info)):
                if ":" in name_info[i]:
                    date_aux.append(date[i])
                    name.append(name_info[i].split(":", 1)[0])
                    info.append(name_info[i].split(":", 1)[1])

            #print(Counter(name))
            name = [x.strip() for x in name]
            #print(name)
            print(len(name))
            print(len(info))
            print(len(date_aux))
            string = name_what
            #print(string)
            #string_name = str(string)
            if string != "Todos los inspectores":
                if string in name :
                    pos_def = []
                    defect = []
                    count_aux = 0

                    pos_name =[]
                    for i in range(len(name)):
                        if string in name[i]:
                            pos_name.append(i)
                    #print(pos_name)#Utilizar posiciones en info
                    #print(info)
                    for i in range(len(info)):
                        print(info[i])
                        if ("Duda" or "duda" or "Dudas" or "dudas" ) in info[i]:
                            #print(info[i])
                            if count_break(info[i])>=0:#Cumplo con condicion de 6 lineas
                                print(date_aux[i] in lst_date[0])
                                if string in name[i] and date_aux[i] in lst_date[0]:
                                    print(name[i])
                                    pos_def.append(i)
                                    print(info[i].split(":"))
                                    #print(info[i].split(":")[6])
                                    defect.append(info[i].split(":")[1].replace("-Ubicacion","").strip().lower())
                            else:
                                count_aux +=1

                    data = dict(Counter(defect))
                    quant = list(data.values())
                    defects_quant = list(data.keys())
                    print(date_aux)
                    print(lst_date)
                    if len(defect) != 0:
                        graph(quant, defects_quant, string, name_what, date1,date2)
                    else:
                        messagebox.showerror(message="La fecha de obtencion de datos no coincide con las fechas establecidas", title="Advertencia")

                else:
                    messagebox.showerror(message="No se encontro elemento", title="Advertencia")


            else:
                pos_def = []
                defect = []
                count_aux = 0

                #pos_name =[]
                #for i in range(len(name)):
                #    if string in name[i]:
                #        pos_name.append(i)
                #print(pos_name)#Utilizar posiciones en info
                #print(info)
                for i in range(len(info)):
                    print(info[i])
                    if ("Duda:" or "duda:" or "Dudas:" or "dudas:" ) in info[i]:
                        #print(info[i])
                        if info[i].split(":")[1].replace("-Ubicacion","").strip().lower() == '' and len(info[i].split(":")[1].replace("-Ubicacion","").strip().lower()) == 0:
                            pass
                        else:
                            if count_break(info[i])>=0:#Cumplo con condicion de 6 lineas
                                print(date_aux[i] in lst_date[0])
                                if date_aux[i] in lst_date[0]:
                                    print(name[i])
                                    pos_def.append(i)
                                    print(info[i].split(":"))
                                    #print(info[i].split(":")[6])
                                    defect.append(info[i].split(":")[1].replace("-Ubicacion","").strip().lower())
                            else:
                                count_aux +=1

                data = dict(Counter(defect))
                quant = list(data.values())
                defects_quant = list(data.keys())
                print(defects_quant[0])
                print(len(defects_quant[0]))
                print(data)
                print(date_aux)
                print(lst_date)
                if len(defect) != 0:
                    graph(quant, defects_quant, string, name_what, date1,date2)
                else:
                    messagebox.showerror(message="La fecha de obtencion de datos no coincide con las fechas establecidas", title="Advertencia")





















        else:
            messagebox.showerror(message="Verifique la fecha", title="Advertencia")
            y2.destroy()



    else:
        messagebox.showerror(message="Seleccione una persona", title="Advertencia")

def getname_file(file_name):
    flag  = True
    print(file_name)
    with open(file_name, mode='r', encoding='utf8') as f:

        data = f.readlines()
    final_data_set = data[1:]
    chat_lines = []
    for i in range(len(final_data_set)):
        #message = line.split("-",1)
        text = isdate(final_data_set[i])
        if text[0]==True and text[1]==True and text[2]==True:
            chat_lines.append(final_data_set[i])
        else:
            pos = len(chat_lines)-1
            str= chat_lines[pos] + " " + final_data_set[i]
            chat_lines.pop()
            chat_lines.append(str)

    date = []
    name_info = []
    split_lines = [i.split('-', 1) for i in chat_lines]
    for i in range(len(chat_lines)):
        date.append(chat_lines[i].split("-", 1)[0])
        name_info.append(chat_lines[i].split("-", 1)[1])

    date = [x.split()[0] for x in date]
    date = [datetime.strptime(x, '%d/%m/%Y') for x in date]
    date = [x.strftime('%d/%m/%Y') for x in date]
    name = []
    info = []
    date_aux = []
    for i in range(len(name_info)):
        if ":" in name_info[i]:
            #name.append(name_info[i].split(":", 1)[0])
            date_aux.append(date[i])
            name.append(name_info[i].split(":", 1)[0])
            info.append(name_info[i].split(":", 1)[1])

    res = list(OrderedDict.fromkeys(name))
    print(len(date))
    print(len(name_info))
    if len(date)==0 or len(name_info)==0 or len(name)==0 or len(date_aux)==0:
        flag= False
    else:
        pass
    print(flag)
    return res, flag

def date_range(start,stop):
    global dates_gb # If you want to use this outside of functions
    Flag = True
    dates_gb = []
    diff = (stop-start).days
    for i in range(diff+1):
        day = start + timedelta(days=i)
        dates_gb.append(day)
    dates_gb = [x.strftime('%d/%m/%Y') for x in dates_gb]#cAMBIAR FORMATO DE FECHA
    #dates_gb = [x.strptime(x,'%d/%m/%Y') for x in dates_gb]
    if dates_gb:
        print(dates_gb) # Print it, or even make it global to access it outside this
    else:
        Flag = False
        #print('Make sure the end date is later than start date')
    return dates_gb, Flag

def add_all_elements(lst):
    string = "Todos los inspectores"
    if string not in lst:
        lst.append(string)
    else:
        pass
    return lst

def create_pareto():
    y2 = Frame()
    y2.place(x=0, y=0, width=600, height=300)
    tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=0, column=0, pady=10, padx=10)

    aux = create_file()
    if aux[0] == True:
        Text_file = aux[1]
        #print(Text_file)
        lab1 = Label( y2, text="Fecha inicial:")
        lab1.grid(column=0, row=1)
        date1 = tkcalendar.DateEntry(y2)
        date1.grid(row=2, column=0, pady=10, padx=10)

        lab2 = Label( y2, text="Fecha final:")
        lab2.grid(column=0, row=3)
        date2 = tkcalendar.DateEntry(y2)
        date2.grid(row=4, column=0, pady=10, padx=10)

        lst_aux = getname_file(Text_file)
        if lst_aux[1] == True:
            lst = lst_aux[0]
            lst = add_all_elements(lst)
            options_list3 = lst
            value_inside3 = tk.StringVar(y2)
            lab4 = Label( y2, text="Seleccione un elemento de análisis")
            lab4.grid(column=1, row=1)
            value_inside3.set("")
            question_menu3 = tk.OptionMenu(y2, value_inside3, *options_list3)
            question_menu3.grid(column=1, row=2, padx=10, pady=10)
            lst_date = date_range(date1.get_date(),date2.get_date())
            #print(lst_date)
            tk.Button(y2, text='Generar Pareto', width=10, bg="black", fg='white', command=lambda:pareto_print(date1.get_date(),date2.get_date(), Text_file, y2, value_inside3.get())).grid(row=3, column=1, pady=10, padx=10)
        else:
            messagebox.showerror(message="Arcchivo no soportado", title="Advertencia")
    else :
        messagebox.showerror(message="No se seleccionó arcchivo", title="Advertencia")
        y2.destroy()

def create_file():
    flag = False
    filetypes = (
        ('text files', '*.txt'),
        ('All files', '*.*')
    )
    # show the open file dialog
    f = fd.askopenfile(filetypes=filetypes)
    try:
        print(f.name)
        name = os.path.abspath(f.name)
        flag = True
        return flag, name
    except :
        flag = False
        name = ""
        return flag, name

def open_file():
    y2 = Frame()
    y2.place(x=0, y=0, width=650, height=100)

    tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=0, column=0, pady=10, padx=10)
    tk.Button(y2, text='Abrir Archivo', width=10, bg="black", fg='white', command=lambda:create_pareto()).grid(row=0, column=1, pady=10, padx=10)
