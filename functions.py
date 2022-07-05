import pandas as pd
import csv
import sys
import os
from collections import defaultdict
import functions as funct
import re
from tkinter import *
import random
import datetime
from tkinter import ttk
from tkinter import scrolledtext
from tkinter.ttk import *
import tkinter as tk
from operator import itemgetter
import win32com.client
import pymsteams
from re import search
#import pywhatkit
from tkinter import messagebox
from sys import intern
import csv
import sys
import os
from collections import defaultdict
import functions as funct
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
import calls
#import pywhatkit

n_array_first = []
count_aux0_first = 0


def whatsapp_analysis():
    Text_file = "Chat de WhatsApp con Equipo de Calidad EWCR.txt"


    with open(Text_file, mode='r', encoding='utf8') as f:

        data = f.readlines()

    final_data_set = data[3:]

    chat_lines = []
    for i in range(len(final_data_set)):
        #message = line.split("-",1)
        text = calls.isdate(final_data_set[i])
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

    print(name_info)
    name = []
    info = []
    for i in range(len(name_info)):
        if ":" in name_info[i]:
            name.append(name_info[i].split(":", 1)[0])
            info.append(name_info[i].split(":", 1)[1])

    #print(Counter(name))
    print(len(name))
    print(len(info))
    string = input("Ingrese un nombre")
    #string_name = str(string)
    if string in name :
        pos_def = []
        defect = []
        count_aux = 0

        pos_name =[]
        for i in range(len(name)):
            if string in name[i]:
                pos_name.append(i)
        print(pos_name)#Utilizar posiciones en info

        for i in range(len(info)):
            if ("Defecto:" or "defecto:") in info[i]:
                if calls.count_break(info[i])>5:#Cumplo con condicion de 6 lineas
                    if string in name[i]:
                        pos_def.append(i)
                        defect.append(info[i].split(":")[6].replace("\n Ubicación","").strip().lower())
                else:
                    count_aux +=1
        print(defect)
        print(len(defect))
        print(Counter(defect))
        data = dict(Counter(defect))
        quant = list(data.values())
        defects_quant = list(data.keys())
        print(quant)
        print(defects_quant)
        #calls.date_today()
        calls.graph(quant, defects_quant, string)

    else:
        print("No se encontro usuario")




def find_latin(string):
    boole = True
    for i in string:
        if i == "á" or i == "é" or i == "í" or i == "ó" or i=="ú" or i == "Á" or i == "É" or i == "Í" or i == "Ó" or i == "Ú":
            boole = False
            break
    return boole

def ene(string):
    count=True
    line=string
    for i in line:
        if i == "ñ" or i == "Ñ":
            count=False
            break
    print("Letra Ñ",count)
    return count

def count_space(string):
    count=0
    line=string
    for i in line:
        if(i.isspace()):
            count=count+1
    print("The number of blank spaces is: ",count)
    return count

def sear(listA, listB, var):
    #print(listB)
    if var == 0 :
        listB = [item.replace(" ","") for item in listB]
        rex = []
        #count = 0
        if len(listA)==1:
            for i in range(len(listB)):
                if listA[0] == listB[i] :
                    op = i + 2
                    rex.append(op)
                else:
                    pass
                #count +=1
        else:
            for i in range(len(listA)):
                for j in range(len(listB)):
                    if listA[i]==listB[j]:
                        op = j + 2
                        rex.append(op)
                    else:
                        pass
                    #count +=1
    elif var == 1:
        listB = [item.replace(" ","").upper() for item in listB]
        rex = []
        #count = 0
        if len(listA)==1:
            for i in range(len(listB)):
                if listA[0] == listB[i]:
                    rex.append(i)
                else:
                    pass
                #count +=1
        else:
            for i in range(len(listA)):
                for j in range(len(listB)):
                    if listA[i] == listB[j]:
                        rex.append(j)
                    else:
                        pass
    else:
        pass                #count +=1
    return rex

def edition_getval(tbn, value_inside, value_inside1, value_inside2, value_inside3,lst, lstbox2, seleccion,y2):
    lst_A = []
    for i in range(len(lst)):
        lst_A.append(lst[i][0].replace(" ", "").upper())
    print(lst_A)
    try:
        all_txt = read_file_1("ALL.TXT",0)
        names_txt = read_file_1(r"Z:\ESD Testing Reports\NAMES.TXT",1)
        #print(all_txt[1])
        index_0 = sear(lst_A, all_txt[1],1)
        lst_aux_names = []
        for i in range(len(names_txt[0])):
            string = names_txt[0][i] + names_txt[1][i]
            lst_aux_names.append(string)
        index_1 = sear(lst_A, lst_aux_names,0)
        print(index_0)
        print(index_1)

        data = []
        loc = str(value_inside.get())
        hor = str(value_inside1.get())
        area = str(value_inside2.get())
        trj = str(value_inside3.get())
        print("len(n_array):")
        print(len(n_array_first))
        aux_var = count_aux0_first - 1
        for row in range(len(n_array_first)):
            print(aux_var)
            if aux_var == row:
                for col in range(3):
                    print(n_array_first[row][col].get())
                    data.append(n_array_first[row][col].get())
            else:
                pass
        if len(data[0]) == 0 or len(data[1]) == 0 or len(data[2]) == 0:
            messagebox.showerror(message="Rellene espacios en blanco", title="Adveretencia")
        else:
            if loc == "Localizacion" or hor == "Horario" or area == "Area" or trj == "Manipulacion de tarjetas":
                messagebox.showerror(message="Ingrese una opcion valida", title="Adveretencia")

            else:

                count_data1 = funct.count_space(data[1].strip())
                count_data0 = funct.count_space(data[0].strip())
                if data[1].replace(" ", "").isalpha() == True and data[0].replace(" ", "").isalpha() == True  and count_data1 <2 and count_data0 <2 and funct.ene(data[1].strip()) == True and funct.ene(data[0].strip()) == True and funct.ene(data[2].strip()) == True and  find_latin(data[1].strip()) == True and find_latin(data[0].strip()) == True and find_latin(data[2].strip()) == True:
                    if len(lst_A)==1:
                        delete_lines(index_0, "ALL.TXT")
                        delete_lines(index_1, r"Z:\ESD Testing Reports\NAMES.TXT")
                        for selected_checkbox in seleccion[::-1]:
                            lstbox2.delete(selected_checkbox)
                        #messagebox.showinfo("Information","Eliminado Exitosamente")
                    view = already(data[1], data[0], data[2],y2)
                    if view == 0:
                        if trj == "Si" :
                            cond = -1
                            result = add(data[1], data[0], data[2], loc, hor, area, cond,y2)
                            messagebox.showinfo("Information","Usuario editado exitosamente")
                        elif trj == "No" :
                            cond = 0
                            result = add(data[1], data[0], data[2], loc, hor, area, cond,y2)
                            messagebox.showinfo("Information","Usuario editado exitosamente")
                    elif view == 1:
                        messagebox.showerror(message="Ingrese una contraseña distinta", title="Adveretencia")
                    else:
                        messagebox.showerror(message="Usuario ya existe", title="Adveretencia")
                else:
                    messagebox.showerror(message="Ingrese unicamente letras, como máximo 2 espacios por texto y no ingrese la letra ñ ni tildes", title="Advertencia")
    except IOError:
        messagebox.showerror(message="Compruebe la conexion a internet", title="Advertencia")
        y2.destroy()
def view_edition(name, apell, id, y2):
  y2 = Frame()
  y2.place(x=0, y=0, width=500, height=500)

  global n_array
  global count_aux0
  count_aux0 += 1

  options_list = ["Administracion", "Administracion*"," Hibrido", "Preformado", "Retrabajos", "Miscelanea", "Soldado", "Ensamble" , "Wave Solder" , "Etiquetado", "Inspeccion", "Materiales", "Test", "SMT", "EMA", "Empaque"]
  value_inside = tk.StringVar(y2)

  options_list1 = ["Diurno", "Nocturno", "Comprimido 1", "Comprimido 2"]
  value_inside1 = tk.StringVar(y2)

  options_list2 = ["SMT", "Preformado","Mantenimiento", "Test", "Calidad", "Limpieza", "Ingenieria", "Wave Solder ", "Contabilidad", "1-2", "3-4", "5-6", "7-8", "9-10", "Compras", "Etiquetado", "Logistica", "Materiales", "Salud Ocupacional", "AOI", "Finanzas", "Set Up", "Entrenamientos", "Produccion", "Retrabajos", "IT"]
  value_inside2 = tk.StringVar(y2)

  options_list3 = ["Si", "No"]
  value_inside3 = tk.StringVar(y2)

  row_array=[]    #array used to store a row
  n_array.append(row_array)
  y=len(n_array)

  for x in range(3):
      tbn="t"+str(y)+str(x)
      tbn1="t"+str(y)+str(x)
      tbn2="t"+str(y)+str(x)   #create entrybox names of the form t10, t11,...
        #print(tbn)

      if x==0:
        pos = x + 1
        label0 = Label(y2,text="Nombre")
        label0.grid(row = pos, column = 1)
        v = StringVar(root, value=name[0])
        tbn=Entry(y2, textvariable=v)
        row_array.append(tbn)
        row_array[x].grid(row=pos, column=2,sticky="nsew", padx=2,pady=2)

      if x==1:
        pos = x + 1
        label1 = Label(y2,text="Apellidos")
        label1.grid(row = pos, column = 1)
        v0 = StringVar(root, value=apell[0])
        tbn=Entry(y2, textvariable=v0)
        row_array.append(tbn)
        row_array[x].grid(row=pos, column=2,sticky="nsew", padx=2,pady=2)

      if x==2:
        pos = x + 1
        label1 = Label(y2,text="Contraseña")
        label1.grid(row = pos, column = 1)
        v1 = StringVar(root, value=id[0])
        tbn=Entry(y2, textvariable=v1,show= '*')
        row_array.append(tbn)
        row_array[x].grid(row=pos, column=2,sticky="nsew", padx=2,pady=2)

  lab1 = Label( y2, text="Localización")
  lab1.grid(column=2, row=4, padx=3, pady=2)
  value_inside.set("")
  question_menu = tk.OptionMenu(y2, value_inside, *options_list)
  question_menu.grid(column=2, row=5, padx=3, pady=2)

  lab2 = Label( y2, text="Horario")
  lab2.grid(column=2, row=6)
  value_inside1.set("")
  question_menu1 = tk.OptionMenu(y2, value_inside1, *options_list1)
  question_menu1.grid(column=2, row=7, padx=3, pady=2)

  lab3 = Label( y2, text="Area")
  lab3.grid(column=2, row=8)
  value_inside2.set("")
  question_menu2 = tk.OptionMenu(y2, value_inside2, *options_list2)
  question_menu2.grid(column=2, row=9, padx=3, pady=2)

  lab4 = Label( y2, text="Manipulacion de tarjetas")
  lab4.grid(column=2, row=10)
  value_inside3.set("")
  question_menu3 = tk.OptionMenu(y2, value_inside3, *options_list3)
  question_menu3.grid(column=2, row=11, padx=0, pady=5)

  #Button(y2, text="Add new row", command=lambda:add_four_entries(y2)).grid(row=0, column=0,)
  Button(y2, text="Confirmar", width=10, bg="#116562", fg='#f7fafa',activebackground='#055959',
activeforeground='#f7fafa', command=lambda:edition_getval(tbn, value_inside, value_inside1, value_inside2, value_inside3,y2)).grid(row=12, column=2)

  tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=0, column=0, pady=10, padx=10)


def edit_tod_dis(lstbox2,y2):
    reslist = list()
    seleccion = lstbox2.curselection()
    lst = []
    for i in seleccion:
        entrada = lstbox2.get(i)
        reslist.append(entrada)
    for val in reslist:
        lst.append(val)
    if len(lst)==0:
        messagebox.showerror(message="Selecione un nombre", title="Advertencia")
    elif len(lst)==1:
        lst = [lst[i].split("-",3) for i in range(len(lst))]
        print(lst)
        print(lst[0][0])
        xname = lst[0][0].split(" ")
        xname = [ele for ele in xname if ele.strip()]
        print(xname)
        print(len(xname))
        ap = ""
        nm = ""
        if len(xname)==2:
            ap = xname[0]
            nm = xname[1]
        elif len(xname)==3:
            ap = xname[0] + " " +xname[1]
            nm = xname[2]
        elif len(xname)==4:
            ap = xname[0] + " " +xname[1]
            nm = xname[2] + " " +xname[3]
        y2 = Frame()
        y2.place(x=0, y=0, width=500, height=500)
        global n_array_first
        global count_aux0_first
        count_aux0_first += 1

        options_list = ["Administracion", "Administracion*"," Hibrido", "Preformado", "Retrabajos", "Miscelanea", "Soldado", "Ensamble" , "Wave Solder" , "Etiquetado", "Inspeccion", "Materiales", "Test", "SMT", "EMA", "Empaque"]
        value_inside = tk.StringVar(y2)

        options_list1 = ["Diurno", "Nocturno", "Comprimido 1", "Comprimido 2"]
        value_inside1 = tk.StringVar(y2)

        options_list2 = ["SMT", "Preformado","Mantenimiento", "Test", "Calidad", "Limpieza", "Ingenieria", "Wave Solder ", "Contabilidad", "1-2", "3-4", "5-6", "7-8", "9-10", "Compras", "Etiquetado", "Logistica", "Materiales", "Salud Ocupacional", "AOI", "Finanzas", "Set Up", "Entrenamientos", "Produccion", "Retrabajos", "IT"]
        value_inside2 = tk.StringVar(y2)

        options_list3 = ["Si", "No"]
        value_inside3 = tk.StringVar(y2)

        row_array=[]    #array used to store a row
        n_array_first.append(row_array)
        y=len(n_array_first)

        for x in range(3):
            tbn="t"+str(y)+str(x)
            tbn1="t"+str(y)+str(x)
            tbn2="t"+str(y)+str(x)   #create entrybox names of the form t10, t11,...
                      #print(tbn)

            if x==0:
                pos = x + 1
                label0 = Label(y2,text="Nombre")
                label0.grid(row = pos, column = 1)
                tbn=Entry(y2)
                tbn.insert(0, nm)
                row_array.append(tbn)
                row_array[x].grid(row=pos, column=2,sticky="nsew", padx=2,pady=2)

            if x==1:
                pos = x + 1
                label1 = Label(y2,text="Apellidos")
                label1.grid(row = pos, column = 1)
                tbn=Entry(y2)
                tbn.insert(0, ap)
                row_array.append(tbn)
                row_array[x].grid(row=pos, column=2,sticky="nsew", padx=2,pady=2)

            if x==2:
                pos = x + 1
                label1 = Label(y2,text="Contraseña")
                label1.grid(row = pos, column = 1)
                tbn=Entry(y2, show= '*')
                row_array.append(tbn)
                row_array[x].grid(row=pos, column=2,sticky="nsew", padx=2,pady=2)

        lab1 = Label( y2, text="Localización")
        lab1.grid(column=2, row=4, padx=3, pady=2)
        value_inside.set("")
        question_menu = tk.OptionMenu(y2, value_inside, *options_list)
        question_menu.grid(column=2, row=5, padx=3, pady=2)

        lab2 = Label( y2, text="Horario")
        lab2.grid(column=2, row=6)
        value_inside1.set("")
        question_menu1 = tk.OptionMenu(y2, value_inside1, *options_list1)
        question_menu1.grid(column=2, row=7, padx=3, pady=2)

        lab3 = Label( y2, text="Area")
        lab3.grid(column=2, row=8)
        value_inside2.set("")
        question_menu2 = tk.OptionMenu(y2, value_inside2, *options_list2)
        question_menu2.grid(column=2, row=9, padx=3, pady=2)

        lab4 = Label( y2, text="Manipulacion de tarjetas")
        lab4.grid(column=2, row=10)
        value_inside3.set("")
        question_menu3 = tk.OptionMenu(y2, value_inside3, *options_list3)
        question_menu3.grid(column=2, row=11, padx=0, pady=5)

                #Button(y2, text="Add new row", command=lambda:add_four_entries(y2)).grid(row=0, column=0,)
        Button(y2, text="Confirmar", width=10, bg="#116562", fg='#f7fafa',activebackground='#055959',
              activeforeground='#f7fafa', command=lambda:edition_getval(tbn, value_inside, value_inside1, value_inside2, value_inside3, lst, lstbox2, seleccion,y2)).grid(row=12, column=2)
                #photo20 = tk.PhotoImage(file = r"return.PNG")
                # Resizing image to fit on button
                #photoimage20 = photo20.subsample(10, 10)
        tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=0, column=0, pady=10, padx=10)

    else:
        messagebox.showerror(message="Selecione unicamente un elemento", title="Advertencia")


def tod_dis(lstbox2,y2):
    reslist = list()
    seleccion = lstbox2.curselection()
    lst = []
    for i in seleccion:
        entrada = lstbox2.get(i)
        reslist.append(entrada)
    for val in reslist:
        lst.append(val)
    if len(lst)==0:
        messagebox.showerror(message="Selecione un nombre", title="Advertencia")
    else:
        lst = [lst[i].split("-",1) for i in range(len(lst))]
        lst_A = []
        for i in range(len(lst)):
            lst_A.append(lst[i][0].replace(" ", "").upper())
        #lst = [lst[0][i].replace(" ", "") for i in range(len(lst))]
        print(lst_A)
        try:
            all_txt = read_file_1("ALL.TXT",0)
            names_txt = read_file_1(r"Z:\ESD Testing Reports\NAMES.TXT",1)
            #print(all_txt[1])
            index_0 = sear(lst_A, all_txt[1],1)
            lst_aux_names = []
            for i in range(len(names_txt[0])):
                string = names_txt[0][i] + names_txt[1][i]
                lst_aux_names.append(string)
            #print(lst_aux_names)
            index_1 = sear(lst_A, lst_aux_names,0)
            print(index_0)
            print(index_1)

            if len(lst_A)==1:
                answer = messagebox.askyesno(title='confirmation',
                        message='Desea eliminar el usuario')
                print(answer)
                if answer == True:
                    delete_lines(index_0, "ALL.TXT")
                    delete_lines(index_1, r"Z:\ESD Testing Reports\NAMES.TXT")
                    for selected_checkbox in seleccion[::-1]:
                        lstbox2.delete(selected_checkbox)
                    messagebox.showinfo("Information","Eliminado Exitosamente")
                else:
                    pass
            else:
                answer = messagebox.askyesno(title='confirmation',
                        message='Desea eliminar los usuarios')
                print(answer)
                if answer == True:
                    delete_lines(index_0, "ALL.TXT")
                    delete_lines(index_1, r"Z:\ESD Testing Reports\NAMES.TXT")
                    for selected_checkbox in seleccion[::-1]:
                        lstbox2.delete(selected_checkbox)
                    messagebox.showinfo("Information","Eliminados Exitosamente")
                else:
                    pass
        except IOError:
            messagebox.showerror(message="Compruebe la conexion a internet", title="Advertencia")
            y2.destroy()

def find_index(index, lst):
    boole = True
    for i in range(len(lst)):
        if i == lst[i]:
            boole = True
        else:
            boole = False
    return boole

def delete_lines(index, file):
    # list to store file lines
    lines = []
    # read file
    with open(file, 'r') as fp:
        # read an store all lines into list
        lines = fp.readlines()

    # Write file
    with open(file, 'w') as fp:
        # iterate each line
        for number, line in enumerate(lines):
            # delete line 5 and 8. or pass any Nth line you want to remove
            # note list index starts from 0
            if number not in index:
                fp.write(line)

def read_file_1(filename, var):
    if var == 0:
        columns = defaultdict(list)
        with open(filename, 'r', encoding="ISO-8859-1") as f:
            reader = csv.reader(f, delimiter=',')
            for row in reader:
                for i in range(len(row)):
                    columns[i].append(row[i])

        columns = dict(columns)
    elif var == 1:
        columns = defaultdict(list)
        with open(filename, 'r', encoding="ISO-8859-1") as f:
            lines_after_2 = f.readlines()[2:]
            reader = csv.reader(lines_after_2, delimiter=',')
            for row in reader:
                for i in range(len(row)):
                    columns[i].append(row[i])

        columns = dict(columns)
    else:
        pass
    return columns

def read_file(filename):
    columns = defaultdict(list)
    with open(filename, 'r', encoding="ISO-8859-1") as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
            for i in range(len(row)):
                columns[i].append(row[i])

    columns = dict(columns)

    return columns

def vec(file):
    if file == r"Z:\ESD Testing Reports\NAMES.TXT":
        columns = defaultdict(list)
        with open(file, 'r', encoding="ISO-8859-1") as f:
            lines_after_2 = f.readlines()[2:]
            reader = csv.reader(lines_after_2, delimiter=',')
            for row in reader:
                for i in range(len(row)):
                    columns[i].append(row[i])
        columns = dict(columns)
    else:
        columns = defaultdict(list)
        with open(file, 'r', encoding="ISO-8859-1") as f:
            reader = csv.reader(f, delimiter=',')
            for row in reader:
                for i in range(len(row)):
                    columns[i].append(row[i])

        columns = dict(columns)

    return columns

def already(apell, nom, id, y2):
    apell = apell.strip()
    nom = nom.strip()
    id = id.strip()
    try:
        NAMES = vec(r"Z:\ESD Testing Reports\NAMES.TXT")
        ALL = vec("ALL.TXT")

        ALL_aux0 = []
        ALL_aux1 = []
        ALL_aux2 = []
        aux0_NAMES = []
        aux1_NAMES = []
        aux2_NAMES = []
        Flag = 0
        for i in range(len(NAMES[0])):
            temp0 = NAMES[0][i].strip()
            aux0_NAMES.append(temp0.replace(" ", ""))
            aux1_NAMES.append(NAMES[1][i].strip())
        aux2_NAMES = [i.split(':', 1)[1] for i in NAMES[2]]
        #aux2_NAMES = aux2_NAMES.strip()
        aux2_NAMES = [i.strip() for i in aux2_NAMES]
        for i in range(len(ALL[0])):
            temp = ALL[1][i].strip()
            ALL_aux0.append(temp.replace(" ", "").upper())
            ALL_aux1.append(ALL[5][i].strip())
        #print(ALL_aux0)
        var_aux = apell.strip().upper().replace(" ", "") + nom.upper().strip()
        var_aux = var_aux.replace(" ", "")
        print("Apellidos y nombre en mayuscula::")
        print(var_aux)
        print("Apellidos mayuscula:")
        print(apell.strip().upper().replace(" ", ""))
        print("Nombre en mayuscula:")
        print(nom.upper().strip().replace(" ", ""))
        print("IDENTIFICACION:")
        print(id.strip())
        print("Lista apellidos")
        print(aux0_NAMES)
        print("Lista nombres")
        print(aux1_NAMES)
        print("Lista identificaciones")
        print(aux2_NAMES)
        if (apell.strip().upper().replace(" ", "")  not in aux0_NAMES and nom.upper().strip().replace(" ", "") not in aux1_NAMES and id.strip() not in  aux2_NAMES) or (var_aux not in ALL_aux0  and id.strip() not in ALL_aux1):
            Flag = 0
        #elif (apell.strip().upper().replace(" ", "")  not in aux0_NAMES and nom.upper().strip().replace(" ", "") not in aux1_NAMES and id.strip() in  aux2_NAMES) or (apell.strip().upper().replace(" ", "")  not in aux0_NAMES and nom.upper().strip().replace(" ", "") in aux1_NAMES and id.strip() in  aux2_NAMES) or (apell.strip().upper().replace(" ", "")   in aux0_NAMES and nom.upper().strip().replace(" ", "") in aux1_NAMES and id.strip() in  aux2_NAMES) or (var_aux not in ALL_aux0  and id.strip() in ALL_aux1)  :
        elif  id.strip() in  aux2_NAMES or id.strip() in ALL_aux1  :
            Flag = 1
        else:
            Flag = 2

        return Flag
    except IOError:
        messagebox.showerror(message="Compruebe la conexion a internet", title="Advertencia")
        y2.destroy()

def intostring(lista):
    listb = '\n'.join(lista)
    return listb


def outlook (list_fail, list_general ):
    today = datetime.date.today()
    #print(len(list_fail))
    #print(len(list_general))
    #Caso donde no hay personal que fallo test
    route = r"Z:\Monitoreo y control de prevención de ESD\Registro de comunicados incumplimiento Test ESD" + today.strftime('%b-%d-%Y') +".msg"
    if len(list_fail)== 0 :
        if len(list_general) == 0 :
            #Caso donde no se encuentra incumplimiento
            string = '<h3></h3>'
            str = 'El día de hoy no se encuentran incumplimientos al reglamento ESD.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str
            #mail.Attachments.Add('c:\\sample.xlsx')
            #mail.Attachments.Add('c:\\sample2.xlsx')
            ##mail.CC = 'EWCR-All@ewmfg.com; testcr@ewmfg.com.com'
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")

        elif len(list_general) == 1 :
            #Caso deonde hay solo una persona que vino y no se testeo
            string = '<h3></h3>'
            str_0 = 'Buenos días,\n'
            str_1 = 'La siguiente persona no reporta haberse testeado el día de hoy. Por favor, indicarle que debe realizar el test pues no deberían estar manipulando material:\n\n '
            str_2 = intostring(list_general)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 + str_1 + str_2 + str_3
                #mail.Attachments.Add('c:\\sample.xlsx')
                #mail.Attachments.Add('c:\\sample2.xlsx')
                ##mail.CC = 'EWCR-All@ewmfg.com; testcr@ewmfg.com.com'
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")
        else:
            #Caso donde vino más de una persona y no se testearon
            string = '<h3></h3>'
            str_0 = 'Buenos días,\n'
            str_1 = 'Las siguientes personas no reportan haberse testeado el día de hoy. Por favor, indicarles que deben realizar el test pues no deberían estar manipulando material:\n\n '
            str_2 = intostring(list_general)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 + str_1 + str_2 + str_3
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")

    elif len(list_fail) == 1 :
        #Caso donde solo una persona fallo test
        if len(list_general) == 1:
            #Caso donde una persona fallo y una no se testo
            string = '<h3></h3>'
            str_0 = 'Buenos días,'
            str_4 = '\nLa siguiente persona se testeó, sin embargo, su equipo está dañado:\n\n'
            str_5 = intostring(list_fail)
            str_1 = '\n\nPor otro lado, la siguiente persona no reporta haberse testeado el día de hoy. Por favor, indicarle que debe realizar el test pues no deberían estar manipulando material:\n'
            str_2 = intostring(list_general)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 +  str_4 + str_5 + str_1 + str_2 + str_3
                #mail.Attachments.Add('c:\\sample.xlsx')
                #mail.Attachments.Add('c:\\sample2.xlsx')
                ##mail.CC = 'EWCR-All@ewmfg.com; testcr@ewmfg.com.com'
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")

        elif len(list_general) == 0:
            #Caso donde una persona fallo y ningunna no se testa
            string = '<h3></h3>'
            str_0 = 'Buenos días,\n'
            str_4 = '\nLa siguiente persona se testeó, sin embargo, su equipo está dañado:\n\n'
            str_5 = intostring(list_fail)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 +  str_4 + str_5 + str_3
                #mail.Attachments.Add('c:\\sample.xlsx')
                #mail.Attachments.Add('c:\\sample2.xlsx')
                ##mail.CC = 'EWCR-All@ewmfg.com; testcr@ewmfg.com.com'
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")

        else:
            #Caso donde una persona fallo y varias  no se testearona
            string = '<h3></h3>'
            str_0 = 'Buenos días,\n\n'
            str_4 = 'La siguiente persona se testeó, sin embargo, su equipo está dañado:\n\n'
            str_5 = intostring(list_fail)
            str_1 = '\n\nPor otro lado, las siguientes personas no reportan haberse testeado el día de hoy. Por favor, indicarles que deben realizar el test pues no deberían estar manipulando material:\n'
            str_2 = intostring(list_general)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 +  str_4 + str_5 + str_1 + str_2 + str_3
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")
    else :
        #Caso donde varias personas falloron test
        if len(list_general) == 1:
            #Caso donde varias fallaron y una no se testo
            string = '<h3></h3>'
            str_0 = 'Buenos días,\n\n'
            str_4 = 'Las siguientes personas se testearon, sin embargo, sus equipos están dañados:\n\n'
            str_5 = intostring(list_fail)
            str_1 = '\n\nPor otro lado, la siguiente persona no reporta haberse testeado el día de hoy. Por favor, indicarle que debe realizar el test pues no deberían estar manipulando material:\n\n'
            str_2 = intostring(list_general)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 +  str_4 + str_5 + str_1 + str_2 + str_3
            #mail.Attachments.Add('c:\\sample.xlsx')
            #mail.Attachments.Add('c:\\sample2.xlsx')
            ##mail.CC = 'EWCR-All@ewmfg.com; testcr@ewmfg.com.com'
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")

        elif len(list_general) == 0:
            #Caso donde una persona fallo y ningunna no se testa
            string = '<h3></h3>'
            str_0 = 'Buenos días,'
            str_4 = '\nLas siguientes personas se testearon, sin embargo, sus equipos estan dañados:\n\n'
            str_5 = intostring(list_fail)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 +  str_4 + str_5 + str_3
                #mail.Attachments.Add('c:\\sample.xlsx')
                #mail.Attachments.Add('c:\\sample2.xlsx')
                ##mail.CC = 'EWCR-All@ewmfg.com; testcr@ewmfg.com.com'
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")

        else:
            #Caso donde varias fallaron y varis  no se testearona
            string = '<h3></h3>'
            str_0 = 'Buenos días,\n\n'
            str_4 = 'Las siguientes personas se testearon, sin embargo, sus equipos están dañados:\n\n'
            str_5 = intostring(list_fail)
            str_1 = '\n\nPor otro lado, las siguientes personas no reportan haberse testeado el día de hoy. Por favor, indicarles que deben realizar el test pues no deberían estar manipulando material:\n'
            str_2 = intostring(list_general)
            str_3 = '\n\nPara consultas y aclaraciones favor responder solo a este correo para evitar el spam a los demás compañeros.'
            outlook = win32com.client.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = 'EWCR-All@ewmfg.com; testcr@ewmfg.com'
            mail.Subject = 'Incumplimiento de reglamento ESD'
            mail.HTMLBody = string
            mail.Body = str_0 +  str_4 + str_5 + str_1 + str_2 + str_3
            mail.Send()
            #mail.SaveAs(Path=route)
            messagebox.showinfo("Information","Mensaje Exitosamente enviado")


def fail0_show_selection(y2, Fail0):
      for pa in Fail0:
          Fail0_display = Text(master=y2, height=20, width=60, bg="Lightgreen",font="Helvetica 10 bold")
          Fail0_display.grid(row=5, column=1, padx=10)
          Fail0_display.insert(END, pa)



def bad_tested(a,name,apell, wok, rok, lok):
  res = []
  for i in range(len(a)):
    pos = a[i]
    #print(pos)
    if ((wok[pos] == " ComFail" or wok[pos] == " Whi" or wok[pos] == " W-") or rok[pos] != " Lok"  or lok[pos] != " Rok") :
      res.append(apell[pos] + ""+ name[pos] + "->" + wok[pos] + "," + rok[pos] + "," + lok[pos])
    pos = 0
  return res


def position(a,b,name, apell):
  pos = []
  count = 0
  for i in range(len(name)):
    #print(a)
    #print(name[i])
    if a == name[i] and b == apell[i]:
      pos.append(count)
    count+=1
  return pos

def add_list_display( y2):



  #print(var_array)
  name = var_array[0]
  apellido = var_array[1]
  ID = var_array[2]



  add_list = funct.add(name, apellido, ID, y2)
  add_list_display = scrolledtext.ScrolledText(y2, height=10, width=50)
  for i in add_list:
    add_list_display.insert(END, (i + "\n"))
    add_list_display.grid(row=5, column=1,padx=10)


def sec(lst):
    ed_22 = []
    ed_23 = []
    Materiales = []
    contain = ["SMT", "AOI", "Etiquetado", "Materiales"]
    display = ""
    for i in range(len(lst)):
        seg = []
        #Entre parentesis estaba funcion search en todos los 3 condicionales siguientes
        if (contain[0] in lst[i]) or (contain[1] in lst[i]) or (contain[2] in lst[i]):
            seg = lst[i].split("->")
            seg = str(seg[0])
            ed_22.append(seg)
        elif (contain[3] in lst[i]):
            seg = lst[i].split("->")
            seg = str(seg[0])
            #print(type(seg))
            Materiales.append(seg)
        else:
            seg = lst[i].split("->")
            seg = str(seg[0])
            #print(type(seg))
            ed_23.append(seg)
    print(len(ed_22))
    print(len(ed_23))
    print(len(Materiales))
    if len(ed_22) == 0 and len(ed_23)== 0 and len(Materiales)==0:
        display = "No falto nadie"
    elif len(ed_22) == 0 and len(ed_23)!=0 and len(Materiales)==0:
        print("Flag2")
        text = "Edificio 23:\n"
        text_temp = intostring(ed_23)
        desplay = text + text_temp
        #print(desplay)
    elif len(ed_22) == 0 and len(ed_23) == 0 and len(Materiales)!=0:
        text = "Materiales:\n"
        text_temp = intostring(Materiales)
        desplay = text + text_temp
    elif len(ed_22) != 0 and len(ed_23) == 0 and len(Materiales)==0:
        text = "Edificio 22:\n"
        text_temp = intostring(ed_22)
        desplay = text + text_temp
    elif len(ed_22) != 0 and len(ed_23) != 0 and len(Materiales)==0:
         text0 = "Edificio 23:\n"
         text_temp0 = intostring(ed_23)
         text = "\nEdificio 22:\n"
         text_temp = intostring(ed_22)
         desplay = text0 + text_tem0 + text + text_temp
    elif len(ed_22) == 0 and len(ed_23) != 0 and len(Materiales)!=0:
         text0 = "Edificio 23:\n"
         text_temp0 = intostring(ed_23)
         text = "\nMateriales:\n"
         text_temp = intostring(Materiales)
         desplay = text0 + text_tem0 + text + text_temp
    elif len(ed_22) != 0 and len(ed_23) == 0 and len(Materiales)!=0:
         text0 = "Edificio 22:\n"
         text_temp0 = intostring(ed_22)
         text = "\nMateriales:\n"
         text_temp = intostring(Materiales)
         desplay = text0 + text_tem0 + text + text_temp
    elif len(ed_22) != 0 and len(ed_23) != 0 and len(Materiales)!=0:
        text0 = "Edificio 23:\n"
        text_temp0 = intostring(ed_23)
        text = "\nEdificio 22:\n"
        text_temp = intostring(ed_22)
        text1 = "\nMateriales:\n"
        text_temp1 = intostring(Materiales)
        desplay = text0 + text_temp0 + text + text_temp + text1 + text_temp1
    else:
        print("here")
        pass
    return desplay

def send_message(lst, fecha):
    myTeamsMessage = pymsteams.connectorcard("https://appriver3651009991.webhook.office.com/webhookb2/59c314f4-3626-4c5c-8c40-1cd6d93142ea@2b27f5d3-a252-44ed-8619-440662038adb/IncomingWebhook/067d65a96b5c4fe08b2b64557dfd04c5/0c02eb15-6535-4dcd-9077-217dd02ae497")
    string = str(sec(lst))
    print(type(string) )
    print(string)
    myTeamsMessage.text(string)
    myTeamsMessage.send()
    messagebox.showinfo("Information","Mensaje Exitosamente enviado")

def whatsapp_list(fecha, y2):
    myTuple = funct.lists(fecha, y2)
  #ok
    first = myTuple[0]
  #not ok
    second = myTuple[1]

    columns = myTuple[2]

    nick_name0 = myTuple[3].values.tolist()
    nick_name = [i[0] for i in nick_name0]
    name0 = myTuple[4].values.tolist()
    name = [i[0] for i in name0]
    Wok0 = myTuple[5].values.tolist()
    Wok = [i[0] for i in Wok0]
    Lok0 = myTuple[6].values.tolist()
    Lok = [i[0] for i in Lok0]
    Rok0 =  myTuple[7].values.tolist()
    Rok = [i[0] for i in Rok0]

    index = []
    for i in range(len(name)):
        index.append(funct.position(name[i],nick_name[i],name,nick_name))

  #print(index)

  #print('########################################')

    res = []
    for i in index:
        if i not in res:
            res.append(i)

  #print(res)
    index_res = []
    for i in range(len(res)):
        index_res.append(max(res[i]))
  #print('########################################')
  #print(index_res)
    #ok sorted
    sorted_list0 = sorted(first)
    #not ok sorted
    sorted_list1 = sorted(second)
  #print(sorted_list1)
    #reporte de test
    report = sorted(sorted_list0 + sorted_list1)
    #personas con test negativo
    bad = funct.compare(sorted_list0, sorted_list1)
  #print(bad)
    #Todas las personas
    every = funct.all()
    #Comprimido 1
    tot_com1 = every[0]
    #Toda la informacion
    data = every[1]
    #Comprimido 2
    tot_com2 = every[2]
    #Gente que no es comprimido 1 ni 2
    get = every[3]
    #Todos maysucula
    com1 = tot_com1 + get
    every0 = [x.upper() for x in com1]
    eve1 = [0]*len(every0)
    #Eliminar espacios inico y final
    for i in range(len(every0)):
        eve1[i] = every0[i].strip()

    #todo maysucula
    report0=[x.upper() for x in report]
    rep1 = [0]*len(report0)
    #Eliminar espacios
    for i in range(len(report0)):
        rep1[i] = report0[i].strip()
    #Personas que no llegaron comprimido 1
    not_came_compri1 = sorted(list(set(eve1) - set(rep1)))
    #Personas que llegaron comprimido 1
    came_compri1 = sorted(list(set(eve1) & set(rep1)))

    #Todos maysucula
    com2 = tot_com2 + get
    EVERYTHING2 = [x.upper() for x in com2]
    EVE2 = [0]*len(EVERYTHING2)
    #Eliminar espacios
    for i in range(len(EVERYTHING2)):
        EVE2[i] =  EVERYTHING2[i].strip()

    #todo maysucula
    REPORT2=[x.upper() for x in report]
    REP2 = [0]*len(REPORT2)
    #Eliminar espacios
    for i in range(len(REPORT2)):
        REP2[i] = REPORT2[i].strip()
    #Personas que no llegaron comprimido 2
    not_came_COMPRI2 = sorted(list(set(EVE2) - set(REP2)))
    #Personas que llegaron comprimido 2
    came_COMPRI2 = sorted(list(set(EVE2) & set(REP2)))

    #No viene comprimido 1
    match1 = funct.print0(data, not_came_compri1)

    SORT_MATCH1 = sorted(match1)

    #No viene comprimido 2
    match2 = funct.print0(data, not_came_COMPRI2)

    SORT_MATCH2 = sorted(match2)
    #print0(SORT_MATCH2)
    #Union con errors
    union = funct.matches(bad, nick_name, name, Wok, Lok, Rok)
    #Testeo fallido
  #fail = list(set(union))
  #print(fail)
    result_bad = funct.bad_tested(index_res, name, nick_name, Wok, Lok, Rok)
  #print(result_bad)
    fail = result_bad
    # pyinstaller --onefile --add-binary='/System/Library/Frameworks/Tk.framework/Tk':'tk' --add-binary='/System/Library/Frameworks/Tcl.framework/Tcl':'tcl' part_manager.py
    # '''
    Fail0 = fail
    Comprimido1 = SORT_MATCH1
    Comprimido2 = SORT_MATCH2
    tot_pers = Comprimido1 + Comprimido2
    tot_pers = list(set(tot_pers))
    tot_pers = sorted(tot_pers)
    print(tot_pers)
    print(len(tot_pers))
    #print(tot_pers)
    #print(len(tot_pers))

    return Comprimido1, Comprimido2, tot_pers

#  myTuple = lists(fecha)
  #ok
#  first = myTuple[0]
    #not ok
#  second = myTuple[1]

#  columns = myTuple[2]

#  nick_name0 = myTuple[3].values.tolist()
#  nick_name = [i[0] for i in nick_name0]
#  name0 = myTuple[4].values.tolist()
#  name = [i[0] for i in name0]
#  Wok0 = myTuple[5].values.tolist()
#  Wok = [i[0] for i in Wok0]
#  Lok0 = myTuple[6].values.tolist()
#  Lok = [i[0] for i in Lok0]
#  Rok0 =  myTuple[7].values.tolist()
#  Rok = [i[0] for i in Rok0]

    #ok sorted
#  sorted_list0 = sorted(first)
    #not ok sorted
#  sorted_list1 = sorted(second)
    #reporte de test
#  report = sorted(sorted_list0 + sorted_list1)
    #personas con test negativo
#  bad = compare(sorted_list0, sorted_list1)

    #Todas las personas
#  every = all()
    #Comprimido 1
#  tot_com1 = every[0]
    #Toda la informacion
#  data = every[1]
    #Comprimido 2
#  tot_com2 = every[2]
    #Gente que no es comprimido 1 ni 2
#  get = every[3]
    #Todos maysucula
#  com1 = tot_com1 + get
#  every0 = [x.upper() for x in com1]
#  eve1 = [0]*len(every0)
    #Eliminar espacios inico y final
#  for i in range(len(every0)):
#    eve1[i] = every0[i].strip()

    #todo maysucula
#  report0=[x.upper() for x in report]
#  rep1 = [0]*len(report0)
    #Eliminar espacios
#  for i in range(len(report0)):
#    rep1[i] = report0[i].strip()
    #Personas que no llegaron comprimido 1
#  not_came_compri1 = sorted(list(set(eve1) - set(rep1)))
    #Personas que llegaron comprimido 1
#  came_compri1 = sorted(list(set(eve1) & set(rep1)))

    #Todos maysucula
#  com2 = tot_com2 + get
#  EVERYTHING2 = [x.upper() for x in com2]
#  EVE2 = [0]*len(EVERYTHING2)
    #Eliminar espacios
#  for i in range(len(EVERYTHING2)):
#    EVE2[i] =  EVERYTHING2[i].strip()

    #todo maysucula
#  REPORT2=[x.upper() for x in report]
#  REP2 = [0]*len(REPORT2)
    #Eliminar espacios
#  for i in range(len(REPORT2)):
#    REP2[i] = REPORT2[i].strip()
    #Personas que no llegaron comprimido 2
#  not_came_COMPRI2 = sorted(list(set(EVE2) - set(REP2)))
    #Personas que llegaron comprimido 2
#  came_COMPRI2 = sorted(list(set(EVE2) & set(REP2)))

    #No viene comprimido 1
#  match1 = print0(data, not_came_compri1)

#  SORT_MATCH1 = sorted(match1)

    #No viene comprimido 2
#  match2 = print0(data, not_came_COMPRI2)

#  SORT_MATCH2 = sorted(match2)
    #print0(SORT_MATCH2)
    #Union con errors
#  union = matches(bad, nick_name, name, Wok, Lok, Rok)

    #Testeo fallido
#  fail = list(set(union))


    # pyinstaller --onefile --add-binary='/System/Library/Frameworks/Tk.framework/Tk':'tk' --add-binary='/System/Library/Frameworks/Tcl.framework/Tcl':'tcl' part_manager.py
    # '''

#  Fail0 = fail
#  Comprimido1 = SORT_MATCH1
#  Comprimido2 = SORT_MATCH2

  #Submit = Button(y2, text = "Eliminar Usuario", command = lambda: C==add_list_display()).grid(row = 2, column = 0)
#  return Comprimido1, Comprimido2

def delete_po(name):
  columns = defaultdict(list)
  word =  "#"
  with open(r"Z:\ESD Testing Reports\NAMES.TXT", 'r', encoding="ISO-8859-1") as f:
    reader = csv.reader(f, delimiter=',')
    for row in reader:
      for i in range(len(row)):
        columns[i].append(row[i])

# Following line is only necessary if you want a key error for invalid column numbers
  name0 = re.split(', |_|-|!', name.upper())

  columns = dict(columns)
  count = 0

  #print(name0[0])
  #print(name0[1])
  for i in range (len(columns[0])):

    if columns[0][i] == name0[0] and columns[0][i] == name0[1]:
      break
    count += 1
  add_list0 = []
  #print(count)
  index = count
  #del add_list0[index]

  a_file = open(r"Z:\ESD Testing Reports\NAMES.TXT", "r")

  lines = a_file.readlines()
  a_file.close()

  del lines[index]

  new_file = open(r"Z:\ESD Testing Reports\NAMES.TXT", "w+")

  for line in lines:
    new_file.write(line)

  new_file.close()

  return add_list0

def add(name, apellido, id, loc0, hor0, area0, cond, y2):
  columns = defaultdict(list)
  word =  "#"
  try:
      with open(r"Z:\ESD Testing Reports\NAMES.TXT", 'r', encoding="ISO-8859-1") as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
          for i in range(len(row)):
            columns[i].append(row[i])

    # Following line is only necessary if you want a key error for invalid column numbers
      name0 = name.upper()
      apellido0 = apellido.upper()
      id0 = id.upper()
      cond = " " +str(cond)
      columns = dict(columns)
      columns[0].append(name0)
      columns[1].append(apellido0)
      columns[2].append("EAST WEST-ID:"+ id)
      columns[3].append(" -1.000")
      columns[4].append(" -1.00")
      columns[5].append(" -1.000")
      columns[6].append(" -1.00")
      columns[7].append(cond)
      columns[8].append(" -1")
      columns[9].append(" A")
      columns[10].append(" 07/30/2020")
      columns[11].append(" 07/30/2020")
      columns[12].append(" 1")
      columns[13].append(" 07/30/2020")
      columns[14].append(" AAA")
      columns[15].append(" 0")
      columns[16].append(" ")
      columns[17] = [" 0"]*len(columns[16])
      add_list = [0]*len(columns[0])

      for i in range (len(columns[0])):
        add_list[i] = columns[0][i] + ", "+ columns[1][i] + ", " +columns[2][i] + "," + columns[3][i] + "," + columns[4][i] + "," + columns[5][i] + "," + columns[6][i] + "," + columns[7][i] + "," + columns[8][i] + "," + columns[9][i] + "," + columns[10][i] + "," + columns[11][i] + "," + columns[12][i] + "," + columns[13][i] + "," + columns[14][i] + "," +  columns[15][i] + "," + columns[16][i] + "," + columns[17][i]

      size = len(add_list) - 2
      final_len = len(add_list) - 1
      first = add_list[0]

      second = add_list[1]

      visitas = add_list[size]

      add_list.remove(first)
      add_list.remove(second)
      add_list.remove(visitas)
      #print(add_list[len(add_list) - 1])
      value = add_list[len(add_list) - 1]  + "\n"
      add_list0 = add_list
      add_list0.sort(key=itemgetter(0,1))

      first = "# On the following lines put last name, first name, ID, wrist min, wrist max, foot min, foot max, wrist enable (-1=enable 0=disable), foot enable, status (Aok Sick Training Vacation), Status start date (mm/dd/yyyy), Status stop date (mm/dd/yyyy), Certification Type, Certification expiration date (mm/dd/yyyy), string of 367 attendance chars, certification warning, password, updateflag"

      second = "# Example:  Gibson, Mel, 123-456:78, .075, 10.12, .05, 50.0, -1, 0, V 10/14/1997, 10/20/1997, 1, 10/20/2001, 367AAAA, 0, pass, 0"

      add_list0.insert(0, first)
      add_list0.insert(1, second)
      add_list0.insert(final_len, visitas)


      palabra = name0+', '+apellido0+', EAST WEST-ID:'+ id0 +', -1.000, -1.00, -1.000, -1.00,'+ cond +', -1, A, 07/30/2020, 07/30/2020, 1, 07/30/2020, AAA, 0, , 0'
      palabra = palabra.strip().replace(" ", "")
      #'HOLIS, HOLIS, EAST WEST-ID:holis, -1.000, -1.00, -1.000, -1.00, -1, -1, A, 07/30/2020, 07/30/2020, 1, 07/30/2020, AAA, 0, , 0'
      #HOLIS, HOLIS, EAST WEST-ID:holis
      print("Palabra en add function")
      print(palabra)
      print(type(palabra))
      hello = [x.strip(' ').replace(" ", "") for x in add_list0]
      #print("add_list0 en add function")
      #print(hello)
      #print(type(hello[0]))
      #palabra = 'ejemplo'
      #res = [(indice, string)for indice, string in enumerate(hello)  if palabra in string ]
      res = []
      #for count, value in enumerate(hello):
          #print(count, value)
          #      if palabra in value:
        #      print(count)
        #      res.append(count)
        #      break
      count = 0
      for i in range(len(hello)):
          if intern(palabra) == intern(hello[i]):
              print("HERE, AQUI ESTOY")
              break
          count+=1
      #print("Res in add function")
      #print(res)
      #print(res[0][0])
      bole = palabra in add_list0
      print("String in addlist0")
      print(bole)
      var = len(hello) -2
      print(hello[var])
      print(palabra)
      #index = res[0]
      index = count
      with open(r"Z:\ESD Testing Reports\NAMES.TXT", "r") as f:
        contents = f.readlines()

      contents.insert(index, value)

      with open(r"Z:\ESD Testing Reports\NAMES.TXT", "w") as f:
        contents = "".join(contents)
        f.write(contents)

      more_lines = "#" + "," + name.strip() + " " + apellido.strip() + "," + loc0 + "," + hor0 + "," + area0 + "," + id +"\n"
      print(more_lines)
      with open('ALL.TXT', 'a') as f:
          f.write(more_lines)

      #already(name, apellido, id)
      return add_list0
  except IOError:
      messagebox.showerror(message="Compruebe la conexion a internet", title="Advertencia")
      y2.destroy()



def lists(fecha, y2):
  #Counters
  count0 = 0
  count1 = 0
  aux_cont = 0

  columns = defaultdict(list)
  try:
      with open('Z:\ESD Testing Reports\ED23\LOG.TXT', 'r', encoding = "ISO-8859-1") as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
          for i in range(len(row)):
            columns[i].append(row[i])
    # Following line is only necessary if you want a key error for invalid column numbers
      columns = dict(columns)

    #Determinate who pass the test
    #columns = dict(columns)
      sum = len(columns[0])

      for i in range(sum):
        if columns[3][i] == fecha:
          if ((columns[5][i] == " Wok") and (columns[6][i] == " Lok") and (columns[7][i] == " Rok")) or ( (columns[5][i] == " W- ") and (columns[6][i] == " Lok") and (columns[7][i] == " Rok") ):
            count0 +=1
          else:
            count1 +=1

      dicts = {}
      keys = sum
      for i in range(keys):
        if columns[3][i] == fecha:
          #dicts[i] = columns[i]
          aux_cont +=1
    #List of names and nicknames people pass the test
      c0 = [0]*count0
      c1 = [0]*count0
      aux0 = 0
      aux1 = 0
    #List of names and nicknames people doesn't pass the test
      d0 = [0]*count1
      d1 = [0]*count1
      nick_name = []
      name = []
      Wok = []
      Lok = []
      Rok = []

      for i in range(sum):
        if columns[3][i] == fecha:
          nick_name.append(columns[0][i])
          name.append(columns[1][i])
          Wok.append(columns[5][i])
          Lok.append(columns[6][i])
          Rok.append(columns[7][i])
          if ((columns[5][i] == " Wok") and (columns[6][i] == " Lok") and (columns[7][i] == " Rok")) or ( (columns[5][i] == " W- ") and (columns[6][i] == " Lok") and (columns[7][i] == " Rok") ):

            c0[aux0] = columns[0][i]
            c1[aux0] = columns[1][i]
            aux0 +=1


          else:
            d0[aux1] = columns[0][i]
            d1[aux1] = columns[1][i]
            aux1 +=1
      #print(d0)
      df_nick_name=pd.DataFrame(nick_name)
      #print(len(nick_name))
      df_name=pd.DataFrame(name)
      #print(len(name))
      df_Wok=pd.DataFrame(Wok)
      #print(len(Wok))
      df_Lok=pd.DataFrame(Lok)
      #print(len(Lok))
      df_Rok=pd.DataFrame(Rok)
      #print(len(Rok))

    #Concatenate lists
      unite0 = [0]*aux0
      unite1 = [0]*aux1

    #ok
      for i in range(aux0):
        unite0[i] = c0[i] + c1[i]

    #not ok
      for i in range(aux1):
        unite1[i] = d0[i] + d1[i]




      return unite0, unite1, columns, df_nick_name, df_name, df_Wok, df_Lok, df_Rok
  except IOError:
      messagebox.showerror(message="Compruebe la conexion a internet", title="Advertencia")
      y2.destroy()


#Personas con test  negativo
def compare(unite0, unite1):
  listOne = set(unite0)
  listTwo = set(unite1)
  res = list(listTwo - listOne)
  return res

def matches (unite0,nick_name, name, Wok, Lok, Rok):
  unite2 = [0]*len(nick_name)
  for i in range(len(unite2)):
    unite2[i] = nick_name[i] + name[i]
  every0 = [x.upper() for x in unite2]
  count = 0
  count1 = 0
  for i in range(len(every0)):
    for j  in range(len(unite0)):
      if every0[i] == unite0[j] :
        count +=1
  match = [0]*count
  for i in range(len(every0)):
    for j  in range(len(unite0)):
      if every0[i] == unite0[j] :
        match[count1] = unite2[i] + " ->" + Wok[i] +","+ Lok[i] + ","+ Rok[i]
        count1 +=1
  return match

def all():
  #Counters
  count0 = 0
  count1 = 0

  columns = defaultdict(list)
  with open('ALL.TXT', 'r', encoding="utf8") as f:
      reader = csv.reader(f, delimiter=',')
      for row in reader:
          for i in range(len(row)):
              columns[i].append(row[i])
# Following line is only necessary if you want a key error for invalid column numbers
  columns = dict(columns)
  sum = len(columns[1])
  count0 = 0
  count1 = 0
  count2 = 0
  count3 = 0

  for i in range(sum):
    if (columns[2][i] == "Limpieza") or (columns[2][i] == "Miscelanea") or (columns[4][i] == "Compras") or (columns[4][i] == "Mantenimiento") or (columns[4][i] == "Logistica") or (columns[4][i] == "Recursos Humanos")  or (columns[2][i] == "Administracion"):
      count1 +=1
    elif columns[3][i] == "Comprimido 1":
      count0 +=1
    elif columns[3][i] == "Comprimido 2":
      count2 +=1
    else:
      count3 +=1

  unite0 = [0]*count0
  unite1 = [0]*count0
  compri2 = [0]*count2
  rest = [0]*count3
  aux = 0
  aux1 = 0
  aux2 = 0
  for i in range(sum):
    if (columns[2][i] == "Limpieza") or (columns[2][i] == "Miscelanea") or (columns[4][i] == "Compras") or (columns[4][i] == "Mantenimiento") or (columns[4][i] == "Logistica")or (columns[4][i] == "Recursos Humanos")  or (columns[2][i] == "Administracion"):
      p = 0
    elif columns[3][i] == "Comprimido 1":
      unite0[aux] = columns[1][i]
      aux +=1
    elif columns[3][i] == "Comprimido 2":
      compri2[aux1] = columns[1][i]
      aux1 += 1
    else:
      rest[aux2] = columns[1][i]
      aux2 += 1
  #Compri1
  unite1 = [x.upper() for x in unite0]
  #Compri2
  compri22 = [x.upper() for x in compri2]
  #Lo que no es comprimidos
  get = [x.upper() for x in rest]

  return unite1, columns, compri22, get

def print0 (data, not_came_compri1):
  data1 = [0]*len(data[1])
  match = [0]*len(not_came_compri1)
  count = 0
  for i in range(len(data[1])):
    data1[i] = data[1][i].strip()
  data0 = [x.upper() for x in data1]
  for i in range (len(data[1])):
    for j in range (len(not_came_compri1)):
      if not_came_compri1[j] == data0[i]:
        match[count] = data[1][i] + "->" + data[2][i] + "," + data[3][i] + "," + data[4][i]
        count +=1

  return match
