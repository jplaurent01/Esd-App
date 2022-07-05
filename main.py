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

n_array = []
count_aux0 = 0
root = tk.Tk()
root.title('Menu ESD')
root.geometry("800x800")
bg = PhotoImage( file = "electronic1.png")
#root.configure(borderwidth="1")
#root.configure(relief="sunken")
#root.configure(cursor="arrow")
#root.configure(bg='#116562')
label1 = Label( root, image = bg)
label1.place(x = 0,y = 0)

#root.configure(highlightcolor="black")

def img():
    photo = tk.PhotoImage(file = r"images_calendar.png")
    # Resizing image to fit on button
    photoimage = photo.subsample(10, 10)
    return photoimage

def Whatsapp():
  y2 = Frame()
  #y2.place(x=0, y=0, width=2500, height=2500)
  y2.place(x=0, y=0, width=650, height=100)
  aux = cal.get_date().replace('/','-')
  oldDate = aux.split('-')
  if len(oldDate[0]) < 2:
    oldDate[0]= '0' + oldDate[0]
  if len(oldDate[1]) < 2:
    oldDate[1]= '0' + oldDate[1]

  fecha0 = '-'.join(oldDate)

  #date.config(text = "Selected Date is: " + newDate)

  fecha = " " + fecha0
  #print(fecha)

  lista = funct.whatsapp_list(fecha, y2)
  #comprimidos1
  comprimido1 = lista[0]

  #comprimido2
  comprimido2 = lista[1]

  #Todos los faltantes
  eve_tot = lista[2]

  btn0 = tk.Button(y2, text = "Enviar mensaje con comprimido1", bg="#116562", fg='#f7fafa',activebackground='#055959',
  activeforeground='#f7fafa', command=lambda: funct.send_message(comprimido1,fecha)).grid(row = 1, column = 0, pady=1, padx=3)

  btn0 = tk.Button(y2, text = "Enviar mensaje con comprimido2", bg="#116562", fg='#f7fafa',activebackground='#055959',
  activeforeground='#f7fafa',command=lambda:funct.send_message(comprimido2, fecha)).grid(row = 1, column = 1, pady=1, padx=3)

  btn0 = tk.Button(y2, text = "Enviar mensaje con todos los faltantes ", bg="#116562", fg='#f7fafa',activebackground='#055959',
  activeforeground='#f7fafa',command=lambda:funct.send_message(eve_tot, fecha)).grid(row = 1, column = 2, pady=1, padx=3)
  #funct.send_message(comrimido2)

  tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=0, column=0, pady=10, padx=10)

def delete_personal():

  y2 = Frame()
  #y2.place(x=0, y=0, width=2500, height=2500)
  y2.place(x=0, y=0, width=500, height=500)

  lstA = funct.read_file("ALL.TXT")
  all_lst = []
  for i in range(len(lstA[1])):
      all_lst.append(lstA[1][i] + "-" + lstA[2][i] + "-" + lstA[3][i] + "-" + lstA[4][i])
  valores2 = StringVar()
  valores2.set(all_lst)

  lab2 = Label(y2,text="Lista de personal")
  lab2.grid(row = 1, column = 1)
  lstbox2 = Listbox(y2, listvariable=valores2, selectmode=MULTIPLE, width=55, height=20)
  lstbox2.grid(column=1, row=2, pady=1, padx=5)
      #return lst
  #data = todos_dispay()
  #print(data)
        #funct.outlook (Fail0, lst)
  #m = Menu(y2, tearoff = 0)
  #m.add_command(label ="Eliminar", command=funct.Eliminate(lstbox2))
  #m.add_command(label ="Editar", command=funct.Edit(lstbox2))
  #m.add_command(label ="Paste")
  #m.add_command(label ="Reload")
  #m.add_separator()
  #m.add_command(label ="Rename")

  #def do_popup(event):
#      try:
#          m.tk_popup(event.x_root, event.y_root)
#      finally:
#          m.grab_release()

 # lstbox2.bind("<Button-3>", do_popup)
  tk.Button(y2, text="Eliminar", command=lambda:funct.tod_dis(lstbox2,y2)).grid(row=4, column=1, pady=10, padx=10)
  tk.Button(y2, text="Editar", command=lambda:funct.edit_tod_dis(lstbox2,y2)).grid(row=5, column=1, pady=10, padx=10)
  tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=0, column=0, pady=10, padx=10)


def getval(tbn, value_inside, value_inside1, value_inside2, value_inside3, y2):
    data = []
    loc = str(value_inside.get())
    hor = str(value_inside1.get())
    area = str(value_inside2.get())
    trj = str(value_inside3.get())
    #print(loc)
    #print(hor)
    #print(area)
    print("len(n_array):")
    print(len(n_array))
    aux_var = count_aux0 - 1
    #if len(n_array) == 1:
    #    print(n_array[0][0].get())
    #    print(n_array[0][1].get())
    #    print(n_array[0][2].get())
    #else :
    #    print(n_array[1][0].get())
    #    print(n_array[1][1].get())
    #    print(n_array[1][2].get())
    for row in range(len(n_array)):
        print(aux_var)
        if aux_var == row:
            for col in range(3):
                print(n_array[row][col].get())
                data.append(n_array[row][col].get())
        else:
            pass
    #print("List data")
    #print(data)
    if len(data[0]) == 0 or len(data[1]) == 0 or len(data[2]) == 0:
        messagebox.showerror(message="Rellene espacios en blanco", title="Adveretencia")
    else:
        if loc == "Localizacion" or hor == "Horario" or area == "Area" or trj == "Manipulacion de tarjetas":
            messagebox.showerror(message="Ingrese una opcion valida", title="Adveretencia")

        else:
            print("data[1]",data[1].replace(" ", ""))
            print("data[0]",data[0].replace(" ", ""))
            count_data1 = funct.count_space(data[1].strip())
            count_data0 = funct.count_space(data[0].strip())
            print("data[1].isalpha()",data[1].isalpha())
            print("data[0].isalpha()",data[0].isalpha())
            print("count_data1",count_data1 )
            print("count_data0", count_data0)
            print("funct.ene(data[1].strip()) ", funct.ene(data[1].strip()) )
            print("funct.ene(data[0].strip()) ", funct.ene(data[0].strip()) )
            print("funct.ene(data[2].strip())", funct.ene(data[2].strip()) )
            if data[1].replace(" ", "").isalpha() == True and data[0].replace(" ", "").isalpha() == True  and count_data1 <2 and count_data0 <2 and funct.ene(data[1].strip()) == True and funct.ene(data[0].strip()) == True and funct.ene(data[2].strip()) == True and funct.find_latin(data[1].strip()) == True and funct.find_latin(data[0].strip()) == True and funct.find_latin(data[2].strip()) == True:
                view = funct.already(data[1], data[0], data[2],y2)

                if view == 0:
                    if trj == "Si" :
                        cond = -1
                        result = funct.add(data[1], data[0], data[2], loc, hor, area, cond, y2)
                        messagebox.showinfo("Information","Usuario agregado exitosamente")
                    elif trj == "No" :
                        cond = 0
                        result = funct.add(data[1], data[0], data[2], loc, hor, area, cond, y2)
                        messagebox.showinfo("Information","Usuario agregado exitosamente")
                elif view == 1:
                    messagebox.showerror(message="Ingrese una contraseña distinta", title="Adveretencia")
                else:
                    messagebox.showerror(message="Usuario ya existe", title="Adveretencia")
            else:
                messagebox.showerror(message="Ingrese unicamente letras, como máximo 2 espacios por texto y no ingrese la letra ñ ni tildes", title="Advertencia")
        #tbn.delete(0, END)
        #funct.fail0_show_selection(y2, result)
        #n_array = []
def personal():
  y2 = Frame()
  #y2.place(x=0, y=0, width=2500, height=2500)
  y2.place(x=0, y=0, width=500, height=500)
  #y2.geometry("800x800")
  #bg2 = PhotoImage( file = "electronic1.png")
  #label2 = Label( y2, image = bg2)
  #label2.place(x = 0,y = 0)
#global frame, n_array
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
        tbn=Entry(y2)
        row_array.append(tbn)
        row_array[x].grid(row=pos, column=2,sticky="nsew", padx=2,pady=2)

      if x==1:
        pos = x + 1
        label1 = Label(y2,text="Apellidos")
        label1.grid(row = pos, column = 1)
        tbn=Entry(y2)
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
activeforeground='#f7fafa', command=lambda:getval(tbn, value_inside, value_inside1, value_inside2, value_inside3, y2)).grid(row=12, column=2)

  #photo20 = tk.PhotoImage(file = r"return.PNG")
  # Resizing image to fit on button
  #photoimage20 = photo20.subsample(10, 10)
  tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=0, column=0, pady=10, padx=10)

def button():
  y2 = Frame()
  y2.place(x=0, y=0, width=1500, height=1000)
  aux = cal.get_date().replace('/','-')
  oldDate = aux.split('-')
  if len(oldDate[0]) < 2:
    oldDate[0]= '0' + oldDate[0]
  if len(oldDate[1]) < 2:
    oldDate[1]= '0' + oldDate[1]

  fecha0 = '-'.join(oldDate)


  fecha = " " + fecha0
  #print(fecha)
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

  valores0 = StringVar()
  valores0.set(Fail0)

  valores1 = StringVar()
  valores1.set(Comprimido1)

  valores2 = StringVar()
  valores2.set(Comprimido2)

  valores3 = StringVar()
  valores3.set(tot_pers)

  lab0 = Label(y2,text="No pasa test")
  lab0.grid(row = 0, column = 3)
  lstbox2 = Listbox(y2, listvariable=valores0, selectmode=MULTIPLE, width=45, height=20)
  #lstbox2.grid(column=3, row=1, columnspan=1)
  lstbox2.grid(column=3, row=1, pady=1, padx=5)
  lab3 = Label(y2,text="Todo el personal")
  lab3.grid(row = 0, column = 2)
  lstbox3 = Listbox(y2, listvariable=valores3, selectmode=MULTIPLE, width=45, height=20)
  lstbox3.grid(column=2, row=1, pady=1, padx=5)

  lab1 = Label(y2,text="Comprimido 1")
  lab1.grid(row = 0, column = 0)
  lstbox1 = Listbox(y2, listvariable=valores1, selectmode=MULTIPLE, width=45, height=20)
  lstbox1.grid(column=0, row=1, pady=1, padx=5)
  #scrollbar = Scrollbar(y2)
  #scrollbar.grid(column=0, row=1)
  #lstbox1.config(yscrollcommand = scrollbar.set)
  #scrollbar.config(command = lstbox1.yview)

  lab2 = Label(y2,text="Comprimido 2")
  lab2.grid(row = 0, column = 1)
  lstbox2 = Listbox(y2, listvariable=valores2, selectmode=MULTIPLE, width=45, height=20)
  lstbox2.grid(column=1, row=1, pady=1, padx=5)

  def tick():
      datenow = datetime.datetime.now()
      time_string = datenow.strftime("%d-%m-%Y %H:%M:%S:%p")
      clock.config(text=time_string)
      clock.after(200,tick)

  def Comprimido2_display():
      reslist = list()
      seleccion = lstbox2.curselection()
      lst = []
      for i in seleccion:
          entrada = lstbox2.get(i)
          reslist.append(entrada)
      for val in reslist:
          #print(val)
          lst.append(val)
      print(lst)
      funct.outlook (Fail0, lst)

      #Comprimido2_display = scrolledtext.ScrolledText(y2, height=10, width=50)

      #for i in Comprimido2:
          #Comprimido2_display.insert(END, (i + "\n"))

      #Comprimido2_display.grid(row=5, column=1,padx=10)

  def Comprimido1_dispay():
      reslist = list()
      seleccion = lstbox1.curselection()
      lst = []
      for i in seleccion:
          entrada = lstbox1.get(i)
          reslist.append(entrada)
      for val in reslist:
          #print(val)
          lst.append(val)
      print(lst)
      funct.outlook (Fail0, lst)


  def todos_dispay():
      reslist = list()
      seleccion = lstbox3.curselection()
      lst = []
      for i in seleccion:
          entrada = lstbox3.get(i)
          reslist.append(entrada)
      for val in reslist:
          #print(val)
          lst.append(val)
      print(lst)
      funct.outlook (Fail0, lst)


  def fail0_dispay():

      fail0_dispay = scrolledtext.ScrolledText(y2, height=10, width=50)

      for i in Fail0:
        if len(Fail0) == 0:
          fail0_dispay.insert(END, ("No se encuentran irregularidades"))
        else:
          fail0_dispay.insert(END, (i + "\n"))


      fail0_dispay.grid(row=5,column=1,padx=10)

  def fail0_show_selection():
      for pa in Fail0:
          Fail0_display = Text(master=y2, height=20, width=60, bg="Lightgreen",font="Helvetica 10 bold")
          Fail0_display.grid(row=5, column=1, padx=10)
          Fail0_display.insert(END, pa)

  def fail0_none_selected():
        # Display message if no options selected
      greeting1 = " "
      Fail0_display = scrolledtext.ScrolledText(y2, height=10, width=50)

      Fail0_display.grid(row=5, column=1, padx=10)
      Fail0_display.insert(END, greeting1)


  #myradiobutton2 = tk.Radiobutton(y2, text="ESD fallido", font=("Helvetica", 12),variable=f, value=Fail0,command=fail0_dispay).grid(row=2, column=0, sticky=W)

  #tk.Radiobutton(y2, text="Personal de comprimido 1", font=("Helvetica", 12), variable=e, value=Comprimido1, command=Comprimido1_dispay).grid(row=3, column=0)

  #tk.Radiobutton(y2, text="Personal de comprimido 2", font=("Helvetica", 12), variable=e, value=Comprimido2, command=Comprimido2_display).grid(row=4, column=0)

  btn_1 = ttk.Button(y2, text="Correo para comprimido 1", command=Comprimido1_dispay)
  btn_1.grid(column=0, row=5)

  btn_2 = ttk.Button(y2, text="Correo para comprimido 2", command=Comprimido2_display)
  btn_2.grid(column=1, row=5)

  btn_2 = ttk.Button(y2, text="Correo para todos", command=todos_dispay)
  btn_2.grid(column=2, row=5)
  #selection_button1 = tk.Button(y2, text="Limpiar ventanas",bg="#116562", fg='#f7fafa',activebackground='#055959',
#activeforeground='#f7fafa', command=fail0_none_selected)
  #selection_button1.grid(row=5, column=0, pady=10, padx=100)

  tk.Button(y2, text='Regresar', width=10, bg="black", fg='white', command=lambda:[y2.destroy()]).grid(row=7, column=2, pady=10, padx=10)


f = IntVar()
e = IntVar()

datenow = datetime.datetime.now()
Year = int(datenow.strftime("%Y"))
Month = int(datenow.strftime("%m"))
Day = int(datenow.strftime("%d"))

cal = Calendar(root, selectmode = 'day', year = Year, month = Month,
 day = Day, date_pattern="mm-dd-yyyy")

cal.grid(row = 3, column = 1, padx=10, pady=15)


#date = Label(root, text = "")
#date.grid(row = 3, column = 1)

photo1 = tk.PhotoImage(file = r"img_add.png")
# Resizing image to fit on button
photoimage1 = photo1.subsample(10, 10)

btn0 = tk.Button(root, text = "Agregar personal ESD ", image = photoimage1,
                    compound = RIGHT,bg="#116562", fg='#f7fafa',activebackground='#055959',
activeforeground='#f7fafa',
command = personal).grid(row = 1, column = 2,  padx=10, pady=10)


photo2 = tk.PhotoImage(file = r"img_delete.png")
# Resizing image to fit on button
photoimage2 = photo2.subsample(10, 10)
btn1 = tk.Button(root, text = "Eliminar y editar personal ESD ", image = photoimage2,
                    compound = RIGHT, bg="#116562", fg='#f7fafa',activebackground='#055959',
activeforeground='#f7fafa',command = delete_personal).grid(row = 1, column = 3,  padx=10, pady=10)

photo3 = tk.PhotoImage(file = r"img_teams.png")
# Resizing image to fit on button
photoimage3 = photo3.subsample(10, 10)

btn2 = tk.Button(root, text = "Mensaje por Teams ", image = photoimage3,
                    compound = RIGHT,bg="#116562", fg='#f7fafa', activebackground='#055959',
activeforeground='#f7fafa',command = Whatsapp).grid(row = 1, column = 4,  padx=10, pady=10)

photo44 = tk.PhotoImage(file = r"pareto.png")
# Resizing image to fit on button
photoimage44 = photo44.subsample(10, 10)

btn22 = tk.Button(root, text = "Pareto Calidad", image = photoimage44,
                    compound = RIGHT,bg="#116562", fg='#f7fafa', activebackground='#055959',
activeforeground='#f7fafa',command =lambda:calls.open_file()).grid(row = 1, column = 5,  padx=10, pady=10)

#photo0 = tk.PhotoImage(file = r"images_calendar.png")
# Resizing image to fit on button
#photoimage0 = photo0.subsample(10, 10)

btn3 = tk.Button(root, text = "Informe por Fecha ",bg="#116562", fg='#f7fafa',activebackground='#055959',
activeforeground='#f7fafa',command = button).grid(row = 1, column = 1)
root.mainloop()
