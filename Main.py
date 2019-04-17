import tkinter as tk
from tkinter import filedialog, messagebox, Menu, StringVar, Entry, Label, Button, Text
from tkinter.ttk import Treeview

import openpyxl


def new():
    global savePath
    savePath = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    a = open('Config/path.txt', 'w')
    a.write(savePath)
    a.close()
    deleteALl()


def guardar():
    global savePath
    wb = openpyxl.Workbook()
    hoja = wb.active
    hoja.title = "Hoja 1"
    fila = 1  # Fila donde empezamos
    for i in tv.get_children():
        hoja.cell(column=1, row=fila, value=tv.item(i)['text'])
        hoja.cell(column=2, row=fila, value=tv.item(i)['values'][0])
        hoja.cell(column=3, row=fila, value=tv.item(i)['values'][1])
        hoja.cell(column=4, row=fila, value=tv.item(i)['values'][2])
        hoja.cell(column=5, row=fila, value=tv.item(i)['values'][3])
        fila += 1

    wb.save(filename=savePath)


def add():
    root2.deiconify()
    btn.place_forget()
    btn2.place(x=110, y=230)


def edit():
    if tv.focus():

        selected = tv.item(tv.focus())
        name.set(selected['text'])
        lastName.set((selected['values'][0]))
        tel.set((selected['values'][1]))
        cel.set((selected['values'][2]))
        email.set((selected['values'][3]))
        btn2.place_forget()
        btn.place(x=110, y=230)
        root2.deiconify()
    else:
        messagebox.showinfo("Error", "Ninguna persona ha sido seleccionada.")


def addEdit():
    if (name.get() == "" or lastName.get() == "" or cel.get() == "" or tel.get() == "" or email.get == ""):
        messagebox.showinfo("Error", "Debe llenar todos los datos.")
    else:
        tv.item(tv.focus(), text=name.get(), values=[lastName.get(), tel.get(), cel.get(), email.get()])
        root2.iconify()
        btn.place_forget()
        cleanBox()


def deleteALl():
    for i in tv.get_children():
        global aux
        aux -= 1
        tv.delete(i)


def delete():
    if tv.focus():
        global aux
        aux -= 1
        tv.delete(tv.focus())

    else:
        messagebox.showinfo("Error", "Ninguna persona ha sido seleccionada.")


def loadTable():
    if (name.get() == "" or lastName.get() == "" or cel.get() == "" or tel.get() == "" or email.get == ""):
        messagebox.showinfo("Error", "Debe llenar todos los datos.")
    else:
        global aux
        aux += 1
        tv.insert('', 'end', text=name.get(), values=(lastName.get(), tel.get(), cel.get(), email.get()))
        root2.iconify()
        cleanBox()
        print(aux)


def cleanBox():
    name.set("")
    lastName.set("")
    tel.set("")
    cel.set("")
    email.set("")


def abrir():
    global savePath
    savePath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    a = open('Config/path.txt', 'w')
    a.write(savePath)
    a.close()
    deleteALl()
    global columnas
    columnas = list(openpyxl.load_workbook(savePath)["Hoja 1"])
    append()


def append():
    print(tv.get_children())
    for i in range(len(columnas)):
        tv.insert('', 'end', text=columnas[i][0].value,
                  values=(columnas[i][1].value, columnas[i][2].value, columnas[i][3].value, columnas[i][4].value))


f = open('Config/path.txt', 'r')
savePath = f.read()
f.close()
columnas = list(openpyxl.load_workbook(savePath)["Hoja 1"])
print(savePath)
aux = len(columnas)

root = tk.Tk()
menubar = Menu(root)
text = Text(root)
name = StringVar()
lastName = StringVar()
tel = StringVar()
email = StringVar()
cel = StringVar()

root.title("Agenda")
root.geometry("720x300")
root.configure(background="white")

root2 = tk.Toplevel(root)
root2.geometry("280x300")
root2.configure(background="white")
root2.iconify()

filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="New", command=new)
filemenu.add_command(label="Open", command=abrir)
filemenu.add_command(label="Save", command=guardar)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)
filemenu.add_separator()

Label(root2, text="AGREGAR CONTACTO", bg="white", fg="black").place(x=80, y=20)
Entry(root2, textvariable=name).place(x=120, y=60)
Label(root2, text="Nombres: ", bg="white", fg="black").place(x=30, y=60)
Entry(root2, textvariable=lastName).place(x=120, y=90)
Label(root2, text="Apellidos: ", bg="white", fg="black").place(x=30, y=90)
Entry(root2, textvariable=tel).place(x=120, y=120)
Label(root2, text="Teléfono: ", bg="white", fg="black").place(x=30, y=120)
Entry(root2, textvariable=cel).place(x=120, y=150)
Label(root2, text="Celular: ", bg="white", fg="black").place(x=30, y=150)
Entry(root2, textvariable=email).place(x=120, y=180)
Label(root2, text="Email: ", bg="white", fg="black").place(x=30, y=180)
btn = Button(root2, text="Guardar", command=addEdit, bg="#CFD4F0", fg="black")
btn2 = Button(root2, text="Guardar", command=loadTable, bg="#CFD4F0", fg="black")

tv = Treeview()
tv['columns'] = ('Apellidos', 'Telefono Fijo', 'Celular', 'Correo electronico')
tv.heading("#0", text='  Nombres', anchor='w')
tv.column("#0", anchor="w", width=150)
tv.heading('Apellidos', text='  Apellidos', anchor='w')
tv.column('Apellidos', anchor="w", width=150)
tv.heading('Telefono Fijo', text='Teléfono Fijo')
tv.column('Telefono Fijo', anchor='center', width=100)
tv.heading('Celular', text='Celular')
tv.column('Celular', anchor='center', width=120)
tv.heading('Correo electronico', text='Correo electrónico')
tv.column('Correo electronico', anchor='center', width=180)
tv.grid(sticky=("N", "S", "W", "E"))
tv.place(x=10, y=10)

append()

Button(root, text="Agregar", command=add, bg="#CFD4F0", fg="black").place(x=100, y=250)
Button(root, text="Editar", command=edit, bg="#CFD4F0", fg="black").place(x=200, y=250)
Button(root, text="Eliminar", command=delete, bg="#CFD4F0", fg="black").place(x=300, y=250)

root.config(menu=menubar)
root.mainloop()
