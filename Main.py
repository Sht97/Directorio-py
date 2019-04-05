# import os
from tkinter import filedialog, messagebox
from tkinter import *

# root = Tk()
#
# def fichero(haylibreta):
#
#     if haylibreta:
#         directorio = filedialog.askopenfilename(initialdir=os.getcwd(), title="Seleccione el archivo para guardar",
#                                                    filetypes=(
#                                                        ("Archivos de texto", "*.txt"),
#                                                        ("Solo archivos de texto", "*.txt*")))
#
#     else:
#         directorio = filedialog.asksaveasfilename(initialdir=os.getcwd(), title="Seleccione el archivo para guardar",
#                                                      filetypes=(
#                                                          ("Archivos de texto", "*.txt"),
#                                                          ("Solo archivos de texto", "*.txt*")))
#
#
#
#
#
# class Contacto:
#
#     def __init__(self, nombre_completo, apodo, telefono, archivo):
#         self.nombre_Completo = nombre_completo
#         self.apodo = apodo
#         self.telefono = telefono
#         self.archivo = archivo
#     # def guardarContacto(self):

# fichero(messagebox.askyesno( "Responda","Ya tiene un directorio creado ?"))
# # nombre_completo = input("Introduce nombre")
# # apodo = input("Introduce apodo")
# # telefono = input("Telefono")
# # persona = Contacto(nombre_completo, apodo, telefono, 4)
from tkinter.ttk import Treeview


def donothing():
    print("n")


def abrir():
    f = filedialog.askopenfile(mode="r")
    linea = f.readline()
    if linea:
        while linea:
            if linea[-1] == "\n":
                linea = linea[:-1]
            print(linea)
            contactos.append(linea)
            linea = f.readline()
    f.close()


def guardar():
    f = filedialog.asksaveasfile(mode='w', defaultextension=".txt")
    if f is None:  # asksaveasfile return `None` if dialog closed with "cancel".
        return
    text2save = str(("Holiiis" + "\n" + "holita"))  # starts from `1.0`, not `0.0`
    f.write(text2save)
    f.close()


contactos = []
root = Tk()
root.geometry("500x500")
menubar = Menu(root)
text = Text(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="New", command=donothing)
filemenu.add_command(label="Open", command=abrir)
filemenu.add_command(label="Save", command=guardar)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)
filemenu.add_separator()
tv = Treeview()
tv['columns'] = ('Telefono Fijo', 'Celular', 'Correo electronico')
tv.heading("#0", text='Nombre completo', anchor='w')
tv.column("#0", anchor="w")
tv.heading('Telefono Fijo', text='Telefono Fijo')
tv.column('Telefono Fijo', anchor='center', width=100)
tv.heading('Celular', text='Celular')
tv.column('Celular', anchor='center', width=100)
tv.heading('Correo electronico', text='Correo electronico')
tv.column('Correo electronico', anchor='center', width=100)
tv.grid(sticky=(N, S, W, E))


def LoadTable():
    tv.insert('', 'end', text="Danielito", values=("454648",45445,"Danielasd@"))
    tv.insert('', 'end', text="Karolsita", values=("allalalla",'meh', 'sadk'))

LoadTable()
root.config(menu=menubar)
root.mainloop()
