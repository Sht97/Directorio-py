# import os
from tkinter.ttk import Treeview
from tkinter import filedialog, messagebox
from tkinter import *


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

def agregar():
    name = StringVar()
    tel = StringVar()
    email = StringVar()
    root2=Tk()
    root2.geometry("400x400")
    root2.configure(background = "white")
    nameBox = Entry(root2, textvariable = name).place(x = 150, y = 60)
    nameLabel = Label(root2, text = "Nombre: ", bg = "white", fg = "black").place(x = 50, y =60)
    telBox = Entry(root2, textvariable = tel).place(x = 150, y = 90)
    telLabel = Label(root2, text = "Teléfono: ", bg = "white", fg = "black").place(x = 50, y =90)
    emailBox = Entry(root2, textvariable = email).place(x = 150, y = 120)
    emailLabel = Label(root2, text = "Email: ", bg = "white", fg = "black").place(x = 50, y =120)
    botonGuardar = Button(root2, text = "Guardar", bg = "#CFD4F0", fg ="black").place(x=150,y = 250)
    #root.destroy()    




def LoadTable():
    tv.insert('', 'end', text="Danielito", values=("454648",45445,"Danielasd@"))
    tv.insert('', 'end', text="Karolsita", values=("allalalla",'meh', 'sadk'))


root = Tk()
root.title("Otalparo")
root.geometry("500x500")
root.configure(background = "white")
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
botonAgregar = Button(root, text = "Agregar", command = agregar,
    bg = "#CFD4F0", fg ="black").place(x=100,y=400)

LoadTable()
root.config(menu=menubar)
root.mainloop()
