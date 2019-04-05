from tkinter.ttk import Treeview
from tkinter import filedialog, messagebox
from tkinter import *
import Main 
def armar():
	root = Tk()
	root.geometry("500x500")
	menubar = Menu(root)
	text = Text(root)
	filemenu = Menu(menubar, tearoff=0)
	filemenu.add_command(label="New", command=Main.donothing)
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
	    bg = "#009", fg ="green").place(x=100,y=400)