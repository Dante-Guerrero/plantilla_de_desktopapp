from tkinter import *
from apoyo.elemetos_de_GUI import Ventana, Cuadro

view = Tk()

c1 = Cuadro(view)

lista = ['Hola', 'Hello', 'Hi']
rejilla = (
    ('L', 0,0,'HOLA'),
    ('L',1,0, 'Irene Aissa'),
    ('E',2,0),
    ('CX',3,0,lista)
)
c1.agregar_rejilla(rejilla)

view.mainloop()