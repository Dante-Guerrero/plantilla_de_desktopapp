import pandas as pd
from apoyo.elemetos_de_GUI import Ventana, Cuadro
from tkinter import *

def printy(x):
    print(x)

elementos = {
    'Pokemon':['Pikachu', 'Squirtle', 'Bulbasaur', 'Charmander'],
    'Tipo':['Electrico', 'Agua', 'Planta', 'Fuego'],
    'Tamaño':['Grande', 'Grande','Mediano', 'Pequeño'],
    'Entrenador':['Ash Ketchup', 'Gary Oak', 'Brenda Gonzales Orbegoso', 'Irene Guerrero'],
    'Poder especial':['Tueno', 'Chorro de agua', 'Hojas navaja', 'Lanzallamas']
}
tabla = pd.DataFrame(elementos)

view = Tk()

c1 = Cuadro(view, True)
c1.agregar_escenario(0, 0, tabla, printy, printy, printy)

view.mainloop()
