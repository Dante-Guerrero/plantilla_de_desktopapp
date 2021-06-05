from tkinter import Tk
from modulos import ejemplo as ej

########################################################################
class Aplicacion(object):
    """"""

    #----------------------------------------------------------------------
    def __init__(self, parent):
        """Constructor"""

        self.root = parent
        self.root.withdraw()
        subFrame = ej.Ventana1(self, 500, 500, "Ventana 1")

#----------------------------------------------------------------------
def main():
    root = Tk()
    app = Aplicacion(root)
    root.mainloop()

