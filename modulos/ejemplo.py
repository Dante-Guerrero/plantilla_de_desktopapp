from apoyo.elemetos_de_GUI import Cuadro, Ventana

class Ventana1(Ventana):
    """"""
    
    #----------------------------------------------------------------------
    def __init__(self, *args):
        """Constructor"""

        Ventana.__init__(self, *args)

        c1 = Cuadro(self)
        c1.agregar_button(0,0,'Ir a ventana 2', self.ir)
    
    #----------------------------------------------------------------------
    def ir(self):
        """"""
        self.desaparecer()
        subframe = Ventana2(self, 500, 200, 'Ventana 2')

class Ventana2(Ventana):

    #----------------------------------------------------------------------
    def __init__(self, *args):
        """Constructor"""

        Ventana.__init__(self, *args)

        c1 = Cuadro(self)
        c1.agregar_button(0,0,'Volver a ventana 1', self.volver)