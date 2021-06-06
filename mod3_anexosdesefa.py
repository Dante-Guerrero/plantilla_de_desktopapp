# Importación externa
import os
import tkinter as tk
from tkinter import *
from tkinter import ttk
import tkinter.messagebox as tkmb
from tkinter.filedialog import askopenfilename, askdirectory
import gspread
from datetime import date, datetime
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from PIL import Image,ImageTk
import pandas as pd
# Importación local
from herramientas import funciones as func
from herramientas import globales as glob
from herramientas import mod1_acceso as acceso

########################################################################
class Demo2(tk.Toplevel):

    #----------------------------------------------------------------------
    def __init__(self, original):
        """Constructor"""
        
        self.original_frame = original
        tk.Toplevel.__init__(self)
        self.box_w = 520
        self.box_h = 200
        self.box_sw = self.winfo_screenwidth()
        self.box_sh = self.winfo_screenheight()
        self.box_x = (self.box_sw - self.box_w)/2
        self.box_y = (self.box_sh - self.box_h)/2
        self.geometry('%dx%d+%d+%d' % (self.box_w, self.box_h, self.box_x, self.box_y))
        self.titulo = "Aplicativo: Herramientas de Sefa - versión " + glob.numero_de_version
        self.title(self.titulo)
        self.resizable(False,False)
        self.config(background= glob.color_fondo)
        self.frame = tk.Frame(self)
        self.frame.config(background= glob.color_fondo)
        self.frame.pack()
        
        if glob.comprobar_coincidencia == "sí":
            
            # Textos

            self.consulta = "Documento consultado: " + glob.nombre_completo
            self.mensaje = "Anteriormente ya se ha creado una carpeta de anexos para el documento consultado."
            self.pregunta = "¿Desea revisar y/o modificar su contenido?"
            self.consulta_label = Label(self.frame, text=self.consulta, fg=glob.color_letras, bg=glob.color_fondo)
            self.mensaje_label = Label(self.frame, text=self.mensaje, fg=glob.color_letras, bg=glob.color_fondo)
            self.pregunta_label = Label(self.frame, text=self.pregunta, fg=glob.color_letras, bg=glob.color_fondo)

            # Botones

            self.boton_si = Button(self.frame, text="Sí", command=self.openFrame, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
            self.boton_no = Button(self.frame, text="No", command=self.volver, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

            # Ubicaciones
            
            self.consulta_label.grid(row=1, column=1, columnspan=2, pady= 10, padx=20)
            self.mensaje_label.grid(row=2, column=1, columnspan=2, pady= 10, padx=20)
            self.pregunta_label.grid(row=3, column=1, columnspan=2, pady= 5, padx=20)
            self.boton_si.grid(row=4, column=1, pady= 10, padx=20)
            self.boton_no.grid(row=4, column=2, pady= 10, padx=20)

            # Efectos

            func.Efecto_de_boton(self.boton_si)
            func.Efecto_de_boton(self.boton_no)

        else:

            # Textos

            self.consulta = "Documento consultado: " + glob.nombre_completo
            self.mensaje = "Aún no se ha creado una carpeta de anexos para el documento consultado."
            self.pregunta = "¿Desea crearla?"
            self.consulta_label = Label(self.frame, text=self.consulta, fg=glob.color_letras, bg=glob.color_fondo)
            self.mensaje_label = Label(self.frame, text=self.mensaje, fg=glob.color_letras, bg=glob.color_fondo)
            self.pregunta_label = Label(self.frame, text=self.pregunta, fg=glob.color_letras, bg=glob.color_fondo)

            # Botones

            self.boton_si = Button(self.frame, text="Sí", command=self.crear_carpeta, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
            self.boton_no = Button(self.frame, text="No", command=self.volver, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

            # Ubicaciones
            
            self.consulta_label.grid(row=1, column=1, columnspan=2, pady= 10, padx=20)
            self.mensaje_label.grid(row=2, column=1, columnspan=2, pady= 10, padx=20)
            self.pregunta_label.grid(row=3, column=1, columnspan=2, pady= 5, padx=20)
            self.boton_si.grid(row=4, column=1, pady= 10, padx=20)
            self.boton_no.grid(row=4, column=2, pady= 10, padx=20)

            # Efectos

            func.Efecto_de_boton(self.boton_si)
            func.Efecto_de_boton(self.boton_no)

    #----------------------------------------------------------------------
    def volver(self):
        """"""
        self.destroy()
        self.original_frame.show()

    #----------------------------------------------------------------------
    def openFrame(self):
        """"""
        self.withdraw()
        subFrame = Demo3(self)

    #----------------------------------------------------------------------
    def crear_carpeta(self):
        """"""

        self.Carpeta_creada = glob.drive.CreateFile({'title': glob.nombre_completo,'mimeType':"application/vnd.google-apps.folder", 'parents':[{'id':glob.carpeta_de_anexos_de_sefa}]})
        self.Carpeta_creada.Upload()
        glob.ID = self.Carpeta_creada['id']

        self.ahora = str(datetime.now())
        self.new = [glob.usuario, self.ahora, glob.tipo_data, glob.numero_data, glob.ano_data, glob.terminacion_data, glob.nombre_completo, glob.ID]
        glob.worksheet_documentos.append_row(self.new)

        self.frame.destroy()

        self.frame2 = tk.Frame(self)
        self.frame2.config(background= glob.color_fondo)
        self.frame2.pack()

        # Textos

        self.resultado = "Acción exitosa" 
        self.mensaje = "La carpeta de anexos de " + glob.nombre_completo + " ha sido creada."
        self.pregunta = "¿Desea cargar documentos ahora?"
        self.consulta_label = Label(self.frame2, text=self.resultado, fg=glob.color_letras, bg=glob.color_fondo)
        self.mensaje_label = Label(self.frame2, text=self.mensaje, fg=glob.color_letras, bg=glob.color_fondo)
        self.pregunta_label = Label(self.frame2, text=self.pregunta, fg=glob.color_letras, bg=glob.color_fondo)

        # Botones

        self.boton_si = Button(self.frame2, text="Sí, vamos", command=self.openFrame, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
        self.boton_no = Button(self.frame2, text="No, lo haré después", command=self.volver, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

        # Ubicaciones

        self.consulta_label.grid(row=1, column=1, columnspan=2, pady= 10, padx=20)
        self.mensaje_label.grid(row=2, column=1, columnspan=2, pady= 10, padx=20)
        self.pregunta_label.grid(row=3, column=1, columnspan=2, pady= 5, padx=20)
        self.boton_si.grid(row=4, column=1, pady= 10, padx=20)
        self.boton_no.grid(row=4, column=2, pady= 10, padx=20)

        # Efectos

        func.Efecto_de_boton(self.boton_si)
        func.Efecto_de_boton(self.boton_no)

    #----------------------------------------------------------------------
    def volver_al_principal_desde_el_siguiente(self):
        """"""
        self.update()
        self.deiconify()
        self.volver()

########################################################################
class Demo3(tk.Toplevel):

    #----------------------------------------------------------------------
    def __init__(self, original):
        """Constructor"""
        
        self.original_frame = original
        tk.Toplevel.__init__(self)
        self.box_w = 800
        self.box_h = 500
        self.box_sw = self.winfo_screenwidth()
        self.box_sh = self.winfo_screenheight()
        self.box_x = (self.box_sw - self.box_w)/2
        self.box_y = (self.box_sh - self.box_h)/2
        self.geometry('%dx%d+%d+%d' % (self.box_w, self.box_h, self.box_x, self.box_y))
        self.titulo = "Carpeta de anexos de " + glob.nombre_completo
        self.title(self.titulo)
        self.resizable(False,False)
        self.config(background= glob.color_fondo)
        self.frame = tk.Frame(self)
        self.frame.config(background= glob.color_fondo)
        self.frame.pack()

        self.coincidencias = glob.worksheet_documentos.find(glob.nombre_completo)
        self.fila_coincidencias = self.coincidencias.row
        self.valores_de_la_fila = glob.worksheet_documentos.row_values(self.fila_coincidencias)

        self.creador_para_set = self.valores_de_la_fila[0]
        self.fecha_de_creacion_para_set = self.valores_de_la_fila[1]
        self.ID_previo_para_set = self.valores_de_la_fila[7]
        self.ID_para_set = "https://drive.google.com/drive/u/2/folders/"+self.ID_previo_para_set

        glob.ID = self.ID_previo_para_set

        # Título

        self.main_title = Label(self.frame, text="CARPETA DE ANEXOS", font=("Arial", 13), fg=glob.color_letrasenbotones, bg=glob.color_botones, width="550", height="2")
        self.main_title.pack()

        # Primer Frame ----------------------------------------

        self.primerframe = Frame(self.frame)
        self.primerframe.pack()
        self.primerframe.config(background= glob.color_fondo)

        # Textos

        self.documento_principal_label = Label(self.primerframe, text="Documento principal", fg=glob.color_letras, bg=glob.color_fondo)
        self.creador_label = Label(self.primerframe, text="Creador", fg=glob.color_letras, bg=glob.color_fondo)
        self.fecha_de_creacion_label = Label(self.primerframe, text="Fecha de creación", fg=glob.color_letras, bg=glob.color_fondo)
        self.link_de_la_carpeta_label = Label(self.primerframe, text="Link de la carpeta", fg=glob.color_letras, bg=glob.color_fondo)

        # Entry

        self.documento_principal = StringVar()
        self.documento_principal.set(glob.nombre_completo)
        self.documento_principal_entry = Entry(self.primerframe, textvariable = self.documento_principal, width= "70", state=DISABLED)

        self.creador = StringVar()
        self.creador.set(self.creador_para_set)
        self.creador_entry = Entry(self.primerframe, textvariable = self.creador, width= "70", state=DISABLED)

        self.fecha_de_creacion = StringVar()
        self.fecha_de_creacion.set(self.fecha_de_creacion_para_set)
        self.fecha_de_creacion_entry = Entry(self.primerframe, textvariable = self.fecha_de_creacion, width= "70", state=DISABLED)

        self.link_de_la_carpeta = StringVar()
        self.link_de_la_carpeta.set(self.ID_para_set)
        self.link_de_la_carpeta_entry = Entry(self.primerframe, textvariable = self.link_de_la_carpeta, width= "70", state=DISABLED)

        # Botones

        self.boton_copiar_al_portapapeles = Button(self.primerframe, text="Copiar link", command=self.copiar_al_portapapeles, width="15", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

        # Ubicaciones

        self.documento_principal_label.grid(row=1, column=1, columnspan=1)
        self.documento_principal_entry.grid(row=1, column=2, columnspan=1)
        self.creador_label.grid(row=2, column=1, columnspan=1)
        self.creador_entry.grid(row=2, column=2, columnspan=1)
        self.fecha_de_creacion_label.grid(row=3, column=1, columnspan=1)
        self.fecha_de_creacion_entry.grid(row=3, column=2, columnspan=1)
        self.link_de_la_carpeta_label.grid(row=4, column=1, columnspan=1)
        self.link_de_la_carpeta_entry.grid(row=4, column=2, columnspan=1)
        self.boton_copiar_al_portapapeles.grid(row=5, column=1, columnspan=2)

        # Efectos

        func.Efecto_de_boton(self.boton_copiar_al_portapapeles)

        # Segundo Frame ----------------------------------------

        self.segundoframe = Frame(self.frame)
        self.segundoframe.pack()
        self.segundoframe.config(background= glob.color_fondo, width=680)

        # Textos

        self.vacio1_label = Label(self.segundoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio2_label = Label(self.segundoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio3_label = Label(self.segundoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio4_label = Label(self.segundoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio5_label = Label(self.segundoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio6_label = Label(self.segundoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.contenido_de_la_carpeta_label = Label(self.segundoframe, text="Contenido de la carpeta", fg=glob.color_letras, bg=glob.color_fondo)

        # Botones

        self.boton_agregar_anexo = Button(self.segundoframe, text="Agregar anexo", command=self.agregar_anexo, width="15", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

        # Ubicaciones

        self.vacio1_label.grid(row=1, column=1, columnspan=1)
        self.vacio2_label.grid(row=1, column=2, columnspan=1)
        self.vacio3_label.grid(row=1, column=3, columnspan=1) 
        self.vacio4_label.grid(row=1, column=4, columnspan=1)
        self.vacio5_label.grid(row=1, column=5, columnspan=1)
        self.vacio6_label.grid(row=1, column=6, columnspan=1)
        self.boton_agregar_anexo.grid(row=2, column=6, columnspan=1)
        self.contenido_de_la_carpeta_label.grid(row=3, column=1, columnspan=6)

        # Efectos

        func.Efecto_de_boton(self.boton_agregar_anexo)

        # Tercer Frame ----------------------------------------

        self.tercerframe = Frame(self.frame)
        self.tercerframe.pack()
        self.tercerframe.config(background= glob.color_fondo, width=680,height=50)

        # Textos

        self.nombre_del_anexo_label = Label(self.tercerframe, text="Nombre del anexo", fg=glob.color_letras, bg=glob.color_fondo, width= "25")
        self.cargado_por_label = Label(self.tercerframe, text="Cargado por", fg=glob.color_letras, bg=glob.color_fondo, width= "25")
        self.fecha_de_carga_label = Label(self.tercerframe, text="Fecha de carga", fg=glob.color_letras, bg=glob.color_fondo, width= "25")
        self.vacio_label = Label(self.tercerframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "9")

        # Ubicaciones

        self.nombre_del_anexo_label.grid(row=1, column=1, columnspan=1)
        self.cargado_por_label.grid(row=1, column=2, columnspan=1)
        self.fecha_de_carga_label.grid(row=1, column=3, columnspan=1)
        self.vacio_label.grid(row=1, column=4, columnspan=1)

        # Cuarto Frame ----------------------------------------

        glob.cuartoframe = Frame(self.frame)
        glob.cuartoframe.pack()
        glob.cuartoframe.config(background= glob.color_fondo, width=680,height=50)

        self.listar()

        # Quinto Frame ----------------------------------------

        self.quintoframe = Frame(self.frame)
        self.quintoframe.pack()
        self.quintoframe.config(background= glob.color_fondo)

        # Texto

        self.vacio1_label = Label(self.quintoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio2_label = Label(self.quintoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio3_label = Label(self.quintoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio4_label = Label(self.quintoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio5_label = Label(self.quintoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")
        self.vacio6_label = Label(self.quintoframe, text="", fg=glob.color_letras, bg=glob.color_fondo, width= "15")

        # Botones

        self.boton_volver_a_busqueda = Button(self.quintoframe, text="Volver a búsqueda", command=self.volver_a_principal, width="15", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
        self.boton_salir = Button(self.quintoframe, text="Salir", command=self.salir, width="15", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

        # Ubicaciones

        self.vacio1_label.grid(row=1, column=1, columnspan=1)
        self.vacio2_label.grid(row=1, column=2, columnspan=1)
        self.vacio3_label.grid(row=1, column=3, columnspan=1)
        self.vacio4_label.grid(row=1, column=4, columnspan=1)
        self.vacio5_label.grid(row=1, column=5, columnspan=1)
        self.vacio6_label.grid(row=1, column=6, columnspan=1)

        self.boton_volver_a_busqueda.grid(row=2, column=5, columnspan=1, padx=1, pady=1)
        self.boton_salir.grid(row=2, column=6, columnspan=1, padx=1, pady=1)

        # Efectos

        func.Efecto_de_boton(self.boton_volver_a_busqueda)
        func.Efecto_de_boton(self.boton_salir)    

    #----------------------------------------------------------------------
    def ver(self, x, xd):
        """"""
        self.x = x
        self.xd = xd
        tkmb.showinfo(x, xd, parent=self)

    #----------------------------------------------------------------------
    def download_anexo(self, m, x):
        """"""
        self.m = m
        self.x = x
        self.elegir_ubicacion = askdirectory(parent=self.master)
        self.ubicacion_archivo = self.elegir_ubicacion + "/" + self.x
        self.anexo_descargado = glob.drive.CreateFile({'id': self.m})
        self.anexo_descargado.GetContentFile(self.ubicacion_archivo)
    
    #----------------------------------------------------------------------
    def delete_anexo(self, x, xd, m):
        """"""
        self.x = x
        self.xd = xd
        self.m = m

        self.anexo_quitado = glob.drive.CreateFile({'id': self.m})
        self.anexo_quitado.Trash()
        self.nombre_del_anexo_quitado = self.x
        self.descripcion_del_anexo_quitado = self.xd
        self.ahora_anexo_quitado = str(datetime.now())
        self.ID_anexo_quitado = self.m
        self.bye_anexo = [glob.nombre_completo, self.nombre_del_anexo_quitado, self.descripcion_del_anexo_quitado, glob.usuario, self.ahora_anexo_quitado, self.ID_anexo_quitado]
        glob.worksheet_eliminados.append_row(self.bye_anexo)
        self.celda_del_ID_a_eliminar = glob.worksheet_anexos.find(self.m)
        self.fila_del_ID_a_eliminar = self.celda_del_ID_a_eliminar.row
        self.fila_del_ID_a_eliminar = str(self.fila_del_ID_a_eliminar)
        self.rango_a_eliminar = "A" + self.fila_del_ID_a_eliminar + ":" + "F" + self.fila_del_ID_a_eliminar
        glob.worksheet_anexos.update(self.rango_a_eliminar, [["Eliminado", "Eliminado", "Eliminado", "Eliminado", "Eliminado", "Eliminado"]])
        self.lista_a_olvidar = glob.cuartoframe.pack_slaves()
        for l in self.lista_a_olvidar:
            l.destroy()
        self.listar()

    #----------------------------------------------------------------------
    def listar(self):
        """"""

        # Imágenes

        self.lupa = func.Dar_formato_a_Imagen(18, 18, 'lupa.png')
        self.lupa = ImageTk.PhotoImage(self.lupa)

        self.descargar = func.Dar_formato_a_Imagen(18, 18, 'descargar.png')
        self.descargar = ImageTk.PhotoImage(self.descargar)

        self.eliminar = func.Dar_formato_a_Imagen(18, 18, 'eliminar.png')
        self.eliminar = ImageTk.PhotoImage(self.eliminar)

        # Información para mostrar el contenido de la carpeta

        self.dataframe = pd.DataFrame(glob.worksheet_anexos.get_all_records())
        self.data_seleccionada = self.dataframe['Documento principal'] == glob.nombre_completo
        self.dataframe_filtrado = self.dataframe[self.data_seleccionada]
        self.all_entries = []
        self.all_boton = []

        for i, row in self.dataframe_filtrado.iterrows():

            glob.x = row[1]
            glob.xd = row[2]
            glob.y = row[3]
            glob.z = row[4]
            glob.m = row[5]

            self.x = glob.x
            self.xd = glob.xd
            self.y = glob.y
            self.z = glob.z
            self.m = glob.m

            self.cuartoframe_dentro = Frame(glob.cuartoframe)
            self.cuartoframe_dentro.pack()
            self.cuartoframe_dentro.config(background= glob.color_fondo)
            
            # Entry

            self.nombre_del_anexo = StringVar()
            self.nombre_del_anexo.set(glob.x)
            self.nombre_del_anexo_entry = Entry(self.cuartoframe_dentro, textvariable = self.nombre_del_anexo, width= "30", state=DISABLED)

            self.cargado_por = StringVar()
            self.cargado_por.set(glob.y)
            self.cargado_por_entry = Entry(self.cuartoframe_dentro, textvariable = self.cargado_por, width= "30", state=DISABLED)

            self.fecha_de_carga = StringVar()
            self.fecha_de_carga.set(glob.z)
            self.fecha_de_carga_entry = Entry(self.cuartoframe_dentro, textvariable = self.fecha_de_carga, width= "30", state=DISABLED)

            # Botones

            self.boton_ver = Button(self.cuartoframe_dentro, image=self.lupa, command= lambda x=glob.x, xd=glob.xd: self.ver(x, xd), compound="top", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
            self.boton_descargar = Button(self.cuartoframe_dentro, image=self.descargar, command= lambda m=glob.m, x=glob.x: self.download_anexo(m, x), compound="top", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
            self.boton_eliminar = Button(self.cuartoframe_dentro, image=self.eliminar, command= lambda m=glob.m, x=glob.x, xd=glob.xd: self.delete_anexo(x, xd, m), compound="top", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

            # Ubicaciones

            self.nombre_del_anexo_entry.grid(row=2, column=1, columnspan=1, pady=1, padx=1)
            self.cargado_por_entry.grid(row=2, column=2, columnspan=1, pady=1, padx=1)
            self.fecha_de_carga_entry.grid(row=2, column=3, columnspan=1, pady=1, padx=1)
            self.boton_ver.grid(row=2, column=4, columnspan=1, pady=1, padx=1)
            self.boton_descargar.grid(row=2, column=5, columnspan=1, pady=1, padx=1)
            self.boton_eliminar.grid(row=2, column=6, columnspan=1, pady=1, padx=1)

            self.all_entries.append( (self.nombre_del_anexo_entry, self.cargado_por_entry, self.fecha_de_carga_entry) )
            self.all_boton.append( (self.boton_ver, self.boton_descargar, self.boton_eliminar))

            # Efectos

            func.Efecto_de_boton(self.boton_ver)
            func.Efecto_de_boton(self.boton_descargar)
            func.Efecto_de_boton(self.boton_eliminar)       

    #----------------------------------------------------------------------
    def volver_a_principal(self):
        """"""
        self.destroy()
        self.original_frame.volver_al_principal_desde_el_siguiente()

    #----------------------------------------------------------------------
    def salir(self):
        """"""
        self.destroy()
    
    #----------------------------------------------------------------------
    def copiar_al_portapapeles(self):
        """"""

        self.primerframe.clipboard_clear()
        self.primerframe.clipboard_append(self.link_de_la_carpeta_entry.get())

    #----------------------------------------------------------------------
    def agregar_anexo(self):
        """"""
        self.withdraw()
        subFrame = Demo4(self)

    #----------------------------------------------------------------------
    def show(self):
        """"""
        self.update()
        self.deiconify()

########################################################################
class Demo4(tk.Toplevel):
    """"""

    #----------------------------------------------------------------------
    def __init__(self, original):
        """Constructor"""
        
        self.original_frame = original
        tk.Toplevel.__init__(self)
        self.box_w = 520
        self.box_h = 450
        self.box_sw = self.winfo_screenwidth()
        self.box_sh = self.winfo_screenheight()
        self.box_x = (self.box_sw - self.box_w)/2
        self.box_y = (self.box_sh - self.box_h)/2
        self.geometry('%dx%d+%d+%d' % (self.box_w, self.box_h, self.box_x, self.box_y))
        self.titulo = "Aplicativo: Herramientas de Sefa - versión " + glob.numero_de_version
        self.title(self.titulo)
        self.resizable(False,False)
        self.config(background= glob.color_fondo)
        
        # Título

        self.tituloframe = tk.Frame(self)
        self.tituloframe.config(background= glob.color_fondo)
        self.tituloframe.pack()

        self.main_title = Label(self.tituloframe, text="CARGA DE ANEXOS", font=glob.fuente_titulos, fg=glob.color_letrasenbotones, bg=glob.color_botones, width="550", height="2")
        self.main_title.pack()

        self.imagendecarga = func.Dar_formato_a_Imagen(100, 100, 'subir.png')
        self.imagendecarga = ImageTk.PhotoImage(self.imagendecarga)
        self.imagen_label = Label(self.tituloframe, image=self.imagendecarga, bg=glob.color_fondo)
        self.imagen_label.pack()
        
        # Cuerpo de la ventana

        self.frame = tk.Frame(self)
        self.frame.config(background= glob.color_fondo)
        self.frame.pack()

        # Textos

        self.ingresar_nombre_del_anexo_label = Label(self.frame, text="Nombre del anexo:", fg=glob.color_letras, bg=glob.color_fondo, anchor="w", width="70")
        self.ingresar_descripcion_label = Label(self.frame, text="Descripción del contenido del anexo:", fg=glob.color_letras, bg=glob.color_fondo, anchor="w", width="70")
        self.escoger_label = Label(self.frame, text="Documento seleccionado:", fg=glob.color_letras, bg=glob.color_fondo, anchor="w", width="70")

        # Entry y texto largo

        self.ingresar_nombre_del_anexo = StringVar()
        self.ingresar_nombre_del_anexo_entry = Entry(self.frame, textvariable = self.ingresar_nombre_del_anexo, width= "80", borderwidth=2, bg=glob.color_entry)

        self.ingresar_descripcion_text = Text(self.frame)
        self.ingresar_descripcion_text.config(width="60", height="5", borderwidth=2, bg=glob.color_entry)

        self.documento_elegido = StringVar()
        self.documento_elegido_entry = Entry(self.frame, textvariable = self.documento_elegido, width= "80", state=DISABLED)

        # Botones

        self.boton_escoger_documento = Button(self.frame, text="Escoger documento", command=self.buscar_documento, width="15", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
        self.boton_cargar_documento = Button(self.frame, text="Cargar documento", command=self.cargar_documento, width="15", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
        self.boton_volver = Button(self.frame, text="Volver", command=self.volver, width="15", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
        
        # Ubicaciones

        self.ingresar_nombre_del_anexo_label.grid(row=1, column=1, columnspan=2)
        self.ingresar_nombre_del_anexo_entry.grid(row=2, column=1, columnspan=2)
        self.ingresar_descripcion_label.grid(row=3, column=1, columnspan=2)
        self.ingresar_descripcion_text.grid(row=4, column=1, columnspan=2)
        self.boton_escoger_documento.grid(row=5, column=1, columnspan=2)
        self.escoger_label.grid(row=6, column=1, columnspan=2)
        self.documento_elegido_entry.grid(row=7, column=1, columnspan=2)
        self.boton_cargar_documento.grid(row=8, column=1, columnspan=1)
        self.boton_volver.grid(row=8, column=2, columnspan=1)

        # Efectos

        func.Efecto_de_boton(self.boton_escoger_documento)
        func.Efecto_de_boton(self.boton_cargar_documento)
        func.Efecto_de_boton(self.boton_volver)

    #----------------------------------------------------------------------
    def buscar_documento(self):
        """"""
        self.nuevo_anexo = askopenfilename(filetypes =[('all files', '.*')], parent=self.master)
        self.documento_elegido.set(self.nuevo_anexo)

    #----------------------------------------------------------------------
    def cargar_documento(self):
        """"""
        self.documento_elegido_data = self.documento_elegido.get()
        self.extension = os.path.splitext(self.documento_elegido_data)[1]
        self.nombre_del_anexo_data = self.ingresar_nombre_del_anexo.get() + self.extension
        self.descripcion_del_anexo_data = self.ingresar_descripcion_text.get("1.0","end-1c")
        self.ahora_anexo = str(datetime.now())

        if self.nombre_del_anexo_data == "":
            tkmb.showerror("Error", "Falta ingresar el nombre del anexo.", parent=self.master)

        elif self.descripcion_del_anexo_data == "":
            tkmb.showerror("Error", "Falta ingresar la descripción del anexo.", parent=self.master)

        elif self.documento_elegido_data == "":
            tkmb.showerror("Error", "Falta seleccionar el documento del anexo.", parent=self.master)

        else:

            self.instancia_de_drive = glob.drive.CreateFile({'title': self.nombre_del_anexo_data, 'parents':[{'id':glob.ID}]})
            self.instancia_de_drive.SetContentFile(self.documento_elegido_data)
            self.instancia_de_drive.Upload()

            self.ID_anexo = self.instancia_de_drive['id']
            self.new_anexo = [glob.nombre_completo, self.nombre_del_anexo_data, self.descripcion_del_anexo_data, glob.usuario, self.ahora_anexo, self.ID_anexo]
            glob.worksheet_anexos.append_row(self.new_anexo)
            self.volver()

            self.lista_a_olvidar = glob.cuartoframe.pack_slaves()
            for l in self.lista_a_olvidar:
                l.destroy()
            self.original_frame.listar()

    #----------------------------------------------------------------------
    def volver(self):
        """"""
        self.destroy()
        self.original_frame.show()

########################################################################
class MyApp(tk.Toplevel):
    """"""

    #----------------------------------------------------------------------
    def __init__(self, parent):
        """Constructor"""

        tk.Toplevel.__init__(self)
        self.box_w = 900
        self.box_h = 450
        self.box_sw = self.winfo_screenwidth()
        self.box_sh = self.winfo_screenheight()
        self.box_x = (self.box_sw - self.box_w)/2
        self.box_y = (self.box_sh - self.box_h)/2
        self.geometry('%dx%d+%d+%d' % (self.box_w, self.box_h, self.box_x, self.box_y))
        self.titulo = "Aplicativo: Herramientas de Sefa - versión " + glob.numero_de_version
        self.title(self.titulo)
        self.resizable(False,False)
        self.config(background= glob.color_fondo)
        
        # Título

        self.tituloframe = tk.Frame(self)
        self.tituloframe.config(background= glob.color_fondo)
        self.tituloframe.pack()
        self.main_title = Label(self.tituloframe, text="ANEXOS DE SEFA EN LA NUBE", font=glob.fuente_titulos, fg=glob.color_letrasenbotones, bg=glob.color_botones, width="550", height="2")
        self.main_title.pack()

        # Imagen Frame

        self.anexosdesefa = func.Dar_formato_a_Imagen(188, 150, 'nube.png')
        self.anexosdesefa = ImageTk.PhotoImage(self.anexosdesefa)
        
        self.imageframe = tk.Frame(self)
        self.imageframe.config(background= glob.color_fondo)
        self.imageframe.pack()
        
        self.vacio_label1 = Label(self.imageframe, text= "", bg=glob.color_fondo)
        self.imagen_label = Label(self.imageframe, image=self.anexosdesefa, bg=glob.color_fondo)
        self.vacio_label2 = Label(self.imageframe, text= "", bg=glob.color_fondo)
        self.vacio_label1.pack()
        self.imagen_label.pack()
        self.vacio_label2.pack()

        # Contenido Frame

        self.frame = tk.Frame(self)
        self.frame.config(background= glob.color_fondo)
        self.frame.pack()
        self.reg = self.frame.register(self.callback)
        
        # Label
        
        self.tipo_label = Label(self.frame, text="Tipo", fg=glob.color_letras, bg=glob.color_fondo)
        self.numero_label = Label(self.frame, text="Número", fg=glob.color_letras, bg=glob.color_fondo)
        self.ano_label = Label(self.frame, text="Año", fg=glob.color_letras, bg=glob.color_fondo)
        self.terminacion_label = Label(self.frame, text="Terminación", fg=glob.color_letras, bg=glob.color_fondo)
        
        # Entry y desplegables

        self.combostyle = ttk.Style()
        self.combostyle.theme_create('combostyle', parent='alt',
            settings = {'TCombobox':
                {'configure':
                    {'selectbackground': glob.color_botones,
                    'fieldbackground': glob.color_entry,
                    'background': glob.color_entry
                    }}})
        self.combostyle.theme_use('combostyle') 

        self.tipo_desplegable = ttk.Combobox(self.frame, width= "25", state="readonly")
        self.tipo_opciones = ["Carta", "Informe", "Memorando", "Oficio"]
        self.tipo_desplegable['values'] = self.tipo_opciones

        self.numero = StringVar()
        self.numero_entry = Entry(self.frame, textvariable = self.numero, width= "25", borderwidth=2, bg=glob.color_entry)
        self.numero_entry.config(validate="key", validatecommand=(self.reg, '%P'))

        self.ano_desplegable = ttk.Combobox(self.frame, width= "25", state="readonly")
        self.ano_opciones = ["2021"]
        self.ano_desplegable['values'] = self.ano_opciones

        self.terminacion_desplegable = ttk.Combobox(self.frame, width= "25", state="readonly")
        self.terminacion_opciones = ["OEFA/DPEF-SEFA", "OEFA/DPEF-SEFA-SINADA", "OEFA/DPEF-SEFA-COFEMA"]
        self.terminacion_desplegable['values'] = self.terminacion_opciones

        # Botones

        self.boton_limpiar = Button(self.frame, text="Limpiar", command=self.limpiar_valores_de_busqueda, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")
        self.boton_buscar = Button(self.frame, text="Buscar", command=self.buscar_coincidencias, width="30", height="1", fg=glob.color_letrasenbotones, bg=glob.color_botones, relief="flat", cursor="hand2")

        # Ubicaciones

        self.tipo_label.grid(row=1, column=1, pady= 2, padx=20)
        self.tipo_desplegable.grid(row=2, column=1, pady= 2, padx=10)
        self.numero_label.grid(row=1, column=2, pady= 2, padx=10)
        self.numero_entry.grid(row=2, column=2, pady= 2, padx=10)
        self.ano_label.grid(row=1, column=3, pady= 2, padx=10)
        self.ano_desplegable.grid(row=2, column=3, pady= 2, padx=10)
        self.terminacion_label.grid(row=1, column=4, pady= 2, padx=10)
        self.terminacion_desplegable.grid(row=2, column=4, pady= 2, padx=10)
        self.boton_limpiar.grid(row=4, column=1, columnspan=2, pady= 20)
        self.boton_buscar.grid(row=4, column=3, columnspan=2, pady= 20)

        # Efectos

        func.Efecto_de_boton(self.boton_limpiar)
        func.Efecto_de_boton(self.boton_buscar)

    #----------------------------------------------------------------------
    def limpiar_valores_de_busqueda(self):
        """"""
        self.tipo_desplegable.set('')
        self.numero_entry.delete(0, END)
        self.ano_desplegable.set('')
        self.terminacion_desplegable.set('')

    #----------------------------------------------------------------------
    def callback(self, input):
        """"""
        if input.isdigit():
            return True
        elif input == "":
            return True
        else:
            return False

    #----------------------------------------------------------------------
    def buscar_coincidencias(self):
        """"""
        glob.tipo_data = self.tipo_desplegable.get()
        glob.numero_data = self.numero.get()
        glob.ano_data = self.ano_desplegable.get()
        glob.terminacion_data = self.terminacion_desplegable.get()

        if glob.tipo_data == "":
            tkmb.showerror("Error", "Falta seleccionar el tipo de documento.", parent=self.root)
            
        elif glob.numero_data == "":
            tkmb.showerror("Error", "Falta ingresar el número del documento.", parent=self.root)

        elif glob.ano_data == "":
            tkmb.showerror("Error", "Falta seleccionar el año del documento.", parent=self.root)

        elif glob.terminacion_data == "":
            tkmb.showerror("Error", "Falta seleccionar la terminación del documento.", parent=self.root)
        
        else:
            glob.numero_data = int(glob.numero_data, base=10)
            glob.numero_data = str(glob.numero_data)
            glob.nombre_completo = glob.tipo_data + " " + glob.numero_data + "-" + glob.ano_data + "-" + glob.terminacion_data
            func.GoogleSheet_Anexos()
            func.Drive()
            self.cell_list = glob.worksheet_documentos.findall(glob.nombre_completo)
            if self.cell_list != []:
                glob.comprobar_coincidencia = "sí"
            else:
                glob.comprobar_coincidencia = "no"
            self.withdraw()
            subFrame = Demo2(self)

    #----------------------------------------------------------------------
    def show(self):
        """"""
        self.update()
        self.deiconify()
