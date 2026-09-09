# -*- coding: utf-8 -*-
"""
Ventana para convertir el Excel de invitados al archivo .glf de Guest List.

Pensada para usarse sin saber nada de programacion: se elige el Excel y la lista
aparece sola. Lo que hay que revisar sale como tarjetas cortas que se despliegan,
no como un muro de texto. Se empaqueta como .exe con construir_exe.py.

Las piezas redondeadas estan en interfaz.py y los colores en tema.py.
"""

import os
import sys
import queue
import webbrowser
import threading
import traceback
import subprocess

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl

import glf
import icono
import actualizacion
import excel_a_glf as conv
import tema
from tema import C
from interfaz import Tarjeta, Boton, Selector, Campo


TITULO = 'Lista de invitados  →  Guest List'
SIN_HOJA = conv.SIN_HOJA

F_TITULO   = ('Segoe UI Semibold', 17)
F_SUB      = ('Segoe UI', 10)
F_ETIQUETA = ('Segoe UI Semibold', 10)
F_TEXTO    = ('Segoe UI', 10)
F_CHICO    = ('Segoe UI', 9)
F_NUMERO   = ('Segoe UI Semibold', 26)

# Lo que se sepa de la version nueva se guarda aqui: None = todavia no se pregunto,
# False = se pregunto y no hay, dict = hay una. Asi rearmar la ventana al cambiar de
# tema no dispara otra consulta a internet.
_VERSION_NUEVA = None


def abrir_carpeta(ruta):
    """Abre el explorador con el archivo seleccionado."""
    try:
        if sys.platform == 'win32':
            subprocess.Popen(['explorer', '/select,', os.path.normpath(ruta)])
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', '-R', ruta])
        else:
            subprocess.Popen(['xdg-open', os.path.dirname(ruta)])
    except OSError:
        pass


class Hallazgo(Tarjeta):
    """Una cosa que revisar. Titulo siempre visible; el detalle se despliega."""

    def __init__(self, padre, dato):
        estilo = C['niveles'][dato['nivel']]
        super().__init__(padre, radio=12, relleno=estilo['fondo'],
                         borde=estilo['borde'], fondo=C['tarjeta'], inset=6)
        self.detalle = dato['detalle']
        self.abierto = False
        cuerpo = self.cuerpo
        cuerpo.columnconfigure(1, weight=1)

        marca = tk.Label(cuerpo, text=estilo['icono'], bg=estilo['fondo'],
                         fg=estilo['texto'], font=('Segoe UI Semibold', 11), width=2)
        marca.grid(row=0, column=0, sticky='nw', padx=(8, 0), pady=8)

        self.titulo = tk.Label(cuerpo, text=dato['titulo'], bg=estilo['fondo'],
                               fg=estilo['texto'], font=F_ETIQUETA, justify='left',
                               anchor='w', wraplength=620)
        self.titulo.grid(row=0, column=1, sticky='ew', pady=8)

        self.flecha = None
        if self.detalle:
            self.flecha = tk.Label(cuerpo, text='ver  ▾', bg=estilo['fondo'],
                                   fg=estilo['texto'], font=F_CHICO, cursor='hand2')
            self.flecha.grid(row=0, column=2, sticky='ne', padx=10, pady=10)

            self.panel = tk.Frame(cuerpo, bg=estilo['fondo'])
            tk.Label(self.panel, text='\n'.join(self.detalle), bg=estilo['fondo'],
                     fg=estilo['texto'], font=F_CHICO, justify='left', anchor='w',
                     wraplength=600).pack(anchor='w')
            for w in (cuerpo, marca, self.titulo, self.flecha):
                w.bind('<Button-1>', self._alternar)
                w.configure(cursor='hand2')

    def _alternar(self, _=None):
        self.abierto = not self.abierto
        if self.abierto:
            self.panel.grid(row=1, column=1, columnspan=2, sticky='ew',
                            padx=(0, 10), pady=(0, 10))
            self.flecha.configure(text='ocultar  ▴')
        else:
            self.panel.grid_forget()
            self.flecha.configure(text='ver  ▾')


class VentanaTexto(tk.Toplevel):
    """Para las listas que llegan por WhatsApp: se pegan y ya.

    Devuelve el texto en `resultado`, o None si se cancela.
    """

    EJEMPLO = ('Pega aqui la lista tal como te la mandaron.\n\n'
               'Un invitado por renglon. Sirve con vinetas, numeracion o sellos de\n'
               'WhatsApp; todo eso se limpia solo.\n\n'
               'Si trae mesas, deja los renglones "Mesa 1", "Mesa 2"... con sus\n'
               'invitados debajo. Y si un renglon son varias personas, se puede\n'
               'escribir el numero al final: "Familia Salas - 5".')

    def __init__(self, padre, texto_previo=''):
        super().__init__(padre, bg=C['fondo'])
        self.resultado = None
        self.title('Pegar la lista de invitados')
        self.configure(padx=18, pady=18)
        self.transient(padre)

        tk.Label(self, text='Pega la lista de invitados', bg=C['fondo'], fg=C['texto'],
                 font=('Segoe UI Semibold', 13)).pack(anchor='w')
        self.lbl_ayuda = tk.Label(self, bg=C['fondo'], fg=C['suave'], font=F_CHICO,
                                  justify='left', anchor='w',
                                  text='Un invitado por renglon. Las vinetas, la numeracion '
                                       'y los sellos de WhatsApp se limpian solos.')
        self.lbl_ayuda.pack(anchor='w', pady=(3, 10))

        caja = Tarjeta(self, radio=12, relleno=C['campo'])
        caja.pack(fill='both', expand=True)
        self.texto = tk.Text(caja.cuerpo, wrap='word', height=18, width=68,
                             bg=C['campo'], fg=C['texto'], insertbackground=C['texto'],
                             relief='flat', bd=0, highlightthickness=0,
                             font=('Segoe UI', 10), padx=10, pady=8)
        self.texto.pack(fill='both', expand=True)

        if texto_previo:
            self.texto.insert('1.0', texto_previo)
        else:
            self._poner_ejemplo()
            self.texto.bind('<FocusIn>', self._quitar_ejemplo)
            self.texto.bind('<Button-1>', self._quitar_ejemplo)

        pie = tk.Frame(self, bg=C['fondo'])
        pie.pack(fill='x', pady=(12, 0))
        Boton(pie, 'Usar esta lista', self._aceptar).pack(side='left')
        Boton(pie, 'Cancelar', self.destroy, tipo='suave').pack(side='left', padx=(10, 0))
        self.lbl_cuenta = tk.Label(pie, text='', bg=C['fondo'], fg=C['suave'], font=F_CHICO)
        self.lbl_cuenta.pack(side='left', padx=14)
        self.texto.bind('<KeyRelease>', self._contar)

        self.bind('<Escape>', lambda e: self.destroy())
        self.texto.focus_set()
        self.grab_set()

    # El texto gris de ejemplo se borra al primer clic, como cualquier campo.
    def _poner_ejemplo(self):
        self.ejemplo_puesto = True
        self.texto.insert('1.0', self.EJEMPLO)
        self.texto.configure(fg=C['suave'])

    def _quitar_ejemplo(self, _=None):
        if getattr(self, 'ejemplo_puesto', False):
            self.ejemplo_puesto = False
            self.texto.delete('1.0', 'end')
            self.texto.configure(fg=C['texto'])

    def _contenido(self):
        if getattr(self, 'ejemplo_puesto', False):
            return ''
        return self.texto.get('1.0', 'end').strip()

    def _contar(self, _=None):
        n = len([l for l in self._contenido().splitlines() if l.strip()])
        self.lbl_cuenta.configure(text='%d renglones' % n if n else '')

    def _aceptar(self):
        contenido = self._contenido()
        if not contenido:
            messagebox.showwarning('Falta la lista',
                                   'Pega la lista de invitados antes de continuar.',
                                   parent=self)
            return
        self.resultado = contenido
        self.destroy()


class Aplicacion(tk.Frame):

    def __init__(self, raiz):
        super().__init__(raiz, bg=C['fondo'])
        self.raiz = raiz
        self.pack(fill='both', expand=True)

        self.texto_pegado = None
        self.excel = None
        self.invitados = None
        self.informe = None
        self.cola = queue.Queue()
        self.cola_version = queue.Queue()
        self.ocupado = False
        self.hojas_visibles = False

        self._estilos()
        self._construir()
        self.buscar_actualizacion()

    # ------------------------------------------------------------- apariencia

    def _estilos(self):
        """Solo quedan widgets de ttk donde no hay alternativa: combos y tabla."""
        s = ttk.Style()
        try:
            s.theme_use('clam')          # el unico tema que deja fijar colores
        except tk.TclError:
            pass
        s.configure('TCombobox', fieldbackground=C['campo'], background=C['campo'],
                    foreground=C['texto'], arrowcolor=C['suave'], bordercolor=C['borde'],
                    lightcolor=C['campo'], darkcolor=C['campo'])
        s.map('TCombobox', fieldbackground=[('readonly', C['campo'])],
              foreground=[('readonly', C['texto'])],
              selectbackground=[('readonly', C['campo'])],
              selectforeground=[('readonly', C['texto'])])
        # la lista desplegable del combo no es ttk, va por opciones de Tk
        self.raiz.option_add('*TCombobox*Listbox.background', C['campo'])
        self.raiz.option_add('*TCombobox*Listbox.foreground', C['texto'])
        self.raiz.option_add('*TCombobox*Listbox.selectBackground', C['azul'])
        self.raiz.option_add('*TCombobox*Listbox.selectForeground', 'white')

        s.configure('Treeview', background=C['tarjeta'], fieldbackground=C['tarjeta'],
                    foreground=C['texto'], rowheight=28, font=F_CHICO, borderwidth=0)
        s.configure('Treeview.Heading', background=C['gris'], foreground=C['suave'],
                    font=('Segoe UI Semibold', 9), relief='flat', padding=(6, 7))
        s.map('Treeview', background=[('selected', C['seleccion'])],
              foreground=[('selected', C['texto'])])
        s.map('Treeview.Heading', background=[('active', C['gris'])])

        s.configure('Vertical.TScrollbar', background=C['borde'], troughcolor=C['tarjeta'],
                    bordercolor=C['tarjeta'], arrowcolor=C['suave'],
                    lightcolor=C['borde'], darkcolor=C['borde'], borderwidth=0)
        s.map('Vertical.TScrollbar', background=[('active', C['suave'])],
              arrowcolor=[('active', C['texto'])])

    def _construir(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(3, weight=1)
        R = dict(padx=22)

        # ---- encabezado
        cab = tk.Frame(self, bg=C['fondo'])
        cab.grid(row=0, column=0, sticky='ew', pady=(20, 12), **R)
        cab.columnconfigure(0, weight=1)
        titulos = tk.Frame(cab, bg=C['fondo'])
        titulos.grid(row=0, column=0, sticky='w')
        tk.Label(titulos, text=TITULO, bg=C['fondo'], fg=C['texto'],
                 font=F_TITULO).pack(anchor='w')
        tk.Label(titulos, bg=C['fondo'], fg=C['suave'], font=F_SUB, anchor='w',
                 text='Elige el Excel del evento y la lista se arma sola.').pack(
                     anchor='w', pady=(3, 0))
        derecha = tk.Frame(cab, bg=C['fondo'])
        derecha.grid(row=0, column=1, sticky='ne')
        self.lbl_tema = tk.Label(
            derecha, bg=C['fondo'], fg=C['suave'], font=F_CHICO, cursor='hand2',
            text='☀  Modo claro' if C['nombre'] == 'oscuro' else '☾  Modo oscuro')
        self.lbl_tema.pack(anchor='e')
        tk.Label(derecha, text='version %s' % actualizacion.VERSION, bg=C['fondo'],
                 fg=C['suave'], font=('Segoe UI', 8)).pack(anchor='e', pady=(2, 0))
        # Aviso de version nueva: se arma solo si la hay, debajo del titulo.
        self.cab = cab

        # ---- archivo
        caja = Tarjeta(self)
        caja.grid(row=1, column=0, sticky='ew', **R)
        fila1 = caja.cuerpo
        fila1.columnconfigure(1, weight=1)

        botones = tk.Frame(fila1, bg=C['tarjeta'])
        botones.grid(row=0, column=0, sticky='w', padx=(8, 0), pady=8)
        Boton(botones, 'Elegir Excel…', self.elegir_excel, tipo='suave',
              fondo=C['tarjeta']).pack(side='left')
        Boton(botones, 'Pegar texto…', self.pegar_texto, tipo='suave',
              fondo=C['tarjeta']).pack(side='left', padx=(8, 0))
        info = tk.Frame(fila1, bg=C['tarjeta'])
        info.grid(row=0, column=1, sticky='ew', padx=14)
        self.lbl_archivo = tk.Label(info, text='Ningun archivo elegido', bg=C['tarjeta'],
                                    fg=C['suave'], font=F_ETIQUETA, anchor='w')
        self.lbl_archivo.pack(anchor='w')
        self.lbl_fuente = tk.Label(info, text='', bg=C['tarjeta'], fg=C['suave'],
                                   font=F_CHICO, anchor='w', justify='left',
                                   wraplength=540)
        self.lbl_fuente.pack(anchor='w')

        self.lbl_cambiar = tk.Label(fila1, text='Cambiar hojas', bg=C['tarjeta'],
                                    fg=C['azul'], font=F_CHICO, cursor='hand2')
        self.lbl_cambiar.grid(row=0, column=2, sticky='e', padx=10)
        self.lbl_cambiar.bind('<Button-1>', self._alternar_hojas)

        # hojas: escondidas, porque casi siempre las adivina bien
        self.marco_hojas = tk.Frame(fila1, bg=C['tarjeta'])
        self.marco_hojas.columnconfigure(1, weight=1)
        self.marco_hojas.columnconfigure(3, weight=1)
        tk.Label(self.marco_hojas, text='Plano de mesas', bg=C['tarjeta'], fg=C['suave'],
                 font=F_CHICO).grid(row=0, column=0, sticky='w')
        self.cbo_mesas = ttk.Combobox(self.marco_hojas, state='disabled', width=30)
        self.cbo_mesas.grid(row=0, column=1, sticky='w', padx=(8, 24))
        tk.Label(self.marco_hojas, text='Hoja de contactos', bg=C['tarjeta'],
                 fg=C['suave'], font=F_CHICO).grid(row=0, column=2, sticky='w')
        self.cbo_lista = ttk.Combobox(self.marco_hojas, state='disabled', width=30)
        self.cbo_lista.grid(row=0, column=3, sticky='w', padx=8)
        for c in (self.cbo_mesas, self.cbo_lista):
            c.bind('<<ComboboxSelected>>', lambda e: self.convertir())

        # ---- datos del evento
        caja2 = Tarjeta(self)
        caja2.grid(row=2, column=0, sticky='ew', pady=(12, 0), **R)
        d = caja2.cuerpo
        d.columnconfigure(1, weight=1)

        tk.Label(d, text='Nombre', bg=C['tarjeta'], fg=C['texto'], font=F_ETIQUETA).grid(
            row=0, column=0, sticky='w', padx=(8, 10), pady=(8, 0))
        self.var_evento = tk.StringVar()
        Campo(d, self.var_evento, fondo=C['tarjeta']).grid(
            row=0, column=1, sticky='ew', padx=(0, 8), pady=(6, 0))
        tk.Label(d, bg=C['tarjeta'], fg=C['suave'], font=F_CHICO, anchor='w',
                 text='Asi se va a llamar el evento dentro de Guest List.').grid(
                     row=1, column=1, sticky='w', padx=6)

        fecha = tk.Frame(d, bg=C['tarjeta'])
        fecha.grid(row=2, column=0, columnspan=2, sticky='w', padx=8, pady=(8, 8))
        tk.Label(fecha, text='Fecha', bg=C['tarjeta'], fg=C['texto'],
                 font=F_ETIQUETA).pack(side='left')
        self.var_fecha = tk.StringVar()
        Campo(fecha, self.var_fecha, ancho=11, tamano=10, fondo=C['tarjeta']).pack(
            side='left', padx=(10, 0))
        tk.Label(fecha, text='a las', bg=C['tarjeta'], fg=C['suave'],
                 font=F_TEXTO).pack(side='left', padx=8)
        self.var_hora = tk.StringVar(value='18:00')
        Campo(fecha, self.var_hora, ancho=6, tamano=10, fondo=C['tarjeta']).pack(side='left')
        tk.Label(fecha, text='Opcional, como 2026-08-29. Vacia queda la fecha de hoy.',
                 bg=C['tarjeta'], fg=C['suave'], font=F_CHICO).pack(side='left', padx=12)

        # ---- resultado
        zona = tk.Frame(self, bg=C['fondo'])
        zona.grid(row=3, column=0, sticky='nsew', pady=(14, 0), **R)
        zona.columnconfigure(0, weight=1)
        zona.rowconfigure(2, weight=1)

        cifras = tk.Frame(zona, bg=C['fondo'])
        cifras.grid(row=0, column=0, sticky='ew')
        self.cifras = {}
        for i, (clave, rotulo) in enumerate((('entradas', 'invitados'),
                                             ('personas', 'personas'),
                                             ('n_mesas', 'mesas'))):
            t = Tarjeta(cifras, radio=14)
            t.grid(row=0, column=i, sticky='ew', padx=(0 if i == 0 else 10, 0))
            cifras.columnconfigure(i, weight=1)
            num = tk.Label(t.cuerpo, text='—', bg=C['tarjeta'], fg=C['texto'], font=F_NUMERO)
            num.pack(anchor='w', padx=12, pady=(8, 0))
            tk.Label(t.cuerpo, text=rotulo, bg=C['tarjeta'], fg=C['suave'],
                     font=F_CHICO).pack(anchor='w', padx=12, pady=(0, 8))
            self.cifras[clave] = num

        self.selector = Selector(zona, ['  Que revisar  ', '  Los invitados  '],
                                 self._cambiar_vista)
        self.selector.grid(row=1, column=0, sticky='w', pady=(14, 8))

        panel = Tarjeta(zona, radio=14)
        panel.grid(row=2, column=0, sticky='nsew')
        panel.cuerpo.columnconfigure(0, weight=1)
        panel.cuerpo.rowconfigure(0, weight=1)

        # vista: que revisar
        self.vista_revisar = tk.Frame(panel.cuerpo, bg=C['tarjeta'])
        self.vista_revisar.grid(row=0, column=0, sticky='nsew')
        self.vista_revisar.columnconfigure(0, weight=1)
        self.vista_revisar.rowconfigure(0, weight=1)
        self.lienzo = tk.Canvas(self.vista_revisar, bg=C['tarjeta'], highlightthickness=0)
        self.lienzo.grid(row=0, column=0, sticky='nsew')
        barra = ttk.Scrollbar(self.vista_revisar, orient='vertical', command=self.lienzo.yview)
        barra.grid(row=0, column=1, sticky='ns')
        self.lienzo.configure(yscrollcommand=barra.set)
        self.lista_hallazgos = tk.Frame(self.lienzo, bg=C['tarjeta'])
        self.ventana_lienzo = self.lienzo.create_window((0, 0), window=self.lista_hallazgos,
                                                        anchor='nw')
        self.lista_hallazgos.bind(
            '<Configure>',
            lambda e: self.lienzo.configure(scrollregion=self.lienzo.bbox('all')))
        self.lienzo.bind(
            '<Configure>',
            lambda e: self.lienzo.itemconfigure(self.ventana_lienzo, width=e.width))
        self.lienzo.bind_all('<MouseWheel>', self._rueda)

        # vista: los invitados
        self.vista_tabla = tk.Frame(panel.cuerpo, bg=C['tarjeta'])
        self.vista_tabla.grid(row=0, column=0, sticky='nsew')
        self.vista_tabla.columnconfigure(0, weight=1)
        self.vista_tabla.rowconfigure(0, weight=1)
        cols = ('nombre', 'mesa', 'personas', 'telefono', 'notas')
        self.tabla = ttk.Treeview(self.vista_tabla, columns=cols, show='headings')
        for c, txt, ancho, centro in (('nombre', 'Invitado', 340, False),
                                      ('mesa', 'Mesa', 60, True),
                                      ('personas', 'Personas', 75, True),
                                      ('telefono', 'Telefono', 150, False),
                                      ('notas', 'Notas', 190, False)):
            lado = 'center' if centro else 'w'
            self.tabla.heading(c, text=txt, anchor=lado)
            self.tabla.column(c, width=ancho, anchor=lado)
        self.tabla.tag_configure('par', background=C['gris'])
        self.tabla.grid(row=0, column=0, sticky='nsew', padx=2, pady=2)
        barra2 = ttk.Scrollbar(self.vista_tabla, orient='vertical', command=self.tabla.yview)
        barra2.grid(row=0, column=1, sticky='ns')
        self.tabla.configure(yscrollcommand=barra2.set)
        self.vista_tabla.grid_remove()

        # ---- opciones
        self.var_sin_mesa = tk.BooleanVar(value=False)
        self.var_todos = tk.BooleanVar(value=False)
        self.var_y = tk.BooleanVar(value=False)
        opciones = tk.Frame(self, bg=C['fondo'])
        opciones.grid(row=4, column=0, sticky='w', pady=(12, 0), **R)
        for var, texto in ((self.var_sin_mesa, 'Agregar los que no estan en el plano'),
                           (self.var_todos, 'No filtrar por color'),
                           (self.var_y, 'Escribir "y" en vez de "&"')):
            # tk y no ttk: el cuadrito de ttk se queda blanco en modo oscuro
            tk.Checkbutton(opciones, text=texto, variable=var, command=self.convertir,
                           bg=C['fondo'], fg=C['suave'], font=F_CHICO,
                           activebackground=C['fondo'], activeforeground=C['texto'],
                           selectcolor=C['campo'], highlightthickness=0, bd=0,
                           cursor='hand2').pack(side='left', padx=(0, 22))

        # ---- pie
        pie = tk.Frame(self, bg=C['fondo'])
        pie.grid(row=5, column=0, sticky='ew', pady=(12, 18), **R)
        self.btn_guardar = Boton(pie, 'Guardar el archivo…', self.guardar)
        self.btn_guardar.pack(side='left')
        self.btn_guardar.configure(state='disabled')
        self.lbl_estado = tk.Label(pie, text='', bg=C['fondo'], fg=C['suave'], font=F_CHICO)
        self.lbl_estado.pack(side='left', padx=14)

        self._vaciar('Elige un Excel para empezar.')

    def _rueda(self, evento):
        try:
            self.lienzo.yview_scroll(int(-evento.delta / 120), 'units')
        except tk.TclError:
            pass

    def _cambiar_vista(self, i):
        if i == 0:
            self.vista_tabla.grid_remove()
            self.vista_revisar.grid()
        else:
            self.vista_revisar.grid_remove()
            self.vista_tabla.grid()

    # ------------------------------------------------------------ version nueva

    def buscar_actualizacion(self):
        """Pregunta a GitHub si hay una version mas nueva, sin estorbar.

        Se hace en otro hilo y el resultado se guarda en el modulo, para no volver
        a preguntar cada vez que se rearma la ventana al cambiar de tema. Si no hay
        internet no pasa nada: la respuesta es None y no se muestra nada.
        """
        global _VERSION_NUEVA
        if _VERSION_NUEVA is False:          # ya se pregunto y no habia
            return
        if _VERSION_NUEVA:                   # ya se sabe cual es: solo repintarla
            self._pintar_aviso_version(_VERSION_NUEVA)
            return
        threading.Thread(target=self._preguntar_version, daemon=True).start()
        self.after(1200, self._revisar_version)

    def _preguntar_version(self):
        self.cola_version.put(actualizacion.buscar() or False)

    def _revisar_version(self):
        global _VERSION_NUEVA
        try:
            nueva = self.cola_version.get_nowait()
        except queue.Empty:
            try:
                self.after(900, self._revisar_version)
            except tk.TclError:
                pass                          # la ventana se rearmo; no pasa nada
            return
        _VERSION_NUEVA = nueva
        if nueva:
            self._pintar_aviso_version(nueva)

    def _pintar_aviso_version(self, nueva):
        try:
            aviso = Tarjeta(self.cab, radio=12, relleno=C['gris'], borde=C['azul'],
                            fondo=C['fondo'], inset=6)
            aviso.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(12, 0))
            cuerpo = aviso.cuerpo
            cuerpo.columnconfigure(0, weight=1)
            tk.Label(cuerpo, bg=C['gris'], fg=C['texto'], font=F_ETIQUETA, anchor='w',
                     text='Hay una version nueva: %s' % nueva['version']).grid(
                         row=0, column=0, sticky='w', padx=10, pady=8)
            tk.Label(cuerpo, bg=C['gris'], fg=C['suave'], font=F_CHICO, anchor='w',
                     text='Tienes la %s.' % actualizacion.VERSION).grid(
                         row=0, column=1, sticky='e', padx=(0, 10))
            Boton(cuerpo, 'Descargar', lambda: webbrowser.open(nueva['url']),
                  fondo=C['gris']).grid(row=0, column=2, sticky='e', padx=(0, 6), pady=5)
        except tk.TclError:
            pass

    # -------------------------------------------------------------------- tema

    def _estado(self):
        """Lo que hay escrito en la ventana, para no perderlo al cambiar de tema."""
        return {'excel': self.excel, 'texto': self.texto_pegado,
                'evento': self.var_evento.get(),
                'fecha': self.var_fecha.get(), 'hora': self.var_hora.get(),
                'mesas': self.cbo_mesas.get(), 'lista': self.cbo_lista.get(),
                'hojas': list(self.cbo_mesas['values']),
                'sin_mesa': self.var_sin_mesa.get(), 'todos': self.var_todos.get(),
                'y': self.var_y.get(), 'hojas_visibles': self.hojas_visibles}

    def _restaurar(self, e):
        self.var_evento.set(e['evento'])
        self.var_fecha.set(e['fecha'])
        self.var_hora.set(e['hora'])
        self.var_sin_mesa.set(e['sin_mesa'])
        self.var_todos.set(e['todos'])
        self.var_y.set(e['y'])
        if e['hojas_visibles']:
            self._alternar_hojas()
        if e.get('texto'):
            self.cargar_texto(e['texto'])
            return
        if not e['excel']:
            return
        self.excel = e['excel']
        self.lbl_archivo.configure(text=os.path.basename(e['excel']), fg=C['texto'])
        for combo, valor in ((self.cbo_mesas, e['mesas']), (self.cbo_lista, e['lista'])):
            combo.configure(state='readonly', values=e['hojas'])
            combo.set(valor)
        self.convertir()

    def cambiar_tema(self):
        """Cambia de paleta rearmando la ventana: es lo unico fiable en Tk, porque
        cada widget se pinta con los colores que habia cuando se creo."""
        estado = self._estado()
        nuevo = 'claro' if C['nombre'] == 'oscuro' else 'oscuro'
        tema.aplicar(nuevo)
        tema.guardar_preferencias(dict(tema.leer_preferencias(), tema=nuevo))
        raiz = self.raiz
        raiz.configure(bg=C['fondo'])
        tema.barra_de_titulo_oscura(raiz, nuevo == 'oscuro')
        self.destroy()
        Aplicacion(raiz)._restaurar(estado)

    # ---------------------------------------------------------------- acciones

    def elegir_excel(self):
        ruta = filedialog.askopenfilename(
            title='Elegir el Excel del evento',
            filetypes=[('Archivos de Excel', '*.xlsx *.xlsm'), ('Todos', '*.*')])
        if ruta:
            self.cargar_excel(ruta)

    def cargar_excel(self, ruta):
        """Deja el Excel listo, adivina las hojas y convierte de una vez."""
        try:
            libro = openpyxl.load_workbook(ruta, data_only=True)
            hojas = list(libro.sheetnames)
            lista, mesas = conv.hojas_probables(libro)
        except Exception as e:
            messagebox.showerror('No pude leer el Excel', str(e))
            return

        self.excel = ruta
        self.texto_pegado = None
        self.lbl_archivo.configure(text=os.path.basename(ruta), fg=C['texto'])
        self.lbl_cambiar.grid()
        for combo, valor in ((self.cbo_mesas, mesas), (self.cbo_lista, lista)):
            combo.configure(state='readonly', values=[SIN_HOJA] + hojas)
            combo.set(valor or SIN_HOJA)
        if not self.var_evento.get().strip():
            self.var_evento.set(os.path.splitext(os.path.basename(ruta))[0].strip().upper())
        self.convertir()

    def pegar_texto(self):
        """Para las listas que llegan por WhatsApp en vez de en un Excel."""
        ventana = VentanaTexto(self.raiz, self.texto_pegado or '')
        self.raiz.wait_window(ventana)
        if ventana.resultado:
            self.cargar_texto(ventana.resultado)

    def cargar_texto(self, texto):
        self.texto_pegado = texto
        self.excel = None
        n = len([l for l in texto.splitlines() if l.strip()])
        self.lbl_archivo.configure(text='Lista pegada (%d renglones)' % n, fg=C['texto'])
        if self.hojas_visibles:          # las hojas no aplican a una lista de texto
            self._alternar_hojas()
        self.lbl_cambiar.grid_remove()
        self.convertir()

    def convertir(self):
        if self.ocupado or not (self.excel or self.texto_pegado):
            return
        self.ocupado = True
        self.btn_guardar.configure(state='disabled')
        self.lbl_estado.configure(text='')

        if self.texto_pegado:
            self._vaciar('Leyendo la lista…')
            args = dict(agregar_sin_mesa=self.var_sin_mesa.get(),
                        amper_y=self.var_y.get())
            tarea = ('texto', self.texto_pegado, args)
        else:
            self._vaciar('Leyendo el Excel…')
            args = dict(hoja=self.cbo_lista.get() or SIN_HOJA,
                        hoja_mesas=self.cbo_mesas.get() or SIN_HOJA,
                        agregar_sin_mesa=self.var_sin_mesa.get(),
                        todos=self.var_todos.get(),
                        amper_y=self.var_y.get())
            tarea = ('excel', self.excel, args)

        threading.Thread(target=self._trabajar, args=tarea, daemon=True).start()
        self.after(80, self._revisar_cola)

    def _trabajar(self, clase, fuente, args):
        # BaseException, no Exception: si algo lanza SystemExit el hilo moriria en
        # silencio y la ventana se quedaria esperando para siempre.
        try:
            hacer = conv.convertir_texto if clase == 'texto' else conv.convertir
            self.cola.put(('ok', hacer(fuente, **args)))
        except BaseException as e:
            self.cola.put(('error', (e, traceback.format_exc())))

    def _revisar_cola(self):
        try:
            tipo, carga = self.cola.get_nowait()
        except queue.Empty:
            self.after(80, self._revisar_cola)
            return

        self.ocupado = False
        if tipo == 'error':
            e, rastro = carga
            entendible = isinstance(e, conv.ExcelNoEntendido)
            self.invitados = None
            self._vaciar(str(e) if entendible else
                         'Ocurrio un error inesperado.\n\n%s' % rastro)
            self.lbl_fuente.configure(text='No pude leer esta lista.')
            for c in self.cifras.values():
                c.configure(text='—')
            if entendible:
                messagebox.showerror('No pude leer esta lista', str(e))
            return

        self.invitados, self.informe = carga
        self._mostrar()

    # ----------------------------------------------------------------- pintar

    def _vaciar(self, mensaje):
        for w in self.lista_hallazgos.winfo_children():
            w.destroy()
        tk.Label(self.lista_hallazgos, text=mensaje, bg=C['tarjeta'], fg=C['suave'],
                 font=F_TEXTO, justify='left', anchor='w', wraplength=620).pack(
                     anchor='w', padx=16, pady=16)

    def _mostrar(self):
        inf, invitados = self.informe, self.invitados
        self.lbl_fuente.configure(text=conv.resumen_fuente(inf))

        for clave, etiqueta in self.cifras.items():
            valor = inf.get(clave)
            etiqueta.configure(text='—' if valor is None else str(valor))

        for w in self.lista_hallazgos.winfo_children():
            w.destroy()
        encontrados = conv.hallazgos(inf)
        if not encontrados:
            tk.Label(self.lista_hallazgos, text='Todo en orden, no hay nada que revisar.',
                     bg=C['tarjeta'], fg=C['suave'], font=F_TEXTO).pack(
                         anchor='w', padx=16, pady=16)
        for h in encontrados:
            Hallazgo(self.lista_hallazgos, h).pack(fill='x', padx=10, pady=(8, 0))
        tk.Frame(self.lista_hallazgos, bg=C['tarjeta'], height=10).pack()
        self.lienzo.yview_moveto(0)

        self.tabla.delete(*self.tabla.get_children())
        for i, g in enumerate(invitados):
            self.tabla.insert('', 'end', tags=('par',) if i % 2 else (), values=(
                g['nombre'], g['mesa'] or '—', 1 + g['acompanantes'],
                g['telefono'] or '—', g['notas'] or ''))

        self.selector.texto(1, '  Los invitados (%d)  ' % len(invitados))
        self.selector.seleccionar(0)      # lo primero es lo que hay que revisar
        self._cambiar_vista(0)
        self.btn_guardar.configure(state='normal')

    # ---------------------------------------------------------------- guardar

    def _alternar_hojas(self, _=None):
        self.hojas_visibles = not self.hojas_visibles
        if self.hojas_visibles:
            self.marco_hojas.grid(row=1, column=0, columnspan=3, sticky='ew',
                                  padx=8, pady=(0, 10))
            self.lbl_cambiar.configure(text='Ocultar hojas')
        else:
            self.marco_hojas.grid_forget()
            self.lbl_cambiar.configure(text='Cambiar hojas')

    def guardar(self):
        if not self.invitados:
            return
        # Con una lista pegada no hay archivo de origen: se guarda en el Escritorio.
        por_defecto = (os.path.splitext(os.path.basename(self.excel))[0]
                       if self.excel else 'LISTA DE INVITADOS')
        carpeta = (os.path.dirname(self.excel) if self.excel
                   else os.path.join(os.path.expanduser('~'), 'Desktop'))
        if not os.path.isdir(carpeta):
            carpeta = os.path.expanduser('~')
        ruta = filedialog.asksaveasfilename(
            title='Guardar el archivo para Guest List',
            initialfile=(self.var_evento.get().strip() or por_defecto) + '.glf',
            defaultextension='.glf', initialdir=carpeta,
            filetypes=[('Archivo de Guest List', '*.glf')])
        if not ruta:
            return
        try:
            glf.escribir_glf(ruta, self.invitados,
                             fecha=self.var_fecha.get().strip() or None,
                             hora=self.var_hora.get().strip() or '18:00')
        except Exception as e:
            messagebox.showerror('No se pudo guardar', str(e))
            return

        self.lbl_estado.configure(text='Guardado: %s' % os.path.basename(ruta),
                                  fg=C['exito'])
        if messagebox.askyesno(
                'Listo',
                'Se guardo "%s" con %d invitados.\n\n¿Abrir la carpeta para enviarlo?'
                % (os.path.basename(ruta), len(self.invitados))):
            abrir_carpeta(ruta)


def main():
    tema.aplicar(tema.leer_preferencias().get('tema', 'oscuro'))

    raiz = tk.Tk()
    raiz.title(TITULO)
    raiz.minsize(840, 660)
    try:
        # La referencia se guarda en la ventana o el recolector se lleva la imagen.
        raiz._icono = tk.PhotoImage(data=icono.PNG)
        raiz.iconphoto(True, raiz._icono)
    except tk.TclError:
        pass
    raiz.configure(bg=C['fondo'])
    tema.recordar_ventana(raiz, '960x840')
    tema.barra_de_titulo_oscura(raiz, C['nombre'] == 'oscuro')
    app = Aplicacion(raiz)

    # Arrastrar el Excel encima del programa lo abre con el archivo ya cargado.
    if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]):
        raiz.after(120, lambda: app.cargar_excel(sys.argv[1]))

    raiz.mainloop()


if __name__ == '__main__':
    main()
