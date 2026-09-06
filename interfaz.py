# -*- coding: utf-8 -*-
"""
Piezas visuales de esquinas redondeadas para la ventana.

Tk no sabe dibujar esquinas redondeadas: todos sus widgets son rectangulos
opacos. Aqui se dibujan sobre un Canvas y el contenido se mete encima, metido
hacia adentro lo suficiente para que sus esquinas cuadradas queden por dentro
de la curva y no se asomen.

Todo lee la paleta `C` del modulo `tema` en el momento de construirse; al
cambiar de tema la ventana se vuelve a armar.
"""

import tkinter as tk
from tkinter import font as tkfont

from tema import C


def contorno_redondeado(x1, y1, x2, y2, r):
    """Puntos de un rectangulo redondeado, para dibujarlo como poligono suavizado."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
            x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]


class Tarjeta(tk.Frame):
    """Recuadro de esquinas redondeadas. El contenido va dentro de `.cuerpo`."""

    def __init__(self, padre, radio=14, relleno=None, borde=None, fondo=None,
                 inset=8, grosor=1):
        self.fondo = fondo or C['fondo']
        super().__init__(padre, bg=self.fondo)
        self.relleno = relleno or C['tarjeta']
        self.borde = borde if borde is not None else C['borde']
        self.radio = radio
        self.grosor = grosor

        self.lienzo = tk.Canvas(self, bg=self.fondo, highlightthickness=0, bd=0)
        self.lienzo.place(x=0, y=0, relwidth=1, relheight=1)

        # El cuerpo va con grid para que la tarjeta adopte el alto del contenido;
        # el lienzo va con place para no influir en ese tamano.
        self.cuerpo = tk.Frame(self, bg=self.relleno)
        self.cuerpo.grid(row=0, column=0, sticky='nsew', padx=inset, pady=inset)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.bind('<Configure>', self._redibujar)

    def _redibujar(self, _=None):
        self.lienzo.delete('fondo')
        ancho, alto = self.winfo_width(), self.winfo_height()
        if ancho < 4 or alto < 4:
            return
        d = self.grosor
        self.lienzo.create_polygon(
            contorno_redondeado(d, d, ancho - d - 1, alto - d - 1, self.radio),
            smooth=True, splinesteps=32, fill=self.relleno,
            outline=self.borde or self.relleno, width=self.grosor, tags='fondo')
        self.lienzo.tag_lower('fondo')


class Boton(tk.Canvas):
    """Boton redondeado dibujado a mano, con estados de reposo, hover y apagado."""

    def __init__(self, padre, texto, orden, tipo='principal', fondo=None,
                 radio=11, ancho_extra=0):
        self.fondo = fondo or C['fondo']
        self.tipo = tipo
        self.orden = orden
        self.texto = texto
        self.radio = radio
        self.encima = False
        self._estado = 'normal'

        self.letra = tkfont.Font(family='Segoe UI Semibold', size=10)
        ancho = self.letra.measure(texto) + 46 + ancho_extra
        alto = self.letra.metrics('linespace') + 20
        super().__init__(padre, width=ancho, height=alto, bg=self.fondo,
                         highlightthickness=0, bd=0, cursor='hand2')

        self.bind('<Button-1>', self._pulsar)
        self.bind('<Enter>', self._entrar)
        self.bind('<Leave>', self._salir)
        self.bind('<Configure>', lambda e: self._pintar())
        self._pintar()

    # --- colores segun tipo y estado
    def _colores(self):
        if self._estado == 'disabled':
            return C['apagado'], C['apagado_texto'], C['apagado']
        if self.tipo == 'principal':
            relleno = C['azul_hover'] if self.encima else C['azul']
            return relleno, 'white', relleno
        relleno = C['gris'] if self.encima else C['tarjeta']
        return relleno, C['texto'], C['borde']

    def _pintar(self):
        self.delete('all')
        ancho, alto = self.winfo_width(), self.winfo_height()
        if ancho < 4:
            ancho, alto = int(self['width']), int(self['height'])
        relleno, tinta, borde = self._colores()
        self.create_polygon(contorno_redondeado(1, 1, ancho - 2, alto - 2, self.radio),
                            smooth=True, splinesteps=32, fill=relleno,
                            outline=borde, width=1)
        self.create_text(ancho / 2, alto / 2, text=self.texto, fill=tinta,
                         font=self.letra)

    # --- interaccion
    def _pulsar(self, _=None):
        if self._estado != 'disabled' and self.orden:
            self.orden()

    def _entrar(self, _=None):
        self.encima = True
        self._pintar()

    def _salir(self, _=None):
        self.encima = False
        self._pintar()

    def configure(self, **kw):
        """Acepta state='normal'/'disabled' como cualquier boton de Tk."""
        if 'state' in kw:
            self._estado = kw.pop('state')
            self.configure(cursor='hand2' if self._estado == 'normal' else 'arrow')
            self._pintar()
        if 'text' in kw:
            self.texto = kw.pop('text')
            self._pintar()
        if kw:
            super().configure(**kw)
        return None

    config = configure


class Selector(tk.Canvas):
    """Selector de dos o mas vistas, dibujado como pastillas redondeadas.

    Reemplaza al Notebook de ttk, que no se puede redondear y ademas cambiaba
    de pestana solo al redimensionarse la ventana.
    """

    def __init__(self, padre, opciones, al_cambiar, fondo=None, radio=10):
        self.fondo = fondo or C['fondo']
        self.opciones = list(opciones)
        self.al_cambiar = al_cambiar
        self.radio = radio
        self.actual = 0
        self.letra = tkfont.Font(family='Segoe UI', size=10)
        self._zonas = []
        super().__init__(padre, bg=self.fondo, highlightthickness=0, bd=0,
                         height=self.letra.metrics('linespace') + 20)
        self.bind('<Button-1>', self._clic)
        self.bind('<Motion>', self._mover)
        self._medir()
        self._pintar()

    def _medir(self):
        self.configure(width=sum(self.letra.measure(t) + 36 for t in self.opciones) + 8)

    def _pintar(self):
        self.delete('all')
        alto = int(self['height'])
        ancho = int(self['width'])
        self.create_polygon(contorno_redondeado(0, 0, ancho - 1, alto - 1, self.radio + 3),
                            smooth=True, splinesteps=32, fill=C['gris'], outline=C['gris'])
        self._zonas = []
        x = 4
        for i, texto in enumerate(self.opciones):
            w = self.letra.measure(texto) + 36
            if i == self.actual:
                self.create_polygon(
                    contorno_redondeado(x, 4, x + w, alto - 5, self.radio),
                    smooth=True, splinesteps=32, fill=C['tarjeta'], outline=C['borde'])
            self.create_text(x + w / 2, alto / 2, text=texto, font=self.letra,
                             fill=C['texto'] if i == self.actual else C['suave'])
            self._zonas.append((x, x + w))
            x += w

    def _indice_en(self, x):
        for i, (a, b) in enumerate(self._zonas):
            if a <= x <= b:
                return i
        return None

    def _clic(self, evento):
        i = self._indice_en(evento.x)
        if i is not None and i != self.actual:
            self.seleccionar(i)
            self.al_cambiar(i)

    def _mover(self, evento):
        self.configure(cursor='hand2' if self._indice_en(evento.x) is not None else 'arrow')

    def seleccionar(self, i):
        self.actual = i
        self._pintar()

    def texto(self, i, valor):
        self.opciones[i] = valor
        self._medir()
        self._pintar()


class Campo(Tarjeta):
    """Caja de texto dentro de una tarjeta redondeada."""

    def __init__(self, padre, variable, ancho=None, tamano=12, fondo=None):
        super().__init__(padre, radio=9, relleno=C['campo'], borde=C['borde'],
                         fondo=fondo, inset=4)
        self.entrada = tk.Entry(self.cuerpo, textvariable=variable, relief='flat', bd=0,
                                bg=C['campo'], fg=C['texto'], insertbackground=C['texto'],
                                font=('Segoe UI', tamano), highlightthickness=0,
                                width=ancho or 20)
        self.entrada.pack(fill='both', expand=True, padx=8, pady=5)
