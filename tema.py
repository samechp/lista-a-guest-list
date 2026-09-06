# -*- coding: utf-8 -*-
"""
Paletas de color de la ventana y la preferencia guardada del usuario.

`C` es la paleta puesta en este momento. Los widgets la leen cuando se
construyen, asi que al cambiar de tema la ventana se vuelve a armar entera.
"""

import os
import sys
import json
import ctypes

TEMAS = {
    'oscuro': {
        'nombre': 'oscuro',
        'fondo': '#15171C', 'tarjeta': '#1E2128', 'borde': '#2F343E',
        'texto': '#E8EAEF', 'suave': '#98A1B0', 'campo': '#272B33',
        'azul': '#3B7BF0', 'azul_hover': '#5590F5', 'gris': '#232730',
        'apagado': '#2A2E37', 'apagado_texto': '#5D6472', 'exito': '#7FD9A6',
        'seleccion': '#2B3A52',
        'niveles': {
            'ojo':  {'fondo': '#2B2416', 'borde': '#5C4A20', 'texto': '#F3C87C', 'icono': '!'},
            'info': {'fondo': '#232730', 'borde': '#333945', 'texto': '#AFB7C4', 'icono': 'i'},
            'ok':   {'fondo': '#15271E', 'borde': '#2C5A43', 'texto': '#7FD9A6', 'icono': '✓'},
        },
    },
    'claro': {
        'nombre': 'claro',
        'fondo': '#F2F4F7', 'tarjeta': '#FFFFFF', 'borde': '#E4E7EC',
        'texto': '#101828', 'suave': '#667085', 'campo': '#FFFFFF',
        'azul': '#2E6BE6', 'azul_hover': '#2559C4', 'gris': '#F1F3F7',
        'apagado': '#E4E7EC', 'apagado_texto': '#98A1B0', 'exito': '#067647',
        'seleccion': '#E8F0FE',
        'niveles': {
            'ojo':  {'fondo': '#FFFAEB', 'borde': '#FEDF89', 'texto': '#B54708', 'icono': '!'},
            'info': {'fondo': '#F9FAFB', 'borde': '#E4E7EC', 'texto': '#475467', 'icono': 'i'},
            'ok':   {'fondo': '#ECFDF3', 'borde': '#ABEFC6', 'texto': '#067647', 'icono': '✓'},
        },
    },
}

C = dict(TEMAS['oscuro'])          # el modo oscuro es el que viene de fabrica

RUTA_PREFERENCIAS = os.path.join(
    os.environ.get('LOCALAPPDATA') or os.path.expanduser('~'),
    'ListaGuestList', 'preferencias.json')


def leer_preferencias():
    try:
        with open(RUTA_PREFERENCIAS, encoding='utf-8') as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def guardar_preferencias(datos):
    try:
        os.makedirs(os.path.dirname(RUTA_PREFERENCIAS), exist_ok=True)
        with open(RUTA_PREFERENCIAS, 'w', encoding='utf-8') as f:
            json.dump(datos, f)
    except OSError:
        pass          # que no se pueda recordar la preferencia no es motivo de error


def aplicar(nombre):
    C.clear()
    C.update(TEMAS.get(nombre, TEMAS['oscuro']))


def barra_de_titulo_oscura(raiz, oscura):
    """Pinta tambien la barra del sistema, para que no quede blanca sobre negro."""
    if sys.platform != 'win32':
        return
    try:
        raiz.update_idletasks()
        ventana = ctypes.windll.user32.GetParent(raiz.winfo_id())
        valor = ctypes.c_int(1 if oscura else 0)
        for atributo in (20, 19):        # 20 en Win11/Win10 nuevo, 19 en Win10 viejo
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                ventana, atributo, ctypes.byref(valor), ctypes.sizeof(valor))
        # Windows solo repinta la barra si la ventana se oculta y se vuelve a
        # mostrar, y eso la saca de pantalla completa: hay que devolverla.
        estaba = raiz.state()
        raiz.withdraw()
        raiz.deiconify()
        if estaba == 'zoomed':
            raiz.state('zoomed')
    except Exception:
        pass


def recordar_ventana(raiz, geometria_por_defecto):
    """Abre la ventana como quedo la ultima vez y guarda el tamano al cerrarla."""
    guardadas = leer_preferencias()
    raiz.geometry(guardadas.get('geometria') or geometria_por_defecto)
    if guardadas.get('maximizada'):
        raiz.state('zoomed')

    # Solo interesa el tamano de la ventana suelta: el de la maximizada lo pone
    # Windows y no sirve para restaurarla despues.
    ultima = {'suelta': guardadas.get('geometria')}

    def anotar(evento):
        if evento.widget is raiz and raiz.state() == 'normal':
            ultima['suelta'] = raiz.geometry()

    def al_cerrar():
        datos = dict(leer_preferencias(), maximizada=raiz.state() == 'zoomed')
        if ultima['suelta']:
            datos['geometria'] = ultima['suelta']
        guardar_preferencias(datos)
        raiz.destroy()

    raiz.bind('<Configure>', anotar, add='+')
    raiz.protocol('WM_DELETE_WINDOW', al_cerrar)
    return al_cerrar
