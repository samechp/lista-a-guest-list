# -*- coding: utf-8 -*-
"""
Avisa si hay una version nueva publicada en GitHub.

Le pregunta a la API de GitHub cual es el ultimo Release y compara su etiqueta
con la version de este programa. Si el usuario dice que si, baja el .exe del
Release y se reemplaza a si mismo:

Windows no deja sobrescribir un .exe abierto, pero si cambiarle el nombre. Asi
que el viejo pasa a llamarse "<nombre>.viejo", el nuevo ocupa su lugar, se abre
el nuevo y este se cierra. El .viejo se borra la proxima vez que arranque.

Todo esto es opcional por definicion. El programa se usa muchas veces sin
internet, asi que cualquier fallo aqui se traga en silencio y la ventana sigue
funcionando igual.
"""

import json
import os
import re
import subprocess
import sys
import urllib.error
import urllib.request

VERSION = '1.4'

REPOSITORIO = 'samechp/lista-a-guest-list'
API_ULTIMO = 'https://api.github.com/repos/%s/releases/latest' % REPOSITORIO
PAGINA_DESCARGA = 'https://github.com/%s/releases/latest' % REPOSITORIO

ESPERA = 6          # segundos; si la red esta lenta no vale la pena insistir


def numeros(etiqueta):
    """'v1.10.2' -> (1, 10, 2). Lo que no sea numero se ignora."""
    return tuple(int(n) for n in re.findall(r'\d+', etiqueta or ''))


def es_mas_nueva(remota, propia=None):
    """Compara por numeros, no por texto: v1.10 es mayor que v1.9."""
    a, b = numeros(remota), numeros(propia or VERSION)
    if not a:
        return False
    largo = max(len(a), len(b))
    return a + (0,) * (largo - len(a)) > b + (0,) * (largo - len(b))


def buscar(version_propia=None):
    """Devuelve {'version', 'url', 'notas'} si hay una mas nueva, o None.

    None significa tanto "esta al dia" como "no se pudo averiguar": para el
    usuario es lo mismo, y no hay nada util que mostrarle en ninguno de los dos
    casos.
    """
    version_propia = version_propia or VERSION
    try:
        peticion = urllib.request.Request(
            API_ULTIMO,
            headers={'Accept': 'application/vnd.github+json',
                     'User-Agent': 'lista-a-guest-list/%s' % version_propia})
        with urllib.request.urlopen(peticion, timeout=ESPERA) as respuesta:
            datos = json.load(respuesta)
    except (urllib.error.URLError, OSError, ValueError, TimeoutError):
        return None

    etiqueta = datos.get('tag_name') or ''
    if not es_mas_nueva(etiqueta, version_propia):
        return None
    exes = [a for a in datos.get('assets') or []
            if (a.get('name') or '').lower().endswith('.exe')]
    return {'version': etiqueta.lstrip('vV'),
            'url': datos.get('html_url') or PAGINA_DESCARGA,
            'notas': (datos.get('body') or '').strip(),
            'descarga': exes[0].get('browser_download_url') if exes else None,
            'tamano': exes[0].get('size') if exes else None}


# ------------------------------------------------------- instalarse sola

def ejecutable_propio():
    """Ruta del .exe que esta corriendo, o None si es el codigo suelto."""
    if getattr(sys, 'frozen', False) and sys.platform == 'win32':
        return os.path.abspath(sys.executable)
    return None


def puede_instalarse(nueva):
    """Si se puede actualizar desde la ventana. Si no, queda el boton a la pagina."""
    exe = ejecutable_propio()
    return bool(exe and nueva and nueva.get('descarga')
                and os.access(os.path.dirname(exe), os.W_OK))


def limpiar_restos():
    """Borra el .exe viejo que dejo la ultima actualizacion, si lo hay."""
    exe = ejecutable_propio()
    for resto in (exe and (exe + '.viejo', exe + '.nuevo')) or ():
        try:
            os.remove(resto)
        except OSError:
            pass                      # no existe, o el viejo aun se esta cerrando


def descargar(nueva, avance=None):
    """Baja el .exe nuevo junto al actual, como "<nombre>.nuevo". Devuelve la ruta.

    `avance(fraccion)` se llama mientras baja. Si algo falla se lanza OSError y
    no queda nada a medias.
    """
    destino = ejecutable_propio() + '.nuevo'
    peticion = urllib.request.Request(
        nueva['descarga'],
        headers={'Accept': 'application/octet-stream',
                 'User-Agent': 'lista-a-guest-list/%s' % VERSION})
    try:
        with urllib.request.urlopen(peticion, timeout=30) as respuesta, \
                open(destino, 'wb') as salida:
            total = nueva.get('tamano') or int(respuesta.headers.get('Content-Length') or 0)
            bajado = 0
            while True:
                bloque = respuesta.read(64 * 1024)
                if not bloque:
                    break
                salida.write(bloque)
                bajado += len(bloque)
                if avance and total:
                    avance(min(1.0, bajado / total))
        with open(destino, 'rb') as f:
            cabeza = f.read(2)
        # Un corte de red deja un archivo mas corto; una pagina de error no empieza por MZ.
        if cabeza != b'MZ' or (total and os.path.getsize(destino) != total):
            raise OSError('La descarga llego incompleta.')
    except (urllib.error.URLError, OSError, ValueError, TimeoutError) as e:
        try:
            os.remove(destino)
        except OSError:
            pass
        raise OSError(str(e)) from e
    return destino


def instalar_y_reabrir(nuevo):
    """Pone el .exe nuevo en el lugar del actual y lo abre. Luego hay que cerrar."""
    exe = ejecutable_propio()
    viejo = exe + '.viejo'
    try:
        os.remove(viejo)
    except OSError:
        pass
    os.replace(exe, viejo)
    try:
        os.replace(nuevo, exe)
    except OSError:
        os.replace(viejo, exe)        # dejarlo todo como estaba
        raise
    # Sin esto el .exe nuevo cree que es parte de este y usa su carpeta temporal,
    # que desaparece en cuanto este se cierra.
    entorno = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT='1')
    subprocess.Popen([exe], env=entorno, close_fds=True,
                     creationflags=getattr(subprocess, 'DETACHED_PROCESS', 0))
